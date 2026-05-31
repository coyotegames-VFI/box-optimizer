"""Small dependency-free QR PNG helper for workbook labels."""

from __future__ import annotations

import struct
import zlib


_QR_L_DATA_CODEWORDS = {1: 19, 2: 34, 3: 55, 4: 80}
_QR_L_EC_CODEWORDS = {1: 7, 2: 10, 3: 15, 4: 20}
_ALIGNMENT_POSITIONS = {1: [], 2: [6, 18], 3: [6, 22], 4: [6, 26]}


def qr_png(value: str, scale: int = 4, border: int = 4) -> bytes:
    """Return a PNG image containing a low-error-correction QR code."""
    return _matrix_png(_qr_matrix(value), scale=scale, border=border)


def _qr_matrix(value: str) -> list[list[bool]]:
    data = value.encode("utf-8")
    version = _version_for_length(len(data))
    size = 21 + (version - 1) * 4
    modules: list[list[bool | None]] = [[None for _ in range(size)] for _ in range(size)]
    function_modules = [[False for _ in range(size)] for _ in range(size)]

    _draw_function_patterns(modules, function_modules, version)
    _draw_codewords(modules, function_modules, _make_codewords(data, version))
    _apply_mask(modules, function_modules)
    _draw_format_bits(modules, function_modules, mask=0)

    return [[bool(value) for value in row] for row in modules]


def _version_for_length(length: int) -> int:
    for version, capacity in _QR_L_DATA_CODEWORDS.items():
        if length <= capacity - 2:
            return version
    raise ValueError("QR value is too long for built-in label QR generator")


def _draw_function_patterns(
    modules: list[list[bool | None]],
    function_modules: list[list[bool]],
    version: int,
) -> None:
    size = len(modules)
    _draw_finder(modules, function_modules, 3, 3)
    _draw_finder(modules, function_modules, size - 4, 3)
    _draw_finder(modules, function_modules, 3, size - 4)
    for index in range(8, size - 8):
        bit = index % 2 == 0
        _set_function(modules, function_modules, 6, index, bit)
        _set_function(modules, function_modules, index, 6, bit)
    for position in _ALIGNMENT_POSITIONS[version]:
        for other in _ALIGNMENT_POSITIONS[version]:
            if (position, other) in {(6, 6), (6, size - 7), (size - 7, 6)}:
                continue
            _draw_alignment(modules, function_modules, position, other)
    _set_function(modules, function_modules, 8, size - 8, True)
    _reserve_format(function_modules)


def _set_function(
    modules: list[list[bool | None]],
    function_modules: list[list[bool]],
    column: int,
    row: int,
    value: bool,
) -> None:
    modules[row][column] = value
    function_modules[row][column] = True


def _draw_finder(
    modules: list[list[bool | None]],
    function_modules: list[list[bool]],
    center_column: int,
    center_row: int,
) -> None:
    size = len(modules)
    for row in range(center_row - 4, center_row + 5):
        for column in range(center_column - 4, center_column + 5):
            if 0 <= row < size and 0 <= column < size:
                distance = max(abs(column - center_column), abs(row - center_row))
                _set_function(modules, function_modules, column, row, distance not in {2, 4})


def _draw_alignment(
    modules: list[list[bool | None]],
    function_modules: list[list[bool]],
    center_column: int,
    center_row: int,
) -> None:
    for row in range(center_row - 2, center_row + 3):
        for column in range(center_column - 2, center_column + 3):
            _set_function(
                modules,
                function_modules,
                column,
                row,
                max(abs(column - center_column), abs(row - center_row)) != 1,
            )


def _reserve_format(function_modules: list[list[bool]]) -> None:
    size = len(function_modules)
    for index in range(9):
        if index != 6:
            function_modules[8][index] = True
            function_modules[index][8] = True
    for index in range(8):
        function_modules[8][size - 1 - index] = True
        function_modules[size - 1 - index][8] = True


def _make_codewords(data: bytes, version: int) -> bytes:
    data_codeword_count = _QR_L_DATA_CODEWORDS[version]
    bits = [0, 1, 0, 0]
    bits.extend((len(data) >> shift) & 1 for shift in range(7, -1, -1))
    for byte in data:
        bits.extend((byte >> shift) & 1 for shift in range(7, -1, -1))
    bits.extend([0] * min(4, data_codeword_count * 8 - len(bits)))
    while len(bits) % 8:
        bits.append(0)
    codewords = [
        sum(bits[index + shift] << (7 - shift) for shift in range(8))
        for index in range(0, len(bits), 8)
    ]
    pad = 0xEC
    while len(codewords) < data_codeword_count:
        codewords.append(pad)
        pad ^= 0xEC ^ 0x11
    return bytes([*codewords, *_reed_solomon_remainder(codewords, _QR_L_EC_CODEWORDS[version])])


def _draw_codewords(
    modules: list[list[bool | None]],
    function_modules: list[list[bool]],
    codewords: bytes,
) -> None:
    bits = [(byte >> shift) & 1 for byte in codewords for shift in range(7, -1, -1)]
    size = len(modules)
    bit_index = 0
    upward = True
    column = size - 1
    while column > 0:
        if column == 6:
            column -= 1
        for offset in range(size):
            row = size - 1 - offset if upward else offset
            for current_column in [column, column - 1]:
                if not function_modules[row][current_column] and bit_index < len(bits):
                    modules[row][current_column] = bits[bit_index] == 1
                    bit_index += 1
        upward = not upward
        column -= 2


def _apply_mask(modules: list[list[bool | None]], function_modules: list[list[bool]]) -> None:
    for row in range(len(modules)):
        for column in range(len(modules)):
            if not function_modules[row][column] and (row + column) % 2 == 0:
                modules[row][column] = not bool(modules[row][column])


def _draw_format_bits(
    modules: list[list[bool | None]],
    function_modules: list[list[bool]],
    mask: int,
) -> None:
    size = len(modules)
    bits = _format_bits(mask)
    for index in range(6):
        _set_function(modules, function_modules, 8, index, _bit(bits, index))
    _set_function(modules, function_modules, 8, 7, _bit(bits, 6))
    _set_function(modules, function_modules, 8, 8, _bit(bits, 7))
    _set_function(modules, function_modules, 7, 8, _bit(bits, 8))
    for index in range(9, 15):
        _set_function(modules, function_modules, 14 - index, 8, _bit(bits, index))
    for index in range(8):
        _set_function(modules, function_modules, size - 1 - index, 8, _bit(bits, index))
    for index in range(8, 15):
        _set_function(modules, function_modules, 8, size - 15 + index, _bit(bits, index))
    _set_function(modules, function_modules, 8, size - 8, True)


def _format_bits(mask: int) -> int:
    data = (1 << 3) | mask
    value = data << 10
    generator = 0x537
    for shift in range(14, 9, -1):
        if (value >> shift) & 1:
            value ^= generator << (shift - 10)
    return ((data << 10) | value) ^ 0x5412


def _bit(value: int, index: int) -> bool:
    return ((value >> index) & 1) != 0


def _reed_solomon_remainder(data: list[int], degree: int) -> list[int]:
    generator = _reed_solomon_generator(degree)
    result = [*data, *([0] * degree)]
    for index, factor in enumerate(data):
        factor = result[index]
        if factor:
            for offset, coefficient in enumerate(generator):
                result[index + offset] ^= _gf_multiply(coefficient, factor)
    return result[-degree:]


def _reed_solomon_generator(degree: int) -> list[int]:
    result = [1]
    for index in range(degree):
        result = _poly_multiply(result, [1, _gf_power(index)])
    return result


def _poly_multiply(left: list[int], right: list[int]) -> list[int]:
    result = [0] * (len(left) + len(right) - 1)
    for left_index, left_value in enumerate(left):
        for right_index, right_value in enumerate(right):
            result[left_index + right_index] ^= _gf_multiply(left_value, right_value)
    return result


def _gf_power(power: int) -> int:
    value = 1
    for _ in range(power):
        value = _gf_multiply(value, 2)
    return value


def _gf_multiply(left: int, right: int) -> int:
    result = 0
    while right:
        if right & 1:
            result ^= left
        left <<= 1
        if left & 0x100:
            left ^= 0x11D
        right >>= 1
    return result


def _matrix_png(matrix: list[list[bool]], scale: int, border: int) -> bytes:
    module_count = len(matrix) + border * 2
    size = module_count * scale
    rows = []
    for y in range(size):
        module_y = y // scale - border
        row = bytearray([0])
        for x in range(size):
            module_x = x // scale - border
            dark = (
                0 <= module_y < len(matrix)
                and 0 <= module_x < len(matrix)
                and matrix[module_y][module_x]
            )
            row.append(0 if dark else 255)
        rows.append(bytes(row))
    raw = b"".join(rows)
    return b"".join(
        [
            b"\x89PNG\r\n\x1a\n",
            _png_chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 0, 0, 0, 0)),
            _png_chunk(b"IDAT", zlib.compress(raw, level=9)),
            _png_chunk(b"IEND", b""),
        ]
    )


def _png_chunk(kind: bytes, data: bytes) -> bytes:
    checksum = zlib.crc32(kind)
    checksum = zlib.crc32(data, checksum)
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", checksum & 0xFFFFFFFF)
