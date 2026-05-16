"""Report generation placeholders."""


def build_summary(rows: list[dict]) -> dict:
    """Build a minimal summary for report consumers."""
    return {"row_count": len(rows)}
