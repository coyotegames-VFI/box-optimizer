from pathlib import Path
from uuid import uuid4

import pytest


@pytest.fixture
def tmp_path() -> Path:
    """Workspace-local temp path for Windows sandboxes with locked system temp."""
    path = Path("test_outputs") / uuid4().hex
    path.mkdir(parents=True, exist_ok=True)
    return path.resolve()
