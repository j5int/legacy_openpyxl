# Fixtures (pre-configured objects) for tests

import pytest

@pytest.fixture
def NumberFormat():
    from legacy_openpyxl.styles import NumberFormat
    return NumberFormat
