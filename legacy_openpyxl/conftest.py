import pytest

# Global objects under tests

@pytest.fixture
def Workbook():
    """Workbook Class"""
    from legacy_openpyxl import Workbook
    return Workbook


@pytest.fixture
def Worksheet():
    """Worksheet Class"""
    from legacy_openpyxl.worksheet import Worksheet
    return Worksheet

