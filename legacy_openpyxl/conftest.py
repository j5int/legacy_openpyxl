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


# Global fixtures

@pytest.fixture
def root_xml():
    """Root XML element <test>"""
    from legacy_openpyxl.xml.functions import Element
    return Element("test")

