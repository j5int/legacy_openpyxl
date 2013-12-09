# Fixtures (pre-configured objects) for tests

import pytest

# objects under test

@pytest.fixture
def NumberFormat():
    from legacy_openpyxl.styles import NumberFormat
    return NumberFormat


@pytest.fixture
def Workbook():
    from legacy_openpyxl import Workbook
    return Workbook

# Charts

@pytest.fixture
def Chart():
    from legacy_openpyxl.chart import Chart
    return Chart


@pytest.fixture
def GraphChart():
    from legacy_openpyxl.chart import GraphChart
    return GraphChart


@pytest.fixture
def Axis():
    from legacy_openpyxl.chart import Axis
    return Axis


@pytest.fixture
def PieChart():
    from legacy_openpyxl.chart import PieChart
    return PieChart


@pytest.fixture
def LineChart():
    from legacy_openpyxl.chart import LineChart
    return LineChart


@pytest.fixture
def BarChart():
    from legacy_openpyxl.chart import BarChart
    return BarChart


@pytest.fixture
def ScatterChart():
    from legacy_openpyxl.chart import ScatterChart
    return ScatterChart


@pytest.fixture
def Reference():
    from legacy_openpyxl.chart import Reference
    return Reference


@pytest.fixture
def Serie():
    from legacy_openpyxl.chart import Serie
    return Serie


@pytest.fixture
def ErrorBar():
    from legacy_openpyxl.chart import ErrorBar
    return ErrorBar

# utility fixtures

@pytest.fixture
def ws(Workbook):
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def ten_row_sheet(ws):
    for i in range(10):
        ws.cell(row=i, column=0).value = i
    return ws


@pytest.fixture
def ten_column_sheet(ws):
    ws.append(list(range(10)))
    return ws

from legacy_openpyxl.shared.xmltools import Element

@pytest.fixture
def root_xml():
    return Element("test")


### Markers ###

def pytest_runtest_setup(item):
    if isinstance(item, item.Function):
        try:
            from PIL import Image
        except ImportError:
            Image = False
        if item.get_marker("pil_required") and Image is False:
            pytest.skip("PIL must be installed")
        elif item.get_marker("pil_not_installed") and Image:
            pytest.skip("PIL is installed")
