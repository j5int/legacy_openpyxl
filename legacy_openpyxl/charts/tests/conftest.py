import pytest

# Charts objects under test

@pytest.fixture
def Chart():
    """Chart class"""
    from legacy_openpyxl.charts.chart import Chart
    return Chart


@pytest.fixture
def GraphChart():
    """GraphicChart class"""
    from legacy_openpyxl.charts.chart import GraphChart
    return GraphChart


@pytest.fixture
def Axis():
    """Axis class"""
    from legacy_openpyxl.charts.axis import Axis
    return Axis


@pytest.fixture
def PieChart():
    """PieChart class"""
    from legacy_openpyxl.charts import PieChart
    return PieChart


@pytest.fixture
def LineChart():
    """LineChart class"""
    from legacy_openpyxl.charts import LineChart
    return LineChart


@pytest.fixture
def BarChart():
    """BarChart class"""
    from legacy_openpyxl.charts import BarChart
    return BarChart


@pytest.fixture
def ScatterChart():
    """ScatterChart class"""
    from legacy_openpyxl.charts import ScatterChart
    return ScatterChart


@pytest.fixture
def Reference():
    """Reference class"""
    from legacy_openpyxl.charts import Reference
    return Reference


@pytest.fixture
def Series():
    """Serie class"""
    from legacy_openpyxl.charts import Series
    return Series


@pytest.fixture
def ErrorBar():
    """ErrorBar class"""
    from legacy_openpyxl.charts import ErrorBar
    return ErrorBar


# Utility fixtures

@pytest.fixture
def ws(Workbook):
    """Empty worksheet titled 'data'"""
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def ten_row_sheet(ws):
    """Worksheet with values 0-9 in the first column"""
    for i in range(10):
        ws.append([i])
    return ws


@pytest.fixture
def sheet(ten_row_sheet):
    ten_row_sheet.title = "reference"
    return ten_row_sheet


@pytest.fixture
def cell(sheet, Reference):
    return Reference(sheet, (1, 1))


@pytest.fixture
def cell_range(sheet, Reference):
    return Reference(sheet, (1, 1), (10, 1))


@pytest.fixture()
def empty_range(sheet, Reference):
    for i in range(10):
        sheet.cell(row=i+1, column=2).value = None
    return Reference(sheet, (1, 2), (10, 2))


@pytest.fixture()
def missing_values(sheet, Reference):
    vals = [None, None, 1, 2, 3, 4, 5, 6, 7, 8]
    for idx, val in enumerate(vals, 1):
        sheet.cell(row=idx, column=3).value = val
    return Reference(sheet, (1, 3), (10, 3))
