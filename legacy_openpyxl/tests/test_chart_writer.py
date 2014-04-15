import pytest
import os

from legacy_openpyxl.xml.functions import Element, fromstring, safe_iterator
from legacy_openpyxl.xml.constants import CHART_NS

from legacy_openpyxl.writer.charts import (ChartWriter,
                                    PieChartWriter,
                                    LineChartWriter,
                                    BarChartWriter,
                                    ScatterChartWriter,
                                    BaseChartWriter
                                    )
from legacy_openpyxl.styles import Color
