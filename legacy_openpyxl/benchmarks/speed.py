"Benchmark some different implementations for cells"

from legacy_openpyxl.cell import Cell
from legacy_openpyxl.reader.iter_worksheet import RawCell
from memory_profiler import memory_usage
import time


def standard():
    c = Cell(None, "A", "0", None)

def iterative():
    c = RawCell(None, None, (0, 0), None, 'n')

def dictionary():
    c = {'ws':'None', 'col':'A', 'row':0, 'value':1}


if __name__ == '__main__':
    initial_use = memory_usage(proc=-1, interval=1)[0]
    for fn in (standard, iterative, dictionary):
        t = time.time()
        container = []
        for i in xrange(1000000):
            container.append(fn())
        print "{0} {1} MB, {2:.2f}s".format(
            fn.func_name,
            memory_usage(proc=-1, interval=1)[0] - initial_use,
            time.time() - t)
