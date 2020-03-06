#!/usr/bin/env python3.5
# vim: set fileencoding=utf-8 fileformat=unix :

"""Usage: {script} [options] SHEETSPEC

Arguments:
  SHEETSPEC             path/to/workbook!sheet or path/to/workbook (the
                        leftmost sheet)

Options:
  -h, --help            show this help message and exit
  -s, --start ADDRESS   start row and column e.g. 'A1', 'R1C1' [default: A1]
  -S, --stop ADDRESS    stop row and column e.g. 'Z99, 'R99C26'
  -r, --start-row ROW   start row of table
  -R, --stop-row ROW    stop row of table or pattern of cell value
  -c, --start-col COL   start col of table
  -C, --stop-col COL    stop col of table or pattern of cell value
  --header              output with header
  --header-rows N       rows to read as field names [default: 1]
  --empty VALUE         default value for empty cells
  --repeat              repeat previous value if blank
  --raw                 don't suppress redundant zeros e.g. 1.0 -> 1.0, not 1
  --version             show version and exit

following notations are available for ROW/COL/DEFVAL:
    A, $A, BZ etc. for column address
    T:... or T(...) for text data
    N:... or N(...) for numeric data
    1, 2, ... for column/row number
    1.0, 1.1 etc. for numeric data (N.B. Excel stores only floats)
"""


import sys
import os
import argparse
import csv
from string import digits, ascii_uppercase, ascii_lowercase
import re

import docopt

from .__init__ import __version__
from . import reader


NBSP = "\xa0"
RADIX26 = "".maketrans(ascii_uppercase + ascii_lowercase,
        digits + ascii_uppercase[:16] + digits + ascii_lowercase[:16])
R1C1FORMAT = re.compile(r"[Rr](\d+)[Cc](\d+)$", re.IGNORECASE)
A1FORMAT = re.compile(r"([A-Z]+)(\d+)$", re.IGNORECASE)


class Args(dict):
    pass


def _inner_col(col):
    if col is None or col == "": return ""
    if isinstance(col, int): return col - 1
    if col.isdigit(): return int(col) - 1
    col = col.lstrip("$")
    try:
        return int(col.translate(RADIX26), 26) + (0, 1, 27, 703)[len(col)] - 1
    except ValueError:
        return col


def _inner_row(row):
    if row is None or row == "": return ""
    if isinstance(row, int): return row - 1
    if row.isdigit(): return int(row) - 1
    return row


def _eval(s):
    if s is None: return ""
    if not isinstance(s, str): return s
    if s.startswith("T:"): return s[2:]
    if s.startswith("T(") and s.endswith(")"): return s[2:-1]
    if s.startswith("N:"): return float(s[2:])
    if s.startswith("N(") and s.endswith(")"): return float(s[2:-1])
    if s.isdigit(): return int(s)
    if s.count(".") == 1 and s.replace(".", "").isdigit(): return float(s)
    return s


def decompose_address(s):
    mo = R1C1FORMAT.match(s)
    if mo: return (mo.group(1), mo.group(2))
    mo = A1FORMAT.match(s)
    if mo: return (mo.group(2), mo.group(1))
    raise IndexError("illegal cell address '{}'".format(s))


def main(args, file=None):
    args = Args(args)
    for k, v in args.items():
        setattr(args, k.lstrip("-").lower().replace("-", "_"), v)
    if hasattr(args, "start"):
        args.start_row, args.start_col = decompose_address(args.start)
    if hasattr(args, "stop"):
        args.stop_row, args.stop_col = decompose_address(args.stop)
    book, sheet = (args.sheetspec.split("!", 1) + [None])[:2]
    table = reader.DictReader(book, sheet,
                start_row=_inner_row(_eval(getattr(args, "start_row", 1))),
                stop_row=_inner_row(_eval(getattr(args, "stop_row", ""))),
                start_col=_inner_col(_eval(getattr(args, "start_col", "A"))),
                stop_col=_inner_col(_eval(getattr(args, "stop_col", ""))),
                header_rows=_eval(getattr(args, "header_rows", 1)),
                empty=_eval(getattr(args, "empty", "")),
                repeat=bool(getattr(args, "repeat", False)),
                trim=not(bool(getattr(args, "raw", False))))
    fieldnames = [f.replace(NBSP, " ") for f in table.fieldnames]
    if not file:
        if bool(getattr(args, "header", False)):
            return fieldnames, table
        return table
    writer = csv.DictWriter(file, fieldnames)
    if bool(getattr(args, "header", False)): writer.writeheader()
    for row in table:
        for k, v in row.items():
            row[k] = str(v or "").replace(NBSP, " ")
        writer.writerow(row)


def __main__():
    main(docopt.docopt(__doc__.format(script="exceltable"),
                       version=__version__),
         file=sys.stdout)


if __name__ == "__main__":
    sys.exit(__main__())
