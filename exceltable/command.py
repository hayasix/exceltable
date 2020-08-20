#!/usr/bin/env python3
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


def main(sheetspec, start, stop,
         header_rows=1,
         empty_value=None,
         repeat=False,
         raw=False,
         file=None,
         output_header=True):
    book, sheet = (sheetspec.split("!", 1) + [None])[:2]
    table = reader.DictReader(book, sheet,
                start_row=start[0], stop_row=stop[0],
                start_col=start[1], stop_col=stop[1],
                header_rows=header_rows,
                empty=empty_value, repeat=repeat, trim=not raw)
    fieldnames = [f.replace(NBSP, " ") for f in table.fieldnames]
    if not file:
        if output_header: return fieldnames, table
        return table
    writer = csv.DictWriter(file, fieldnames, lineterminator="\n")
    if output_header: writer.writeheader()
    for row in table:
        for k, v in row.items():
            row[k] = str(v or "").replace(NBSP, " ")
        writer.writerow(row)


def __main__():
    args = docopt.docopt(__doc__.format(script="exceltable"),
                       version=__version__)
    sr = sc = er = ec = None
    if args["--start"]: sr, sc = decompose_address(args["--start"])
    if args["--stop"]: er, ec = decompose_address(args["--stop"])
    if args["--start-row"]: sr = args["--start-row"]
    if args["--stop-row"]: er = args["--stop-row"]
    if args["--start-col"]: sc = args["--start-col"]
    if args["--stop-col"]: ec = args["--stop-col"]
    sr = _inner_row(_eval(sr or 1))
    sc = _inner_col(_eval(sc or "A"))
    er = _inner_row(_eval(er or ""))
    ec = _inner_col(_eval(ec or ""))
    header_rows = _eval(args["--header-rows"])
    empty_value = _eval(args["--empty"])
    main(args["SHEETSPEC"], (sr, sc), (er, ec),
            header_rows=int(args["--header-rows"]),
            empty_value=empty_value,
            repeat=args["--repeat"],
            raw=args["--raw"],
            file=sys.stdout,
            output_header=args["--header"])


if __name__ == "__main__":
    sys.exit(__main__())
