#!/usr/bin/env python3.5
# vim: set fileencoding=utf-8 fileformat=unix :

import datetime
from itertools import count
from collections import OrderedDict

import xlrd


__all__ = ("Reader", "DictReader")


NEWLINE = "\n"
NOTIME = datetime.time(0)
CONVERT = {
        #xlrd.XL_CELL_EMPTY: (is added at runtime)
        xlrd.XL_CELL_TEXT: str,
        xlrd.XL_CELL_NUMBER: float,
        #xlrd.XL_CELL_DATE: (is added at runtime)
        xlrd.XL_CELL_BOOLEAN: bool,
        xlrd.XL_CELL_ERROR: lambda n: xlrd.error_text_from_code.get(n),
        xlrd.XL_CELL_BLANK: str,
        }


class BaseReader(object):

    """Reader for a table on a Excel sheet"""

    def __init__(self, source, sheet,
            start_row=0, stop_row="", start_col=0, stop_col="",
            header_rows=1, empty="", repeat=False, trim=True):
        """Initiator.

        source          (xlrd.book.Book) Excel book
                        (str) pathname of Excel book (.xl*)
        sheet           (str) worksheet name; None=leftmost
        start_row       (int) top row number; starts with 0
        stop_row        (int) bottom row number [1]_ ; starts with 0
                        (float or str) boundary marker string to stop scan
                        (callable) see below [2]_
        start_col       (int) left column number; starts with 0
        stop_col        (int) right column number [3]_ ; starts with 0
                        (float or str) boundary marker string to stop scan
                        (callable) see below [4]_
        header_rows     (int) rows to read as field names; default=1
        empty           (float or str) alternative value for empty cells
        repeat          (bool) repeat cell value of the previous row if blank
        trim            (bool) suppress redundant zeros [5]_; default=True

        .. [1]  stop_row itself is not included in the data read.
        .. [2]  this callable (function) is to accept a list with the row
                values and to return True if the row is the boundary.
        .. [3]  stop_col itself is not included in the data read.
        .. [4]  this callable (function) is to accept a (cell) value and
                to return True if the column of the cell is the boundary.
        .. [5]  e.g. 3.0 -> 3, 2000-12-31 00:00:00 -> 2000-12-31
        """
        self.book = (xlrd.open_workbook(source) if isinstance(source, str)
                        else source)
        self.sheet = (self.book.sheet_by_name(sheet) if sheet else
                    self.book.sheet_by_index(0))
        self.start_row = start_row
        self.stop_row = stop_row
        self.start_col = start_col
        self.stop_col = stop_col
        self.header_rows = header_rows
        self.empty = empty if isinstance(empty, str) else float(empty)
        self.repeat = repeat
        self.trim = trim
        self.fieldnames = self._get_fields(header_rows)
        self._convert = CONVERT.copy()
        self._convert[xlrd.XL_CELL_EMPTY] = lambda x: self.empty
        self._convert[xlrd.XL_CELL_DATE] = lambda f: self._mkdt(f)

    def _mkdt(self, f):
        t = xlrd.xldate_as_tuple(f, self.book.datemode)
        # Trimming time part is done in _trim().
        #if f.is_integer(): return datetime.date(*t[:3])
        if f < 1: return datetime.time(*t[3:])
        return datetime.datetime(*t)

    def _mergearea(self, row, col):
        """Get the associated merge area.

        row         (int) row number of a cell; starts with 0
        col         (int) column number of a cell; starts with 0

        Returns a tuple of 4 integers as (rlo, rhi, clo, chi), where:

            rlo     (int) min. row number; starts with 0
            rhi     (int) max. row number + 1; starts with 0
            clo     (int) min. column number; starts with 0
            chi     (int) max. column number + 1; starts with 0
        """
        for (rlo, rhi, clo, chi) in self.sheet.merged_cells:
            if (rlo <= row < rhi and clo <= col < chi):
                return (rlo, rhi, clo, chi)
        return None

    @staticmethod
    def _isbreak_factory(criteria):
        if isinstance(criteria, int):
            def isbreak(row_or_col, _): return row_or_col == criteria
        elif isinstance(criteria, (float, str)):
            def isbreak(_, value): return value == criteria
        elif callable(criteria):
            def isbreak(_, value): return criteria(value)
        else:
            def isbreak(_, value): return not(bool(value))
        return isbreak

    def _get_fields(self, header_rows=1):
        """Get the field names.

        header_rows (int)   number of rows of the header; default=1

        Returns a list of field names.

        If 1 < header_rows, each field name consists of the values of
        vertically continuous cells joined by '_' (underscore).
        """
        fields = []
        r0 = self.start_row
        c0 = self.start_col
        rows = [self.sheet.row_values(r0 + r) for r in range(header_rows)]
        maxcol = c0 + max(len(rows[r]) for r in range(header_rows))
        isbreak = self._isbreak_factory(self.stop_col)
        for col in range(c0, maxcol + 1):
            f = []
            for row in range(header_rows):
                ma = self._mergearea(r0 + row, col)
                if not ma:
                    try:
                        v = rows[row][col]
                    except IndexError:
                        v = ""
                elif ma[0] - r0 == row: v = rows[row][ma[2]]
                else: v = None
                if v:
                    v = str(v)
                    if v.endswith(".0"):
                        v = v[:-2]
                    f.append(v)
            field = "_".join(f).replace(NEWLINE, "")
            fields.append(field or xlrd.colname(col))
            if isbreak(col, field): break
        fields.pop()
        # Add subscriptions for duplicate field names.
        for k, v in enumerate(fields):
            if v not in fields[:k]: continue
            for n in count(1):
                alt = "{}_{}".format(v, n)
                if alt not in fields[:k]:
                    fields[k] = alt
                    break
        return fields

    @staticmethod
    def _trim(values):
        for v in values:
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            elif isinstance(v, datetime.datetime) and v.time() == NOTIME:
                v = v.date()
            yield v

    def _build(self, kv):
        raise NotImplementedError

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __iter__(self):
        cols = slice(self.start_col, self.start_col + len(self.fieldnames))
        isbreak = self._isbreak_factory(self.stop_row)
        if self.repeat: prev = [None] * len(self.fieldnames)
        for row in count(self.start_row + self.header_rows):
            try:
                types = self.sheet.row_types(row)[cols]
                values = self.sheet.row_values(row)[cols]
            except IndexError:
                break
            values = [self._convert[t](v) for (t, v) in zip(types, values)]
            if len(values) < 1 or isbreak(row, values[0]): break
            if self.repeat:
                values = [p if v in (None, "") else v
                        for (v, p) in zip(values, prev)]
            if self.trim: values = self._trim(values)
            yield self._build(self.fieldnames, values)
            if self.repeat: prev = values


class _CSVRecord(object):
    pass


class Reader(BaseReader):

    def _build(self, keys, values):
        rec = _CSVRecord()
        for k, v in zip(keys, values):
            setattr(rec, k, v)
        return rec


class DictReader(BaseReader):

    def _build(self, keys, values):
        return OrderedDict(zip(keys, values))
