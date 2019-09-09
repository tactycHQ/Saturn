#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#
# Copyright 2016-2019 European Commission (JRC);
# Licensed under the EUPL (the 'Licence');
# You may not use this work except in compliance with the Licence.
# You may obtain a copy of the Licence at: http://ec.europa.eu/idabc/eupl
import os
import copy
import dill
import time
import unittest
import os.path as osp
import schedula as sh
from formulas.excel import ExcelModel, BOOK, ERR_CIRCULAR
from formulas.excel.xlreader import load_workbook
from formulas.functions import is_number

EXTRAS = os.environ.get('EXTRAS', 'all')

mydir = osp.join(osp.dirname(__file__), 'test_files')
_filename = 'test.xlsx'
_filename_compile = 'excel.xlsx'
_link_filename = 'test_link.xlsx'
_filename_circular = 'circular.xlsx'


def _book2dict(book):
    res = {}
    for ws in book.worksheets:
        s = res[ws.title.upper()] = {}
        for k, cell in ws._cells.items():
            value = getattr(cell, 'value', None)
            if value is not None:
                s[cell.coordinate] = value
    return res


def _res2books(res):
    return {k.upper(): _book2dict(v[BOOK]) for k, v in res.items()}


def _file2books(*fpaths):
    return {
        osp.basename(fp).upper(): _book2dict(load_workbook(fp, data_only=True))
        for fp in fpaths
    }


@unittest.skipIf(EXTRAS not in ('all', 'excel'), 'Not for extra %s.' % EXTRAS)
class TestExcelModel(unittest.TestCase):
    def setUp(self):
        self.filename = osp.join(mydir, _filename)
        self.link_filename = osp.join(mydir, _link_filename)
        self.filename_compile = osp.join(mydir, _filename_compile)
        self.filename_circular = osp.join(mydir, _filename_circular)

        self.results = _file2books(self.filename, self.link_filename)
        self.results_compile = _book2dict(
            load_workbook(self.filename_compile, data_only=True)
        )['DATA']
        self.results_circular = _file2books(self.filename_circular)
        self.maxDiff = None

    def _compare(self, books, results):
        it = sorted(sh.stack_nested_keys(results, depth=3))
        errors = []
        for k, res in it:
            value = sh.get_nested_dicts(books, *k, default=lambda: '')
            msg = '[{}]{}!{}'.format(*k)
            try:
                if is_number(value) and is_number(res):
                    self.assertAlmostEqual(float(res), float(value), msg=msg)
                else:
                    self.assertEqual(res, value, msg=msg)
            except AssertionError as ex:
                errors.append(str(ex))
        self.assertFalse(bool(errors), 'Errors:\n%s' % '\n'.join(errors))
        return len(it)

    def test_excel_model(self):
        start = time.time()
        _msg = '[info] test_excel_model: '
        xl_mdl = ExcelModel()

        print('\n%sLoading excel-model.' % _msg)
        s = time.time()

        xl_mdl.loads(self.filename)
        xl_mdl.add_book(self.link_filename)

        msg = '%sLoaded excel-model in %.2fs.\n%sFinishing excel-model.'
        print(msg % (_msg, time.time() - s, _msg))
        s = time.time()

        xl_mdl.finish()

        print('%sFinished excel-model in %.2fs.' % (_msg, time.time() - s))

        n_test = 0
        for i in range(3):
            print('%sCalculate excel-model.' % _msg)
            s = time.time()

            xl_mdl.calculate()

            msg = '%sCalculated excel-model in %.2fs.\n%s' \
                  'Comparing overwritten results.'
            print(msg % (_msg, time.time() - s, _msg))
            s = time.time()

            books = _res2books(xl_mdl.write(xl_mdl.books))
            n_test += self._compare(books, self.results)

            msg = '%sCompared overwritten results in %.2fs.\n' \
                  '%sComparing fresh written results.'
            print(msg % (_msg, time.time() - s, _msg))
            s = time.time()

            n_test += self._compare(_res2books(xl_mdl.write()), self.results)

            msg = '%sCompared fresh written results in %.2fs.'
            print(msg % (_msg, time.time() - s))

            if i == 0:
                print('%sSaving excel-model dill.' % _msg)
                s = time.time()

                xl_copy = dill.dumps(xl_mdl)

                msg = '%sSaved excel-model dill in %.2fs.\n' \
                      '%sLoading excel-model dill.'
                print(msg % (_msg, time.time() - s, _msg))
                s = time.time()

                xl_mdl = dill.loads(xl_copy)
                del xl_copy

                msg = '%sLoaded excel-model dill in %.2fs.'
                print(msg % (_msg, time.time() - s))

            elif i == 1:
                print('%sDeep-copying excel-model.' % _msg)
                s = time.time()

                xl_mdl = copy.deepcopy(xl_mdl)

                msg = '%sDeep-copied excel-model in %.2fs.'
                print(msg % (_msg, time.time() - s))

        print('%sSaving excel-model xlsx.' % _msg)
        s = time.time()

        dirpath = osp.join(mydir, 'tmp')
        xl_mdl.write(dirpath=dirpath)

        msg = '%sSaved excel-model exls in %.2fs.\n%sComparing saved results.'
        print(msg % (_msg, time.time() - s, _msg))
        s = time.time()

        n_test += self._compare(_file2books(*(
            osp.join(dirpath, fp) for fp in xl_mdl.books
        )), self.results)

        msg = '%sCompared saved results in %.2fs.\n%sRan %d tests in %.2fs'
        print(msg % (_msg, time.time() - s, _msg, n_test, time.time() - start))

    def test_excel_model_compile(self):
        xl_model = ExcelModel().loads(self.filename_compile).finish()

        inputs = ["A%d" % i for i in range(2, 5)]
        outputs = ["C%d" % i for i in range(2, 5)]
        func = xl_model.compile(
            ["'[EXCEL.XLSX]DATA'!%s" % i for i in inputs],
            ["'[EXCEL.XLSX]DATA'!%s" % i for i in outputs]
        )
        i = sh.selector(inputs, self.results_compile, output_type='list')
        res = sh.selector(outputs, self.results_compile, output_type='list')
        self.assertEqual([x.value[0, 0] for x in func(*i)], res)
        self.assertIsNot(xl_model, copy.deepcopy(xl_model))
        self.assertIsNot(func, copy.deepcopy(func))

        xl_model = ExcelModel().loads(self.filename_circular).finish(circular=1)
        func = xl_model.compile(
            ["'[CIRCULAR.XLSX]DATA'!A10"], ["'[CIRCULAR.XLSX]DATA'!E10"]
        )
        self.assertEqual(func(False).value[0, 0], 2.0)
        self.assertIs(func(True).value[0, 0], ERR_CIRCULAR)
        self.assertIsNot(xl_model, copy.deepcopy(xl_model))
        self.assertIsNot(func, copy.deepcopy(func))

    def test_excel_model_cycles(self):
        xl_model = ExcelModel().loads(self.filename_circular).finish(circular=1)
        xl_model.calculate()
        books = xl_model.books
        books = {k: _book2dict(v[BOOK])
                 for k, v in xl_model.write(books).items()}

        self._compare(books, self.results_circular)
