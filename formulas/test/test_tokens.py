#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#
# Copyright 2016-2019 European Commission (JRC);
# Licensed under the EUPL (the 'Licence');
# You may not use this work except in compliance with the Licence.
# You may obtain a copy of the Licence at: http://ec.europa.eu/idabc/eupl

import unittest
import ddt
from formulas.tokens.operand import String, Error, Range, Number
from formulas.tokens.operator import OperatorToken
from formulas.errors import TokenError


@ddt.ddt
class TestTokens(unittest.TestCase):
    @ddt.data(('A:A', 'A:A'), ('1:1', '1:1'), ('A1', 'A1'), ('A:A5', 'A:A5'),
              ('A1:B2', 'A1:B2'))
    def test_range(self, case):
        inputs, result = case
        token = Range(inputs)
        self.assertFalse(bool({'r1', 'r2', 'n1', 'n2'} - set(token.attr)))
        output = token.name
        self.assertEqual(result, output, '%s != %s' % (result, output))

    @ddt.data('1A', '.Sheet1!A1', '1Sheet1!A1', '.A1')
    def test_invalid_range(self, inputs):
        with self.assertRaises(TokenError):
            Range(inputs)

    @ddt.data(('1.2 a', 1.2), ('TrUe', True), ('FAlse', False),
              ('1e+10', 1e10), ('5', 5), ('3e-2', 3e-2))
    def test_number(self, case):
        inputs, result = case
        output = Number(inputs).compile()
        self.assertEqual(result, output, '%r != %r' % (result, output))

    @ddt.data('+1', '-1', '.4', '2e10', '3 :', '4.5a', 'TRUE1', 'FALSE3')
    def test_invalid_number(self, inputs):
        with self.assertRaises(TokenError):
            Number(inputs)

    @ddt.data(
        (' <=', '<='), (' <>', '<>'), ('  <', '<'), ('>', '>'), ('>=', '>='),
        ('=', '='), ('++', '+'), ('- - -', '-'), ('++- -+', '+'), (' *', '*'),
        ('^ ', '^'), (' & ', '&'), ('/', '/'), ('%', '%'), (' : ', ':'),
        ('*+', '*'))
    def test_valid_operators(self, case):
        inputs, result = case
        output = OperatorToken(inputs).name
        self.assertEqual(result, output, '%s != %s' % (result, output))

    @ddt.data(
        '=<', '> <', ' < <', '>>', '=>', '==', '+*', '**', '^ ^', 'z&', '\/',
        '', '%%', ', ,', ' : : ')
    def test_invalid_operators(self, inputs):
        with self.assertRaises(TokenError):
            OperatorToken(inputs)

    @ddt.data(
        ('""', ''), (' " "', ' '), ('  "a\'b"', 'a\'b'), ('" a " b"', ' a '),
        ('" "" "', ' "" '))
    def test_valid_strings(self, case):
        inputs, result = case
        output = String(inputs).name
        self.assertEqual(result, output, '%s != %s' % (result, output))

    @ddt.data('"', '" ', 'a\'b"', ' a " b')
    def test_invalid_strings(self, inputs):
        with self.assertRaises(TokenError):
            String(inputs)

    @ddt.data(
        ('#NULL! ', '#NULL!'), ('  #DIV/0!', '#DIV/0!'), ('#VALUE!', '#VALUE!'),
        ('#REF!', '#REF!'), ('#NAME?', '#NAME?'), ('#NUM!', '#NUM!'),
        ('#N/A  ', '#N/A'))
    def test_valid_errors(self, case):
        inputs, result = case
        output = Error(inputs).name
        self.assertEqual(result, output, '%s != %s' % (result, output))

    @ddt.data(
        '#NUL L!', ' #DIV\/0!', '# VALUE!', '#REF', '#NME?', '#NUM !', '#N\A ')
    def test_invalid_errors(self, inputs):
        with self.assertRaises(TokenError):
            Error(inputs)

    def test_deepcopy(self):
        import copy
        obj = Range('A:A')
        self.assertIsNot(obj, copy.deepcopy(obj))

    def test_dill(self):
        import dill
        obj = Range('A:A')
        self.assertEqual(str(obj), str(dill.loads(dill.dumps(obj))))
