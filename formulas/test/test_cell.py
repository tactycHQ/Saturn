#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#
# Copyright 2016-2019 European Commission (JRC);
# Licensed under the EUPL (the 'Licence');
# You may not use this work except in compliance with the Licence.
# You may obtain a copy of the Licence at: http://ec.europa.eu/idabc/eupl

import unittest
import ddt
import schedula as sh
from formulas.cell import Cell
from formulas.functions import Error
from formulas.functions.date import DEFAULT_DATE

DEFAULT_DATE[0] = 2019


def inp_ranges(*rng):
    return dict.fromkeys(rng, sh.EMPTY)


@ddt.ddt
class TestCell(unittest.TestCase):
    @ddt.idata((
        ('A1', '=LARGE({-1.1,10.1;"40",-2},1.1)', {}, '<Ranges>(A1)=[[-1.1]]'),
        ('A1', '=LARGE(A2:H2,"01/01/1900")', {
            'A2:H2':[[0.1, -10, 0.9, 2.2, -0.1, sh.EMPTY, "02/01/1900", True]]
        }, '<Ranges>(A1)=[[2.2]]'),

        ('A1:B1', '=SMALL(A2:B2,A3:B3)', {
            'A2:B2': [[4, sh.EMPTY]], 'A3:B3': [[1, Error.errors['#N/A']]]
        }, '<Ranges>(A1:B1)=[[4.0 #N/A]]'),
        ('A1', '=SMALL({-1.1,10.1;4.1,"40"},4)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=SMALL({-1.1,10.1;"40",TRUE},2)', {}, '<Ranges>(A1)=[[10.1]]'),
        ('A1', '=LARGE({-1.1,4.1,#REF!},"c")', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=LARGE({-1.1,4.1,#REF!},#N/A)', {}, '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=LARGE({-1.1,10.1;4.1,"40"},4)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=LARGE({-1.1,10.1;"40",-2},2)', {}, '<Ranges>(A1)=[[-1.1]]'),
        ('A1', '=LOOKUP(2,{-1.1,2.1,3.1,4.1},{#REF!,2.1,3.1,4.1})', {},
         '<Ranges>(A1)=[[#REF!]]'),
        ('A1', '=XIRR({-10000,2750,3250,4250,2},'
               '      {"39448",39508,39859,"30/10/2008", TRUE}, 0)',
         {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=XIRR({-10,1,5,0.0001,4.5},{1,20,4,4,5},"26/08/1987")',
         {}, '<Ranges>(A1)=[[38.321500577844446]]'),
        ('A1', '=XIRR({-10000,2750,3250,4250},'
               '      {"39448",39508,39859,"30/10/2008"}, 0)',
         {}, '<Ranges>(A1)=[[0.03379137764398378]]'),
        ('A1', '=XIRR({-10,1,1,2,3},{3,6,7,4,9})',
         {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=XIRR({-10,1,1,2,7},{3,6,7,4,9})', {},
         '<Ranges>(A1)=[[1937.5566679300437]]'),
        ('A1', '=XIRR({-10000,2750,3250,4250},{39448,39508,39859,39751}, 1)',
         {}, '<Ranges>(A1)=[[0.03379137764432629]]'),
        ('A1', '=IRR({-7,2,-1,4,-3;7,2,-1,4,4},A2:C2)',
         {'A2:C2': [[1, 20, 4]]}, '<Ranges>(A1)=[[0.19086464188385843]]'),
        ('A1', '=IRR({2,0,4},A2)',
         {'A2': [[sh.EMPTY]]}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=IRR({7,2,-1,4,-3;7,2,-1,4,4},A2)',
         {'A2': [[3]]}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=XNPV(0.02,{-10000,2750,3250,4250},{39448,39508,39859,39751})',
         {}, '<Ranges>(A1)=[[100.10102845727761]]'),
        ('A1', '=XNPV(0.02,{-10000,2750,3250,4250},{0,39508,39859,39751})',
         {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=NPV(A2:B2, A3:C3)',
         {'A2:B2': [[1, "ciao"]], 'A3:C3': [[5, 4, 3]]},
         '<Ranges>(A1)=[[3.875]]'),
        ('A1', '=NPV(D2, A2:C2)',
         {'D2': [["ciao"]], 'A2:C2': [[5, sh.EMPTY, Error.errors['#N/A']]]},
         '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=NPV(5, A2:D2)',
         {'A2:D2': [[5, sh.EMPTY, 'ciao', Error.errors['#N/A']]]},
         '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=NPV(-0.1, {-0.1,2,0,4,5})', {},
         '<Ranges>(A1)=[[16.922200206608068]]'),
        ('A1', '=COUNT(0,345,TRUE,#VALUE!,"26/08")', {}, '<Ranges>(A1)=[[4]]'),
        ('A1', '=MAX("")', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=MAXA("")', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=COUNTIF(A2:B2, "<>-1")', {'A2:B2': [[-1, 2]]},
         '<Ranges>(A1)=[[1]]'),
        ('A1', '=COUNTA(A2)', {'A2': [[sh.EMPTY]]}, '<Ranges>(A1)=[[0]]'),
        ('A1:C1', '=COUNTIF(A2:F2,{"60","29/02/1900","*0"})',
         {'A2:F2': [[60, "29/02/1900", sh.EMPTY, 0, "*", "AUG-98"]]},
         '<Ranges>(A1:C1)=[[2 2 1]]'),
        ('A1:G1', '=COUNTIF(A2:E2,{"<=FALSE",0,"",#VALUE!,"~*",FALSE})',
         {'A2:E2': [[sh.EMPTY, 0, Error.errors['#VALUE!'], "*", False]]},
         '<Ranges>(A1:G1)=[[1 1 1 1 1 1 #N/A]]'),
        ('A1', '=MAX("29/02/1900")', {}, '<Ranges>(A1)=[[60.0]]'),
        ('A1', '=COUNT(A2:E2,"26/08")',
         {'A2:E2': [[0, 345, True, Error.errors['#VALUE!'], "26/08"]]},
         '<Ranges>(A1)=[[3]]'),
        ('A1:G1', '=SUMIF(A2:E2,{"<=FALSE",0,"",#VALUE!,"~*",FALSE},A3:E3)',
         {'A2:E2': [[sh.EMPTY, 0, Error.errors['#VALUE!'], "*", False]],
          'A3:E3': [[11, 7, 5, 9, 2]]},
         '<Ranges>(A1:G1)=[[2 7 11 5 9 2 #N/A]]'),
        ('A1:C1', '=SUMIF(A2:F2,{"60","29/02/1900","*0"},A3:F3)',
         {'A2:F2': [[60, "29/02/1900", sh.EMPTY, 0, "*", "AUG-98"]],
          'A3:F3': [[1, 3, 11, 7, 9, 13]]},
         '<Ranges>(A1:C1)=[[4 4 3]]'),
        ('A1', '=YEARFRAC(0,345,TRUE)', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=YEARFRAC("26/8/1987 05:00 AM",345,4)', {},
         '<Ranges>(A1)=[[86.71111111111111]]'),
        ('A1', '=YEARFRAC("26/8/1987 05:00 AM",345,1)', {},
         '<Ranges>(A1)=[[86.71043215830248]]'),
        ('A1', '=YEARFRAC(2,1462,1)', {},
         '<Ranges>(A1)=[[3.9978094194961664]]'),
        ('A1', '=YEARFRAC(0,4382,1)', {}, '<Ranges>(A1)=[[12.0]]'),
        ('A1', '=YEARFRAC(1462,4382,1)', {},
         '<Ranges>(A1)=[[7.994524298425736]]'),
        ('A1', '=YEARFRAC(1462,4383,1)', {},
         '<Ranges>(A1)=[[7.997262149212868]]'),
        ('A1', '=YEARFRAC(3,368,1)', {}, '<Ranges>(A1)=[[1.0]]'),
        ('A1', '=YEARFRAC(366,2000,1)', {},
         '<Ranges>(A1)=[[4.4746691008671835]]'),
        ('A1', '=YEARFRAC(1,6000,1)', {}, '<Ranges>(A1)=[[16.4250281848929]]'),
        ('A1', '=YEARFRAC(1,2000,1)', {}, '<Ranges>(A1)=[[5.474212688270196]]'),
        ('A1', '=YEARFRAC(1,2000,0)', {}, '<Ranges>(A1)=[[5.475]]'),
        ('A1', '=YEARFRAC(1,2000,2)', {}, '<Ranges>(A1)=[[5.552777777777778]]'),
        ('A1', '=YEARFRAC(1,2000,3)', {}, '<Ranges>(A1)=[[5.476712328767123]]'),
        ('A1', '=YEARFRAC(1,2000,4)', {}, '<Ranges>(A1)=[[5.475]]'),
        ('A1', '=HOUR(-1)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=HOUR(0.4006770833)', {}, '<Ranges>(A1)=[[9]]'),
        ('A1', '=MINUTE(2.4)', {}, '<Ranges>(A1)=[[36]]'),
        ('A1', '=MINUTE(0.4006770833)', {}, '<Ranges>(A1)=[[36]]'),
        ('A1', '=SECOND(0.4006770833)', {}, '<Ranges>(A1)=[[58]]'),
        ('A1', '=SECOND(0.4006770834)', {}, '<Ranges>(A1)=[[59]]'),
        ('A1', '=SECOND(0.4)', {}, '<Ranges>(A1)=[[0]]'),
        ('A1', '=SECOND("22-Aug-2011 9:36 AM")', {}, '<Ranges>(A1)=[[0]]'),
        ('A1', '=TIMEVALUE("22-Aug-2011 9:36 AM")', {}, '<Ranges>(A1)=[[0.4]]'),
        ('A1', '=TIMEVALUE("9:36 AM")', {}, '<Ranges>(A1)=[[0.4]]'),
        ('A1', '=DAY(0)', {}, '<Ranges>(A1)=[[0]]'),
        ('A1', '=MONTH(0.7)', {}, '<Ranges>(A1)=[[1]]'),
        ('A1', '=TIME(24,0,0)', {}, '<Ranges>(A1)=[[0.0]]'),
        ('A1', '=TIME(36,12*60,6*60*60)', {}, '<Ranges>(A1)=[[0.25]]'),
        ('A1', '=TIME(36,0,6*60*60)', {}, '<Ranges>(A1)=[[0.75]]'),
        ('A1', '=TIME(12,0,6*60*60)', {}, '<Ranges>(A1)=[[0.75]]'),
        ('A1', '=TIME(0,0,6*60*60)', {}, '<Ranges>(A1)=[[0.25]]'),
        ('A1', '=DAY(2958466)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=DAY(60)', {}, '<Ranges>(A1)=[[29]]'),
        ('A1', '=DAY(50)', {}, '<Ranges>(A1)=[[19]]'),
        ('A1', '=DAY(100)', {}, '<Ranges>(A1)=[[9]]'),
        ('A1', '=DAY("29/2/1900")', {}, '<Ranges>(A1)=[[29]]'),
        ('A1', '=DAY("22 August 20")', {}, '<Ranges>(A1)=[[22]]'),
        ('A1', '=DATEVALUE("22 August 20")', {}, '<Ranges>(A1)=[[44065]]'),
        ('A1', '=DATEVALUE("01/01/00")', {}, '<Ranges>(A1)=[[36526]]'),
        ('A1', '=DATEVALUE("01/01/99")', {}, '<Ranges>(A1)=[[36161]]'),
        ('A1', '=DATEVALUE("01/01/29")', {}, '<Ranges>(A1)=[[47119]]'),
        ('A1', '=DATEVALUE("01/01/30")', {}, '<Ranges>(A1)=[[10959]]'),
        ('A1', '=DATEVALUE("8/22/2011")', {}, '<Ranges>(A1)=[[40777]]'),
        ('A1', '=DATEVALUE("22-MAY-2011")', {}, '<Ranges>(A1)=[[40685]]'),
        ('A1', '=DATEVALUE("2011/02/23")', {}, '<Ranges>(A1)=[[40597]]'),
        ('A1', '=DATEVALUE("5-JUL")', {}, '<Ranges>(A1)=[[43651]]'),
        ('A1', '=DATEVALUE("8/1987")', {}, '<Ranges>(A1)=[[31990]]'),
        ('A1', '=DATEVALUE("26/8/1987")', {}, '<Ranges>(A1)=[[32015]]'),
        ('A1', '=DATEVALUE("29/2/1900")', {}, '<Ranges>(A1)=[[60]]'),
        ('A1', '=DATE(1900,3,"")', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=DATE(1900,3,"-1")', {}, '<Ranges>(A1)=[[59]]'),
        ('A1', '=DATE(1900,3,-1)', {}, '<Ranges>(A1)=[[59]]'),
        ('A1', '=DATE(1900,3,0)', {}, '<Ranges>(A1)=[[60]]'),
        ('A1', '=DATE(9999,12,32)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=DATE(9999,12,31)', {}, '<Ranges>(A1)=[[2958465]]'),
        ('A1', '=DATE(-1,2,29)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=DATE(1899,2,29)', {}, '<Ranges>(A1)=[[693657]]'),
        ('A1', '=DATE(1900,1,0)', {}, '<Ranges>(A1)=[[0]]'),
        ('A1', '=DATE(1900,0,0)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=DATE(1900,2,28)', {}, '<Ranges>(A1)=[[59]]'),
        ('A1', '=DATE(1900,2,29)', {}, '<Ranges>(A1)=[[60]]'),
        ('A1', '=DATE(1904,2,29)', {}, '<Ranges>(A1)=[[1521]]'),
        ('A1', '=DATE(0,11.9,-40.1)', {}, '<Ranges>(A1)=[[264]]'),
        ('A1', '=DATE(0,11.4,1.9)', {}, '<Ranges>(A1)=[[306]]'),
        ('A1', '=DATE(0,11,0)', {}, '<Ranges>(A1)=[[305]]'),
        ('A1', '=DATE(1,-1.9,1.1)', {}, '<Ranges>(A1)=[[275]]'),
        ('A1', '=DATE(1,-1,0)', {}, '<Ranges>(A1)=[[305]]'),
        ('A1', '=DATE(0,0,0)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=DATE(2020,2,29)', {}, '<Ranges>(A1)=[[43890]]'),
        ('A1', '=OR({0,0,0},FALSE,"0")', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=OR(B1,FALSE)', {'B1': [['0']]}, '<Ranges>(A1)=[[False]]'),
        ('A1', '=OR("0",FALSE)', {'B1': [['0']]}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=XOR({0,0,0},FALSE,FALSE)', {}, '<Ranges>(A1)=[[False]]'),
        ('A1', '=XOR({0,0},FALSE,FALSE)', {}, '<Ranges>(A1)=[[False]]'),
        ('A1', '=XOR(TRUE,TRUE)', {}, '<Ranges>(A1)=[[False]]'),
        ('A1', '=OR(TRUE,#REF!,"0")', {}, '<Ranges>(A1)=[[#REF!]]'),
        ('A1', '=OR(FALSE,"0",#REF!)', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=INDEX({2,3;4,5},FALSE,"0")', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=INDEX({2,3;4,5}, -1)', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=INDEX(B1:C1, 1, 1)', {'B1:C1': [[sh.EMPTY, 2]]},
         '<Ranges>(A1)=[[0]]'),
        ('A1', '=INDEX(B1:C2, -1)', {'B1:C2': [[1, 2], [3, 4]]},
         '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=SUM(B1:D1 (B1:D1,B1:C1))',
         {'B1:D1': [[2, 3, 4]], 'B1:C1': [[2, 3]]}, '<Ranges>(A1)=[[14]]'),
        ('A1', '=INDEX(B1:D2 (B1:C1,B2:C2), 1, 1, 2)',
         {'B1:C1': [[2, 3]], 'B2:C2': [[4, 5]]}, '<Ranges>(A1)=[[4]]'),
        ('A1', '=INDEX((D1:D2:B1:C1, B2:C2), 1, 1, 2)',
         {'B1:D2': [[2, 3, 6], [4, 5, 7]], 'B2:C2': [[4, 5]]},
         '<Ranges>(A1)=[[4]]'),
        ('A1', '=INDEX({2,3;4,5},#NAME?)', {}, '<Ranges>(A1)=[[#NAME?]]'),
        ('A1:B2', '=INDEX({2,3;4,5},{1,2})', {},
         '<Ranges>(A1:B2)=[[2 4]\n [2 4]]'),
        ('A1', '=INDEX(C1:D2,1)', {'C1:D2': [[2, 3], [4, 5]]},
         '<Ranges>(A1)=[[#REF!]]'),
        ('A1:B1', '=INDEX(C1:D1,1)', {'C1:D1': [2, 3]},
         '<Ranges>(A1:B1)=[[2 2]]'),
        ('A1:B1', '=INDEX({2,3;4,5},1)', {}, '<Ranges>(A1:B1)=[[2 3]]'),
        ('A1', '=INDEX({2,3;4,5},1)', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=INDEX({2,3,4},1,1)', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=INDEX({2,3,4},2,1)', {}, '<Ranges>(A1)=[[#REF!]]'),
        ('A1', '=INDEX({2,3,4},2)', {}, '<Ranges>(A1)=[[3]]'),
        ('A1', '=LOOKUP(2,{-1.1,2.1,3.1,4.1})', {}, '<Ranges>(A1)=[[-1.1]]'),
        ('A1', '=LOOKUP(3,{-1.1,2.1,3.1,4.1})', {}, '<Ranges>(A1)=[[2.1]]'),
        ('A1', '=SWITCH(TRUE,1,0,,,TRUE,1,7)', {}, '<Ranges>(A1)=[[1]]'),
        ('A1:D1', '=SWITCH({0,1,TRUE},1,0,,,TRUE,1,7)', {},
         '<Ranges>(A1:D1)=[[0 0 1 #N/A]]'),
        ('A1', '=SWITCH(1,2,0,1,4,,4,5)', {}, '<Ranges>(A1)=[[4]]'),
        ('A1', '=GCD(5.2, -1, TRUE)', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=GCD(5.2, -1)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=GCD(5.2, 10)', {}, '<Ranges>(A1)=[[5]]'),
        ('A1', '=GCD(#NAME?, #VALUE!, #N/A)', {}, '<Ranges>(A1)=[[#NAME?]]'),
        ('A1', '=GCD(55, 15, 5)', {}, '<Ranges>(A1)=[[5]]'),
        ('A1', '=5%', {}, '<Ranges>(A1)=[[0.05]]'),
        ('A1', '=IF(#NAME?, #VALUE!, #N/A)', {}, '<Ranges>(A1)=[[#NAME?]]'),
        ('A1', '=IF(TRUE, #VALUE!, #N/A)', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=IF(FALSE, #VALUE!, #N/A)', {}, '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=IF(TRUE, "1a", "2b")', {}, '<Ranges>(A1)=[[\'1a\']]'),
        ('A1', '=ROW(4:7)', inp_ranges('4:7'), '<Ranges>(A1)=[[4]]'),
        ('A1', '=ROW(B8:D8:F7:H8 D7:E8)',
         inp_ranges('B8:D8', 'F7:H8', 'D7:E8'), '<Ranges>(A1)=[[7]]'),
        ('A1', '=COLUMN(B8:D8:F7:H8 D7:E7)',
         inp_ranges('B8:D8', 'F7:H8', 'D7:E7'), '<Ranges>(A1)=[[4]]'),
        ('A1:C3', '=ROW(D1:E1)', inp_ranges('D1:E1'),
         '<Ranges>(A1:C3)=[[1 1 1]\n [1 1 1]\n [1 1 1]]'),
        ('A1:C3', '=ROW(D1:D2)', inp_ranges('D1:D2'),
         '<Ranges>(A1:C3)=[[1 1 1]\n [2 2 2]\n [#N/A #N/A #N/A]]'),
        ('A1:C3', '=ROW(D1:E2)', inp_ranges('D1:E2'),
         '<Ranges>(A1:C3)=[[1 1 1]\n [2 2 2]\n [#N/A #N/A #N/A]]'),
        ('A11', '=ROW(B55:D55:F54:H55 D54:E54)',
         inp_ranges('B55:D55', 'F54:H55', 'D54:E54'), '<Ranges>(A11)=[[54]]'),
        ('A11', '=ROW(B53:D54 C54:E54)', inp_ranges('B53:D54', 'C54:E54'),
         '<Ranges>(A11)=[[54]]'),
        ('A11', '=ROW(L45)', inp_ranges('L45'), '<Ranges>(A11)=[[45]]'),
        ('A11', '=ROW()', {}, '<Ranges>(A11)=[[11]]'),
        ('A1', '=REF', {}, "<Ranges>(A1)=[[#REF!]]"),
        ('A1', '=(-INT(2))', {}, '<Ranges>(A1)=[[-2.0]]'),
        ('A1', '=(1+1)+(1+1)', {}, '<Ranges>(A1)=[[4.0]]'),
        ('A1', '=IFERROR(INDIRECT("aa") * 100,"")', {}, "<Ranges>(A1)=[['']]"),
        ('A1', '=( 1 + 2 + 3)*(4 + 5)^(1/5)', {},
         '<Ranges>(A1)=[[9.311073443492159]]'),
        ('A1', '={1,2;1,2}', {}, '<Ranges>(A1)=[[1]]'),
        ('A1:B2', '={1,2;1,2}', {}, '<Ranges>(A1:B2)=[[1 2]\n [1 2]]'),
        ('A1', '=PI()', {}, '<Ranges>(A1)=[[3.141592653589793]]'),
        ('A1', '=INT(1)%+3', {}, '<Ranges>(A1)=[[3.01]]'),
        ('A1', '=SUM({1, 3; 4, 2})', {}, '<Ranges>(A1)=[[10]]'),
        ('A1', '=" "" a"', {}, '<Ranges>(A1)=[[\' " a\']]'),
        ('A1', '=#NULL!', {}, "<Ranges>(A1)=[[#NULL!]]"),
        ('A1', '=1 + 2', {}, '<Ranges>(A1)=[[3.0]]'),
        ('A1', '=AVERAGE(((123 + 4 + AVERAGE({1,2}))))', {},
         '<Ranges>(A1)=[[128.5]]'),
        ('A1', '="a" & "b"""', {}, '<Ranges>(A1)=[[\'ab"\']]'),
        ('A1', '=SUM(B2:B4)', {'B2:B4': ('', '', '')}, '<Ranges>(A1)=[[0]]'),
        ('A1', '=SUM(B2:B4)', {'B2:B4': ('', 1, '')}, '<Ranges>(A1)=[[1]]'),
        ('A1', '=MATCH("*b?u*",{"a",2.1,"ds  bau  dsd",4.1},0)', {},
         '<Ranges>(A1)=[[3]]'),
        ('A1', '=MATCH(4.1,{FALSE,2.1,TRUE,4.1},-1)', {},
         '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=HLOOKUP(-1.1,{-1.1,2.1,3.1,4.1;5,6,7,8},2,0)', {},
         '<Ranges>(A1)=[[5]]'),
        ('A1', '=HLOOKUP(-1.1,{-1.1,2.1,3.1,4.1;5,6,7,8},3,0)', {},
         '<Ranges>(A1)=[[#REF!]]'),
        ('A1', '=MATCH(1.1,{"b",4.1,"a",1.1})', {}, '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=MATCH(1.1,{4.1,2.1,3.1,1.1})', {}, '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=MATCH(4.1,{4.1,"b","a",1.1})', {}, '<Ranges>(A1)=[[4]]'),
        ('A1', '=MATCH(4.1,{"b",4.1,"a",1.1})', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=MATCH(4.1,{4.1,"b","a",5.1},-1)', {}, '<Ranges>(A1)=[[1]]'),
        ('A1', '=MATCH(4.1,{"b",4.1,"a",5.1},-1)', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=MATCH("b",{"b",4.1,"a",1.1})', {}, '<Ranges>(A1)=[[3]]'),
        ('A1', '=MATCH(3,{-1.1,2.1,3.1,4.1})', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=MATCH(-1.1,{"b",4.1,"a",1.1})', {}, '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=MATCH(-1.1,{4.1,2.1,3.1,1.1},-1)', {}, '<Ranges>(A1)=[[4]]'),
        ('A1', '=MATCH(-1.1,{-1.1,2.1,3.1,4.1})', {}, '<Ranges>(A1)=[[1]]'),
        ('A1', '=MATCH(2.1,{4.1,2.1,3.1,1.1})', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=MATCH(2.1,{4.1,2.1,3.1,1.1},-1)', {}, '<Ranges>(A1)=[[2]]'),
        ('A1', '=MATCH(2,{4.1,2.1,3.1,1.1},-1)', {}, '<Ranges>(A1)=[[3]]'),
        ('A1', '=LOOKUP(2.1,{4.1,2.1,3.1,1.1},{"L","ML","MR","R"})', {},
         '<Ranges>(A1)=[[\'ML\']]'),
        ('A1', '=LOOKUP("b",{"b",4.1,"a",1.1},{"L","ML","MR","R"})', {},
         '<Ranges>(A1)=[[\'MR\']]'),
        ('A1', '=LOOKUP(TRUE,{TRUE,4.1,FALSE,1.1},{"L","ML","MR","R"})', {},
         '<Ranges>(A1)=[[\'MR\']]'),
        ('A1', '=LOOKUP(4.1,{"b",4.1,"a",1.1},{"L","ML","MR","R"})', {},
         '<Ranges>(A1)=[[\'ML\']]'),
        ('A1', '=LOOKUP(2,{"b",4.1,"a",1.1},{"L","ML","MR","R"})', {},
         '<Ranges>(A1)=[[#N/A]]'),
        ('A1', '=LOOKUP(4.1,{4.1,2.1,3.1,1.1},{"L","ML","MR","R"})', {},
         '<Ranges>(A1)=[[\'R\']]'),
        ('A1', '=LOOKUP(4,{4.1,2.1,3.1,1.1},{"L","ML","MR","R"})', {},
         '<Ranges>(A1)=[[\'R\']]'),
        ('A1:D1', '=IF({0,-0.2,0},2,{1})', {},
         '<Ranges>(A1:D1)=[[1 2 1 #N/A]]'),
        ('A1', '=HEX2DEC(9999999999)', {}, '<Ranges>(A1)=[[-439804651111]]'),
        ('A1', '=HEX2BIN(9999999999)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=HEX2BIN("FFFFFFFE00")', {}, '<Ranges>(A1)=[[\'1000000000\']]'),
        ('A1', '=HEX2BIN("1ff")', {}, '<Ranges>(A1)=[[\'111111111\']]'),
        ('A1', '=HEX2OCT("FF0000000")', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=HEX2OCT("FFE0000000")', {}, '<Ranges>(A1)=[[\'4000000000\']]'),
        ('A1', '=HEX2OCT("1FFFFFFF")', {}, '<Ranges>(A1)=[[\'3777777777\']]'),
        ('A1', '=DEC2HEX(-439804651111)', {},
         '<Ranges>(A1)=[[\'9999999999\']]'),
        ('A1', '=DEC2BIN(TRUE)', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=DEC2BIN(#DIV/0!)', {}, '<Ranges>(A1)=[[#DIV/0!]]'),
        ('A1', '=DEC2BIN("a")', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        ('A1', '=DEC2BIN(4,6)', {}, '<Ranges>(A1)=[[\'000100\']]'),
        ('A1', '=DEC2BIN(4,-2)', {}, '<Ranges>(A1)=[[#NUM!]]'),
        ('A1', '=DEC2BIN(4,"a")', {}, '<Ranges>(A1)=[[#VALUE!]]'),
        # ('A1:D1', '=IF({0,-0.2,0},{2,3},{1})', {},
        #  '<Ranges>(A1:D1)=[[1 2 1 #N/A]]'),
        # ('A1:D1', '=IF({0,-2,0},{2,3},{1,4})', {},
        #  '<Ranges>(A1:D1)=[[1 2 #N/A #N/A]]')
    ))
    def test_output(self, case):
        reference, formula, inputs, result = case
        dsp = sh.Dispatcher()
        cell = Cell(reference, formula).compile()
        assert cell.add(dsp)
        output = str(dsp(inputs)[cell.output])
        self.assertEqual(
            result, output,
            'Formula({}): {} != {}'.format(formula, result, output)
        )

    @ddt.idata((
        ('A1:D1', '=IF({0,-0.2,0},{2,3},{1})', {}),  # BroadcastError
        ('A1:D1', '=IF({0,-2,0},{2,3},{1,4})', {}),  # BroadcastError
    ))
    def test_invalid(self, case):
        reference, formula, inputs = case
        with self.assertRaises(sh.DispatcherError):
            dsp = sh.Dispatcher(raises=True)
            cell = Cell(reference, formula).compile()
            assert cell.add(dsp)
            dsp(inputs)