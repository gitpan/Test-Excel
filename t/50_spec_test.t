#!/usr/bin/perl

use strict; use warnings;

use Test::More tests => 1;

use Test::Excel;
use File::Spec::Functions;

is(compare_excel(
    catfile('t', 'got-4.xls'),
    catfile('t', 'exp-4.xls'),
    { sheet => 'Ignore', tolerance => 10**-12, sheet_tolerance => 0.20, spec => catfile('t', 'spec.txt') }
), 1);