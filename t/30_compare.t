﻿#!/usr/bin/perl

use strict; use warnings;

use Test::More tests => 5;
use File::Spec::Functions;

BEGIN { use_ok('Test::Excel'); }

is(compare(
    catfile('t', 'hello_world.xls'),
    catfile('t', 'hello_world.xls'),
), 1);


is(compare(
    catfile('t', 'got-1.xls'),
    catfile('t', 'exp-1.xls'),
    { sheet => 'MySheet1|MySheet2', tolerance => 10**-12, sheet_tolerance => 0.20 }
), 1);

is(compare(
    catfile('t', 'got-2.xls'),
    catfile('t', 'exp-2.xls'),
    { sheet => 'MySheet1|MySheet2', tolerance => 10**-12, sheet_tolerance => 0.20 }
), 0);

is(compare(
    catfile('t', 'got-3.xls'),
    catfile('t', 'exp-3.xls'),
    { sheet => 'MySheet1|MySheet2', tolerance => 10**-12, sheet_tolerance => 0.20 }
), 0);

done_testing();