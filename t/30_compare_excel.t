#!/usr/bin/perl

use strict; use warnings;

use Test::More tests => 5;
use File::Spec::Functions;

BEGIN { use_ok('Test::Excel'); }

is(compare_excel(
    catfile('t', 'hello_world.xls'),
    catfile('t', 'hello_world.xls'),
), 1);


is(compare_excel(
    catfile('t', 'got-1.xls'),
    catfile('t', 'exp-1.xls'),
    { ignore => 'Ignore' }
), 1);

is(compare_excel(
    catfile('t', 'got-2.xls'),
    catfile('t', 'exp-2.xls'),
    { ignore => 'Ignore' }
), 0);

is(compare_excel(
    catfile('t', 'got-3.xls'),
    catfile('t', 'exp-3.xls'),
    { ignore => 'Ignore' }
), 0);

done_testing();