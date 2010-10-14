#!/usr/bin/perl

use strict; use warnings;

use Test::More;
use File::Spec::Functions;

BEGIN { use_ok('Test::Excel'); }

is(compare_excel(
    catfile('t', 'hello_world.xls'), 
    catfile('t', 'hello_world.xls'), 
), 1);

done_testing();