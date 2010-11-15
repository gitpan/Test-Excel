#!/usr/bin/perl

use strict; use warnings;

use Test::More tests => 3;
use File::Spec::Functions;

BEGIN { use_ok('Test::Excel'); }

cmp_excel(
    catfile('t', 'hello_world.xls'), 
    catfile('t', 'hello_world.xls'), 
    {message => 'Our Excels were essentially the same.'}
);

cmp_excel(
    catfile('t', 'got-7.xls'),
    catfile('t', 'exp-7.xls'),
    { swap_check => 1, error_limit => 2, sheet => 'MySheet1|MySheet2', tolerance => 10**-12, sheet_tolerance => 0.20 }
);

done_testing();