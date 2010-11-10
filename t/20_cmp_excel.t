#!/usr/bin/perl

use strict; use warnings;

use Test::More;
use File::Spec::Functions;

BEGIN { use_ok('Test::Excel'); }

cmp_excel(
    catfile('t', 'hello_world.xls'), 
    catfile('t', 'hello_world.xls'), 
    {message => 'Our Excels were essentially the same.'}
);

done_testing();