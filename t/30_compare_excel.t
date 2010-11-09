#!/usr/bin/perl

use strict; use warnings;

use Test::More tests => 6;
use File::Spec::Functions;

BEGIN { use_ok('Test::Excel'); }

is(compare_excel(
    catfile('t', 'hello_world.xls'),
    catfile('t', 'hello_world.xls'),
), 1);


is(compare_excel(
    catfile('t', 'got-1.xls'),
    catfile('t', 'exp-1.xls'),
    { sheet => 'Ignore', tolerance => 10**-12, sheet_tolerance => 0.20 }
), 1);

is(compare_excel(
    catfile('t', 'got-2.xls'),
    catfile('t', 'exp-2.xls'),
    { sheet => 'Ignore', tolerance => 10**-12, sheet_tolerance => 0.20 }
), 0);

is(compare_excel(
    catfile('t', 'got-3.xls'),
    catfile('t', 'exp-3.xls'),
    { sheet => 'Ignore', tolerance => 10**-12, sheet_tolerance => 0.20 }
), 0);


eval
{
    compare_excel(
        catfile('t', 'got-1.xls'),
        catfile('t', 'exp-1.xls'),
        { sheet => 'Ignore' }
    );
};
my $got = $@;
chomp($got);
my $exp = "ERROR: Invalid RULE definition.";
like($got, qr/$exp/);

done_testing();