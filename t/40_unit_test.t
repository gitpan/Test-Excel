#!/usr/bin/perl

use strict; use warnings;

use Test::More tests => 20;

use Test::Excel;
use File::Spec::Functions;

my ($got_col, $got_row);
my ($exp_col, $exp_row);
my ($cell, $range, $result);
my ($got_error, $exp_error);
my ($got_cells, $exp_cells);
my ($got_number, $exp_number);
my ($got_letter, $exp_letter);

$cell = 'A23';
$exp_col = 'A'; $exp_row = 23;
($got_col, $got_row) = Test::Excel::column_row($cell);
is($got_col, $exp_col);
is($got_row, $exp_row);

$range = 'A1:B3';
$exp_cells = [{row => 1, col => 0},
              {row => 1, col => 1},
              {row => 2, col => 0},
              {row => 2, col => 1},
              {row => 3, col => 0},
              {row => 3, col => 1}];
$got_cells = Test::Excel::cells_within_range($range);
ok(eq_array($got_cells, $exp_cells));

$exp_number = 27;
$got_number = Test::Excel::letter_to_number('AB');
is($got_number, $exp_number);

$exp_letter = 'AB';
$got_letter = Test::Excel::number_to_letter(27);
is($got_letter, $exp_letter);

eval
{
    $result = cmp_excel('x.xls','y.xls');
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Unable to locate got file.";
like($got_error, qr/$exp_error/);

eval
{
    $result = cmp_excel(catfile('t','got-1.xls'),'y.xls');
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Unable to locate expected file.";
like($got_error, qr/$exp_error/);

eval
{
    $result = compare_excel('x.xls','y.xls');
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Unable to locate got file.";
like($got_error, qr/$exp_error/);

eval
{
    $result = compare_excel(catfile('t','got-1.xls'),'y.xls');
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Unable to locate expected file.";
like($got_error, qr/$exp_error/);

eval
{
    $result = cmp_excel(catfile('t','got-1.xls'),
                        catfile('t','exp-1.xls'), 
                        'Test Message');
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Invalid RULE definitions. It has to be reference to a HASH.";
like($got_error, qr/$exp_error/);

eval
{
    $result = compare_excel(catfile('t','got-1.xls'),
                            catfile('t','exp-1.xls'), 
                            'Test Message');
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Invalid RULE definitions. It has to be reference to a HASH.";
like($got_error, qr/$exp_error/);

eval
{
    $result = cmp_excel(catfile('t','got-1.xls'),
                        catfile('t','exp-1.xls'), 
                        { name => 'Test Message'});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Invalid key found in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    $result = compare_excel(catfile('t','got-1.xls'),
                            catfile('t','exp-1.xls'), 
                            { name => 'Test Message'});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Invalid key found in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    $result = cmp_excel(catfile('t','got-1.xls'),
                        catfile('t','exp-1.xls'), 
                        { message         => 'Testing', 
                          sheet           => 'Test Message', 
                          sheet_tolerance => 0.2});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Missing key spec in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    $result = compare_excel(catfile('t','got-1.xls'),
                            catfile('t','exp-1.xls'), 
                            { message         => 'Testing', 
                              sheet           => 'Test Message', 
                              sheet_tolerance => 0.2});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Missing key spec in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    $result = cmp_excel(catfile('t','got-1.xls'),
                        catfile('t','exp-1.xls'), 
                        { message => 'Testing', 
                          spec    => catfile('t','spec-1.txt')});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Missing key sheet_tolerance in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    $result = compare_excel(catfile('t','got-1.xls'),
                            catfile('t','exp-1.xls'), 
                           { message => 'Testing', 
                             spec    => catfile('t','spec-1.txt')});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Missing key sheet_tolerance in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    $result = cmp_excel(catfile('t','got-1.xls'),
                        catfile('t','exp-1.xls'), 
                        { message         => 'Testing', 
                          sheet_tolerance => 0.2, 
                          spec            => catfile('t','spec-1.txt')});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Missing key tolerance in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    $result = compare_excel(catfile('t','got-1.xls'),
                            catfile('t','exp-1.xls'), 
                           { message         => 'Testing',
                             sheet_tolerance => 0.2,                            
                             spec            => catfile('t','spec-1.txt')});
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Missing key tolerance in the rule definitions.";
like($got_error, qr/$exp_error/);

eval
{
    Test::Excel::parse(catfile('t','spec-3.txt'));
};
$got_error = $@;
chomp($got_error);
$exp_error = "ERROR: Unable to locate spec file.";
like($got_error, qr/$exp_error/);