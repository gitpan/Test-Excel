#!/usr/bin/perl

use strict; use warnings;

use Test::More tests => 5;

use Test::Excel;

my ($cell, $range);
my ($got_col, $got_row);
my ($exp_col, $exp_row);
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