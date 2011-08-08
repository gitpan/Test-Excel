package Test::Excel;

use strict; use warnings;

use Carp;
use IO::File;
use Readonly;
use Data::Dumper;
use Test::Builder ();
use Scalar::Util 'blessed';
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::Utility qw(int2col col2int);

require Exporter;

our @ISA    = qw(Exporter);
our @EXPORT = qw(cmp_excel compare_excel column_row letter_to_number number_to_letter cells_within_range);

=head1 NAME

Test::Excel - Interface to test and compare Excel files.

=head1 VERSION

Version 1.24

=head1 AWARD

Test::Excel has been granted the "Famous Software Award" by Download.FamousWhy.com on Wed 17 Nov 2010.

http://download.famouswhy.com/test_excel/

=cut

our $VERSION = '1.24';

$|=1;

our $DEBUG = 0;
Readonly my $ALMOST_ZERO          => 10**-16;
Readonly my $IGNORE               => 1;
Readonly my $SPECIAL_CASE         => 2;
Readonly my $MAX_ERRORS_PER_SHEET => 0;

=head1 DESCRIPTION

This  module  is  meant to be used for testing custom generated  Excel  files, it provides two
functions at  the  moment,  which  is C<cmp_excel>  and C<compare_excel>. These can be used to
compare_excel 2  Excel files to see if they are I<visually> similar. The function C<cmp_excel>
is for testing purpose where function C<compare_excel> can be used as standalone.

=head1 RULE

The  new paramter has been added to both method cmp_excel() and  method compare_excel() called
RULE.  This is optional, however, this would allow to apply your own rule for comparison. This
should  be passed in as a reference to a HASH with the keys sheet, tolerance,  sheet_tolerance
and optionally swap_check, error_limit and message (only relevant to method cmp_excel()).

    +-----------------+---------------------------------------------------------------------+
    | Key             | Description                                                         |
    +-----------------+---------------------------------------------------------------------+
    | sheet           | "|" seperated sheet names.                                          |
    | tolerance       | Number. Apply to all NUMBERS except on 'sheet'/'spec'. e.g. 10**-12 |
    | sheet_tolerance | Number. Apply to sheets/ranges in the spec. e.g. 0.20               |
    | spec            | Path to the specification file.                                     |
    | swap_check      | Number (optional) (1 or 0). Row swapping check. Default is 0.       |
    | error_limit     | Number (optional). Limit error per sheet. Default is 0.             |
    | message         | String (optional). Only required when calling method cmp_excel().   |
    +-----------------+---------------------------------------------------------------------+

=head1 What is "Visually" Similar?

This module uses the C<Spreadsheet::ParseExcel> module to parse Excel files, then compares the
parsed  data structure for differences. We ignore cetain components of the Excel file, such as
embedded fonts,  images, forms and annotations, and focus entirely on the layout of each Excel
page instead.  Future versions will likely support font and image comparisons, but not in this
initial release.

=head1 METHODS

=head2 cmp_excel($got, $exp, { ...rule... })

This function  will  tell  you  whether the two Excel files are "visually" different, ignoring
differences in  embedded fonts/images and metadata. Both $got and $exp can be either instances
of Spreadsheet::ParseExcel / file path (which is in turn passed to the Spreadsheet::ParseExcel
constructor). This one is for use in TEST MODE.

    use strict; use warnings;
    use Test::More no_plan => 1;
    use Test::Excel;

    cmp_excel('foo.xls', 'bar.xls', { message => 'EXCELSs are identical.' });

    # or

    my $foo = Spreadsheet::ParseExcel::Workbook->Parse('foo.xls');
    my $bar = Spreadsheet::ParseExcel::Workbook->Parse('bar.xls');
    cmp_excel($foo, $bar, { message => 'EXCELs are identical.' });

=cut

sub cmp_excel
{
    my $got  = shift;
    my $exp  = shift;
    my $rule = shift;

    _validate_rule($rule);
    $rule->{test} = 1;
    compare_excel($got, $exp, $rule);
}

=head2 compare_excel($got, $exp, { ...rule... })

This function  will  tell  you  whether the two Excel files are "visually" different, ignoring
differences in  embedded fonts/images and metadata. Both $got and $exp can be either instances
of Spreadsheet::ParseExcel / file path (which is in turn passed to the Spreadsheet::ParseExcel
constructor). This one is for use in STANDALONE MODE.

    use strict; use warnings;
    use Test::Excel;

    print "EXCELs are identical.\n"
        if compare_excel("foo.xls", "bar.xls");

=cut

sub compare_excel
{
    my $got  = shift;
    my $exp  = shift;
    my $rule = shift;

    croak("ERROR: Unable to locate file [$got].\n") unless (-f $got);
    croak("ERROR: Unable to locate file [$exp].\n") unless (-f $exp);
    _log_message("INFO: Excel comparison [$got] [$exp]\n") if $DEBUG;

    unless (blessed($got) && $got->isa('Spreadsheet::ParseExcel::WorkBook'))
    {
        $got = Spreadsheet::ParseExcel::Workbook->Parse($got)
            || croak("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$got]\n");
    }
    unless (blessed($exp) && $exp->isa('Spreadsheet::ParseExcel::WorkBook'))
    {
        $exp = Spreadsheet::ParseExcel::Workbook->Parse($exp)
            || croak("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$exp]\n");
    }

    my (@gotWorkSheets, @expWorkSheets);
    my ($message, $status, $error, $error_limit, $spec, $test, $TESTER);

    $status = 1;
    $test = $rule->{test}                if ((ref($rule) eq 'HASH') && exists($rule->{test}));
    _validate_rule($rule)                unless (defined($test) && ($test));
    $spec = parse($rule->{spec})         if exists($rule->{spec});
    $error_limit = $rule->{error_limit}  if exists($rule->{error_limit});
    $message     = $rule->{message}      if exists($rule->{message});
    $error_limit = $MAX_ERRORS_PER_SHEET unless defined $error_limit;

    @gotWorkSheets = $got->worksheets();
    @expWorkSheets = $exp->worksheets();

    $TESTER = Test::Builder->new if (defined($test) && ($test));
    if (scalar(@gotWorkSheets) != scalar(@expWorkSheets))
    {
        $error = "ERROR: Sheets count mismatch. ";
        $error .= "Got: [".scalar(@gotWorkSheets)."] exp: [".scalar(@expWorkSheets)."]\n";
        _log_message($error);
        if (defined($test) && ($test))
        {
            $TESTER->ok(0, $message);
            return;
        }
        return 0;
    }

    my ($i, @sheets);
    @sheets = split(/\|/,$rule->{sheet})
        if (exists($rule->{sheet}) && defined($rule->{sheet}));

    for ($i=0; $i<scalar(@gotWorkSheets); $i++)
    {
        my ($error_on_sheet);
        my ($gotWorkSheet, $expWorkSheet);
        my ($gotSheetName, $expSheetName);
        my ($gotRowMin, $gotRowMax, $gotColMin, $gotColMax);
        my ($expRowMin, $expRowMax, $expColMin, $expColMax);

        $error_on_sheet = 0;
        $gotWorkSheet   = $gotWorkSheets[$i];
        $expWorkSheet   = $expWorkSheets[$i];
        $gotSheetName   = $gotWorkSheet->get_name();
        $expSheetName   = $expWorkSheet->get_name();
        if (uc($gotSheetName) ne uc($expSheetName))
        {
            $error = "ERROR: Sheetname mismatch. Got: [$gotSheetName] exp: [$expSheetName].\n";
            _log_message($error);
            if (defined($test) && ($test))
            {
                $TESTER->ok(0, $message);
                return;
            }
            return 0;
        }

        ($gotRowMin, $gotRowMax) = $gotWorkSheet->row_range();
        ($gotColMin, $gotColMax) = $gotWorkSheet->col_range();
        ($expRowMin, $expRowMax) = $expWorkSheet->row_range();
        ($expColMin, $expColMax) = $expWorkSheet->col_range();

        if ($DEBUG > 1)
        {
            _log_message("\n");
            _log_message("INFO:[$gotSheetName]:[$gotRowMin][$gotColMin]:[$gotRowMax][$gotColMax]");
            _log_message("INFO:[$expSheetName]:[$expRowMin][$expColMin]:[$expRowMax][$expColMax]");
        }

        if (defined($gotRowMax) && defined($expRowMax) && ($gotRowMax != $expRowMax))
        {
            $error  = "\nERROR: Max row counts mismatch in sheet [$gotSheetName]. ";
            $error .= "Got[$gotRowMax] Expected: [$expRowMax]\n";
            _log_message($error);
            if (defined($test) && ($test))
            {
                $TESTER->ok(0, $message);
                return;
            }
            return 0;
        }

        if (defined($gotColMax) &&  defined($expColMax) && ($gotColMax != $expColMax))
        {
            $error  = "\nERROR: Max column counts mismatch in sheet [$gotSheetName]. ";
            $error .= "Got[$gotColMax] Expected: [$expColMax]\n";
            _log_message($error);
            if (defined($test) && ($test))
            {
                $TESTER->ok(0, $message);
                return;
            }
            return 0;
        }

        my ($row, $col, $swap);
        for ($row=$gotRowMin; $row<=$gotRowMax; $row++)
        {
            for ($col=$gotColMin; $col<=$gotColMax; $col++)
            {
                my ($gotData, $expData, $error);
                $gotData = $gotWorkSheet->{Cells}[$row][$col]->{Val};
                $expData = $expWorkSheet->{Cells}[$row][$col]->{Val};

                next if ( defined($spec)
                          &&
                          exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                          &&
                          ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $IGNORE) );

                if (defined($gotData) && defined($expData))
                {
                    if (($gotData =~ /^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$/)
                        &&
                        ($expData =~ /^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$/))
                    {
                        if (($gotData < $ALMOST_ZERO) && ($expData < $ALMOST_ZERO))
                        {
                            # Can be treated as the same.
                            next;
                        }
                        else
                        {
                            if (defined($rule))
                            {
                                my ($compare_with, $difference);
                                $difference = abs($expData - $gotData) / abs($expData);

                                if ( ( defined($spec)
                                       &&
                                       exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                                       &&
                                       ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $SPECIAL_CASE)
                                     )
                                     ||
                                     ( scalar(@sheets)
                                       &&
                                       grep(/$gotSheetName/,@sheets)
                                     ) )
                                {
                                    print "\nINFO: [NUMBER]:[$gotSheetName]:[SPC][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... "
                                        if $DEBUG > 1;
                                    $compare_with = $rule->{sheet_tolerance};
                                }
                                else
                                {
                                    print "\nINFO: [NUMBER]:[$gotSheetName]:[STD][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... "
                                        if $DEBUG > 1;
                                    $compare_with = $rule->{tolerance};
                                }

                                if ($compare_with < $difference)
                                {
                                    print "[FAIL]" if $DEBUG > 1;
                                    $difference = sprintf("%02f", $difference);
                                    $status = 0;
                                }
                                else
                                {
                                    $status = 1;
                                    print "[PASS]" if $DEBUG > 1;
                                }
                            }
                            else
                            {
                                print "\nINFO: [NUMBER]:[$gotSheetName]:[N/A][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... "
                                    if $DEBUG > 1;
                                if ($expData != $gotData)
                                {
                                    print "[FAIL]" if $DEBUG > 1;
                                    $status = 0;
                                }
                                else
                                {
                                    $status = 1;
                                    print "[PASS]" if $DEBUG > 1;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (uc($gotData) ne uc($expData))
                        {
                            _log_message("INFO: [STRING]:[$gotSheetName]:[$expData][$gotData] ... [FAIL]");
                            $status = 0;
                        }
                        else
                        {
                            $status = 1;
                            _log_message("INFO: [STRING]:[$gotSheetName]:[STD][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... [PASS]")
                                if $DEBUG > 1;
                        }
                    }

                    if (exists($rule->{swap_check}) && defined($rule->{swap_check}) && ($rule->{swap_check}))
                    {
                        if ($status == 0)
                        {
                            $error_on_sheet++;
                            push @{$swap->{exp}->{number_to_letter($col-1)}}, $expData;
                            push @{$swap->{got}->{number_to_letter($col-1)}}, $gotData;

                            if (($error_on_sheet >= $error_limit) && ($error_on_sheet % 2 == 0) && !_is_swapping($swap))
                            {
                                _log_message("ERROR: Max error per sheet reached.[$error_on_sheet]\n");
                                if (defined($test) && ($test))
                                {
                                    $TESTER->ok($status, $message);
                                    return;
                                }
                                return $status;
                            }
                        }
                    }
                }
            } # col

        if (($error_on_sheet >= $error_limit) && ($error_on_sheet % 2 == 0) && !_is_swapping($swap))
        {
            if (defined($test) && ($test))
            {
                $TESTER->ok($status, $message);
                return;
            }
            return $status;
        }

        } # row

        if (exists($rule->{swap_check}) && defined($rule->{swap_check}) && ($rule->{swap_check}))
        {
            if (($error_on_sheet > 0) && _is_swapping($swap))
            {
                print "\n\nWARN: SWAP OCCURRED.\n\n";
                $status = 1;
            }
        }
        print "INFO: [$gotSheetName]: ..... [OK].\n" if $DEBUG == 1;
    } # sheet


    if (defined($test) && ($test))
    {
        $TESTER->ok($status, $message);
        return;
    }
    return $status;
}

=head2 parse()

This method parse specification file provided by the user.  It  expects spec  file  to be in a
format mentioned below. Key and values are space seperated.

    sheet       Sheet1
    range       A3:B14
    range       B5:C5
    sheet       Sheet2
    range       A1:B2
    ignorerange B3:B8
    
They are grouped as sheet followed by one or more ranges.

    use strict; use warnings;
    use Test::Excel;

    my $data = Test::Excel::parse('spec-1.txt');

=cut

sub parse
{
    my $spec = shift;
    return unless defined $spec;

    croak("ERROR: Unable to locate spec file [$spec].\n")
        unless (-f $spec);

    my ($handle, $row, $sheet, $cells, $data);
    $handle = IO::File->new($spec)
        || croak("ERROR: Couldn't open file [$spec][$!].\n");

    $sheet = undef;
    $data  = undef;
    while ($row = <$handle>)
    {
        chomp($row);
        next unless $row =~ /\w/;
        next if $row =~ /^#/;

        if ($row =~ /^sheet\s+(.*)/i)
        {
            $sheet = $1;
        }
        elsif (defined($sheet) && ($row =~ /^range\s+(.*)/i))
        {
            $cells = Test::Excel::cells_within_range($1);
            foreach (@{$cells})
            {
                $data->{uc($sheet)}->{$_->{col}+1}->{$_->{row}} = $SPECIAL_CASE;
            }
        }
        elsif (defined($sheet) && ($row =~ /^ignorerange\s+(.*)/i))
        {
            $cells = Test::Excel::cells_within_range($1);
            foreach (@{$cells})
            {
                $data->{uc($sheet)}->{$_->{col}+1}->{$_->{row}} = $IGNORE;
            }
        }
        else
        {
            croak("ERROR: Invalid format data [$row] found in spec file.\n");
        }
    }
    $handle->close();

    return $data;
}

=head2 column_row()

This method accepts a cell address and returns column and row address as a list.

    use strict; use warnings;
    use Test::Excel;

    my $cell = 'A23';
    my ($col, $row) = Test::Excel::column_row($cell);

=cut

sub column_row
{
    my $cell = shift;
    return unless defined $cell;

    croak("ERROR: Invalid cell address [$cell].\n")
        unless ($cell =~ /([A-Za-z]+)(\d+)/);

    return ($1, $2);
}

=head2 letter_to_number()

This  method accepts a letter and returns back its equivalent number. This simply wraps around
Spreadsheet::ParseExcel::Utility::col2int().

    use strict; use warnings;
    use Test::Excel;

    my $number = Test::Excel::letter_to_number('AB');

=cut

sub letter_to_number
{
    my $letter = shift;
    return col2int($letter);
}

=head2 number_to_letter()

This  number  accepts  a  number  and  returns its equivalent letter. This simply wraps around 
Spreadsheet::ParseExcel::Utility::int2col().

    use strict; use warnings;
    use Test::Excel;

    my $letter = Test::Excel::number_to_letter(27);

=cut

sub number_to_letter
{
    my $number = shift;
    return int2col($number);
}

=head2 cells_within_range()

This method accepts address range and returns all cell address within the range.

    use strict; use warnings;
    use Test::Excel;

    my $range = 'A1:B3';
    my $cells = Test::Excel::cells_within_range($range);

=cut

sub cells_within_range
{
    my $range = shift;
    return unless defined $range;

    croak("ERROR: Invalid range [$range].\n")
        unless ($range =~ /(\w+\d+):(\w+\d+)/);

    my ($from, $to, $row, $col, $cells);
    my ($min_row, $min_col, $max_row, $max_col);

    $from = $1; $to = $2;
    ($min_col, $min_row) = column_row($from);
    ($max_col, $max_row) = column_row($to);
    $min_col = letter_to_number($min_col);
    $max_col = letter_to_number($max_col);

    for ($row = $min_row; $row <= $max_row; $row++)
    {
        for ($col = $min_col; $col <= $max_col; $col++)
        {
            push @{$cells}, { col => $col, row => $row };
        }
    }

    return $cells;
}

sub _is_swapping
{
    my $data = shift;
    return 0 unless defined $data;

    foreach (keys %{$data->{exp}})
    {
        my $exp = $data->{exp}->{$_};
        my $out = $data->{out}->{$_};

        return 0 if grep(/$exp->[0]/,@{$out});
    }
    return 1;
}

sub _log_message
{
    my $message = shift;
    return unless defined($message);

    print {*STDOUT} "\n".$message;
}

sub _validate_rule
{
    my $rule = shift;
    return unless defined $rule;

    croak("ERROR: Invalid RULE definitions. It has to be reference to a HASH.\n")
        unless (ref($rule) eq 'HASH');

    my ($keys, $valid);
    $keys = scalar(keys(%{$rule}));
    return if (($keys == 1) && exists($rule->{message}));

    croak("ERROR: Rule has more than 8 keys defined.\n")
        if $keys > 8;

    $valid = {'message'         => 1,
              'sheet'           => 2,
              'spec'            => 3,
              'tolerance'       => 4,
              'sheet_tolerance' => 5,
              'error_limit'     => 6,
              'swap_check'      => 7,
              'test'            => 8,};
    foreach (keys %{$rule})
    {
        croak("ERROR: Invalid key found in the rule definitions.\n")
            unless exists($valid->{$_});
    }

    if ((exists($rule->{spec}) && defined($rule->{spec}))
        ||
        (exists($rule->{sheet}) && defined($rule->{sheet})))
    {
        croak("ERROR: Missing key sheet_tolerance in the rule definitions.\n")
            unless (exists($rule->{sheet_tolerance}) && defined($rule->{sheet_tolerance}));
        croak("ERROR: Missing key tolerance in the rule definitions.\n")
            unless (exists($rule->{tolerance}) && defined($rule->{tolerance}));
    }
    else
    {
        if ( (exists($rule->{sheet_tolerance}) && defined($rule->{sheet_tolerance}))
             ||
             (exists($rule->{tolerance}) && defined($rule->{tolerance})) )
        {
            croak("ERROR: Missing key sheet/spec in the rule definitions.\n")
                unless ((exists($rule->{sheet}) && defined($rule->{sheet}))
                        ||
                        (exists($rule->{spec}) && defined($rule->{spec})));
        }
    }
}

=head1 DEBUG

Debug mode can be turned on or off by setting package variable $DEBUG, for example,

   $Test::Excel::DEBUG = 1;

You can set it anything greater than 1 for fine grained debug information. i.e.

   $Test::Excel::DEBUG = 2;

=head1 NOTES

It should be clearly noted that this module does not claim to provide a fool-proof  comparison
of  generated  Excels.  In fact there are still a number of ways in which I want to expand the 
 existing comparison functionality. This module I<is> actively being developed for a number of 
projects  I  am  currently  working on,  so  expect  many  changes  to happen. If you have any 
suggestions/comments/questions please feel free to contact me.

=head1 CAVEATS

Testing  of  large  Excels  can take a long time, this is because, well, we are doing a lot of 
computation.  In  fact,  this  module  test suite includes tests against several large Excels, 
however I am not including those in this distibution for obvious reasons.

=head1 BUGS

None  that I am aware of. Of course, if you find a bug, let me know, and I will be sure to fix
it.  This  is still a very early version, so it is always possible that I have just "gotten it 
wrong" in some places.

=head1 SEE ALSO

=over 4

=item C<Spreadsheet::ParseExcel> - I could not have written this without this module.

=back

=head1 ACKNOWLEDGEMENTS

=over 4

=item John McNamara (author of Spreadsheet::ParseExcel).

=item Kawai Takanori (author of Spreadsheet::ParseExcel::Utility).

=item Stevan Little (author of Test::PDF).

=back

=head1 AUTHOR

Mohammad S Anwar, E<lt>mohammad.anwar@yahoo.comE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2010-2011 by Mohammad S Anwar.

This  library  is free software; you can redistribute it and/or modify it under the same terms
as Perl itself.

=head1 DISCLAIMER

This  program  is  distributed  in  the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

=cut

1;
__END__
