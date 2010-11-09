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

Test::Excel - A module for testing and comparing Excel files

=head1 VERSION

Version 0.05

=cut

our $VERSION = '0.05';

$|=1;

our $DEBUG = 0;
Readonly my $ALMOST_ZERO  => 10**-16;
Readonly my $IGNORE       => 1;
Readonly my $SPECIAL_CASE => 2;

=head1 SYNOPSIS

  use Test::More no_plan => 1;
  use Test::Excel;

  cmp_excel('foo.xls', 'bar.xls', 'EXCELSs are identical.');

  # or

  my $foo = Spreadsheet::ParseExcel::Workbook->Parse('foo.xls');
  my $bar = Spreadsheet::ParseExcel::Workbook->Parse('bar.xls');
  cmp_excel($foo, $bar, undef, 'EXCELs are identical.');

  # or even in standalone mode:

  use Test::Excel;
  print "EXCELs are identical.\n"
      if compare_excel("foo.xls", "bar.xls");

=head1 DESCRIPTION

This module is meant to be used for testing custom generated Excel files, it 
provides two functions at the moment, which is C<cmp_excel> and C<compare_excel>. 
These can be used to compare two Excel files to see if they are I<visually> 
similar. The function C<cmp_excel> is for testing purpose where function C<compare_excel>
can be used as standalone. Future versions may include other testing functions.

=head2 Definition of Rule

The new paramter has been added to both method cmp_excel() and method compare_excel() 
called rule. This is optional, however, this would allow to apply your own rule for
comparison. This should be passed in as reference to a HASH with the keys 'sheet',
'tolerance' and 'sheet_tolerance'.

=over 3

=item sheet: "|" seperated sheet name.

The tolerance defined by the key sheet_tolerance would apply on these sheets.
Example: 'Sheet1|Sheet2'

=item tolerance: Number.

This would apply to all the sheets in the excel when comparing numbers.
Example: 10**-12

=item sheet_tolerance: Number.

These rule would be applied to all the sheets defined by the key sheet.
Example: 0.20

=back

=head2 What is "Visually" Similar?

This module uses the C<Spreadsheet::ParseExcel> module to parse Excel files, 
then compares the parsed data structure for differences. We ignore cetain 
components of the Excel file, such as embedded fonts, images, forms and 
annotations, and focus entirely on the layout of each Excel page instead. 
Future versions will likely support font and image comparisons, but not 
in this initial release.

=cut

my $Test = Test::Builder->new;

=head1 METHODS

=head2 cmp_excel(got, expected, rule, message)

This function will tell you whether the two Excel files are "visually" 
different, ignoring differences in embedded fonts/images and metadata.

Both $got and $expected can be either instances of Spreadsheet::ParseExcel 
or a file path (which is in turn passed to the Spreadsheet::ParseExcel constructor).

=cut

sub cmp_excel
{
    my $got     = shift;
    my $exp     = shift;
    my $rule    = shift;
    my $message = shift;

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

    my (@gotWorkSheets, @expWorkSheets, $error, $spec);
    if (defined($rule))
    {
        croak("ERROR: Invalid RULE definition.\n")
            unless ((ref($rule) eq 'HASH') && (exists $rule->{tolerance}) && (exists $rule->{sheet_tolerance}));

        if (exists $rule->{spec})
        {
            $spec = parse($rule->{spec});
        }
        else
        {
            croak("ERROR: Missing sheet key in the rule definitions.\n")
                unless (exists $rule->{sheet});
        }
    }

    @gotWorkSheets = $got->worksheets();
    @expWorkSheets = $exp->worksheets();

    if (scalar(@gotWorkSheets) != scalar(@expWorkSheets))
    {
        $error = "ERROR: Sheets count mismatch. ";
        $error .= "Got: [".scalar(@gotWorkSheets)."] exp: [".scalar(@expWorkSheets)."]\n";
        _dump_error($error);
        return 0;
    }

    my ($i);
    for ($i=0; $i<scalar(@gotWorkSheets); $i++)
    {
        my ($gotWorkSheet, $expWorkSheet);
        my ($gotSheetName, $expSheetName);
        my ($gotRowMin, $gotRowMax, $gotColMin, $gotColMax);
        my ($expRowMin, $expRowMax, $expColMin, $expColMax);

        $gotWorkSheet = $gotWorkSheets[$i];
        $expWorkSheet = $expWorkSheets[$i];
        $gotSheetName = $gotWorkSheet->get_name();
        $expSheetName = $expWorkSheet->get_name();
        if (uc($gotSheetName) ne uc($expSheetName))
        {
            $Test->ok(0, $message);
            return;
        }

        ($gotRowMin, $gotRowMax) = $gotWorkSheet->row_range();
        ($gotColMin, $gotColMax) = $gotWorkSheet->col_range();
        ($expRowMin, $expRowMax) = $expWorkSheet->row_range();
        ($expColMin, $expColMax) = $expWorkSheet->col_range();

        if (defined($gotRowMax) && defined($expRowMax) && ($gotRowMax != $expRowMax))
        {
            $Test->ok(0, $message);
            return;
        }
        if (defined($gotColMax) &&  defined($expColMax) && ($gotColMax != $expColMax))
        {
            $Test->ok(0, $message);
            return;
        }

        my ($row, $col);
        for ($row=$gotRowMin; $row<=$gotRowMax; $row++)
        {
            for ($col=$gotColMin; $col<=$gotColMax; $col++)
            {
                my ($gotData, $expData);
                $gotData = $gotWorkSheet->{Cells}[$row][$col]->{Val};
                $expData = $expWorkSheet->{Cells}[$row][$col]->{Val};

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
                            if (defined($rule) && ref($rule) eq 'HASH')
                            {
                                my ($compare_with, $sheet, $difference);

                                $sheet = $rule->{sheet}; 
                                if ( defined($spec) 
                                    && 
                                    exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                                    &&
                                    ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $IGNORE) )
                                {
                                    # Data can be ignored.
                                    next;
                                }
                                elsif ( ( defined($spec) 
                                          &&    
                                          exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                                          &&
                                          ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $SPECIAL_CASE)
                                        )
                                        ||
                                        ($gotSheetName =~ /$sheet/) )
                                {
                                    $compare_with = $rule->{sheet_tolerance};
                                }
                                else
                                {        
                                    $compare_with = $rule->{tolerance};
                                }
                                $difference = abs($expData - $gotData) / abs($expData);
                                if ($compare_with < $difference)
                                {
                                    $Test->ok(0, $message);
                                    return;
                                }
                                print "[PASS]\n" if $DEBUG;
                            }
                            else
                            {
                                if ($expData != $gotData)
                                {
                                    $Test->ok(0, $message);
                                    return;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (uc($gotData) ne uc($expData))
                        {
                            $Test->ok(0, $message);
                            return;
                        }
                    }
                }
            } # col
        } # row
    } # sheet
    
    $Test->ok(1, $message);
}

=head2 compare_excel(got, expected, rule)

This function will tell you whether the two Excel files are "visually" 
different, ignoring differences in embedded fonts/images and metadata in standalone mode.

Both $got and $expected can be either instances of Spreadsheet::ParseExcel 
or a file path (which is in turn passed to the Spreadsheet::ParseExcel constructor).

=cut

sub compare_excel
{
    my $got  = shift;
    my $exp  = shift;
    my $rule = shift;

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

    my (@gotWorkSheets, @expWorkSheets, $error, $spec);
    if (defined($rule))
    {
        croak("ERROR: Invalid RULE definition.\n")
            unless ((ref($rule) eq 'HASH') && (exists $rule->{tolerance}) && (exists $rule->{sheet_tolerance}));
        
        if (exists $rule->{spec})
        {
            $spec = parse($rule->{spec});
        }
        else
        {
            croak("ERROR: Missing sheet key in the rule definitions.\n")
                unless (exists $rule->{sheet});
        }
    }

    @gotWorkSheets = $got->worksheets();
    @expWorkSheets = $exp->worksheets();

    if (scalar(@gotWorkSheets) != scalar(@expWorkSheets))
    {
        $error = "ERROR: Sheets count mismatch. ";
        $error .= "Got: [".scalar(@gotWorkSheets)."] exp: [".scalar(@expWorkSheets)."]\n";
        _dump_error($error);
        return 0;
    }

    my ($i);
    for ($i=0; $i<scalar(@gotWorkSheets); $i++)
    {
        my ($gotWorkSheet, $expWorkSheet);
        my ($gotSheetName, $expSheetName);
        my ($gotRowMin, $gotRowMax, $gotColMin, $gotColMax);
        my ($expRowMin, $expRowMax, $expColMin, $expColMax);

        $gotWorkSheet = $gotWorkSheets[$i];
        $expWorkSheet = $expWorkSheets[$i];
        $gotSheetName = $gotWorkSheet->get_name();
        $expSheetName = $expWorkSheet->get_name();
        if (uc($gotSheetName) ne uc($expSheetName))
        {
            $error = "ERROR: Sheetname mismatch. Got: [$gotSheetName] exp: [$expSheetName].\n";
            _dump_error($error);
            return 0;
        }

        ($gotRowMin, $gotRowMax) = $gotWorkSheet->row_range();
        ($gotColMin, $gotColMax) = $gotWorkSheet->col_range();
        ($expRowMin, $expRowMax) = $expWorkSheet->row_range();
        ($expColMin, $expColMax) = $expWorkSheet->col_range();

        if (defined($gotRowMax) && defined($expRowMax) && ($gotRowMax != $expRowMax))
        {
            $error = "ERROR: Max row counts mismatch in sheet [$gotSheetName]. ";
            $error .= "Got[$gotRowMax] Expected: [$expRowMax]\n";
            _dump_error($error);
            return 0;
        }
        if (defined($gotColMax) &&  defined($expColMax) && ($gotColMax != $expColMax))
        {
            $error = "ERROR: Max column counts mismatch in sheet [$gotSheetName]. ";
            $error .= "Got[$gotColMax] Expected: [$expColMax]\n";
            _dump_error($error);
            return 0;
        }

        my ($row, $col);
        for ($row=$gotRowMin; $row<=$gotRowMax; $row++)
        {
            for ($col=$gotColMin; $col<=$gotColMax; $col++)
            {
                my ($gotData, $expData);
                $gotData = $gotWorkSheet->{Cells}[$row][$col]->{Val};
                $expData = $expWorkSheet->{Cells}[$row][$col]->{Val};

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
                            if (defined($rule) && ref($rule) eq 'HASH')
                            {
                                my ($compare_with, $sheet, $difference);

                                $sheet = $rule->{sheet}; 
                                if ( defined($spec) 
                                    && 
                                    exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                                    &&
                                    ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $IGNORE) )
                                {
                                    # Data can be ignored.
                                    next;
                                }
                                elsif ( ( defined($spec) 
                                          &&    
                                          exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                                          &&
                                          ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $SPECIAL_CASE)
                                        )
                                        ||
                                        ($gotSheetName =~ /$sheet/) )
                                {
                                    print "INFO: [NUMBER]:[$gotSheetName]:[SPC][".($row+1)."][".($col+1)."] ... "
                                        if $DEBUG;
                                    $compare_with = $rule->{sheet_tolerance};
                                }
                                else
                                {        
                                    print "INFO: [NUMBER]:[$gotSheetName]:[STD][".($row+1)."][".($col+1)."] ... "
                                        if $DEBUG;
                                    $compare_with = $rule->{tolerance};
                                }
                                $difference = abs($expData - $gotData) / abs($expData);
                                if ($compare_with < $difference)
                                {
                                    $difference = sprintf("%02f", $difference);
                                    $error = "ERROR: [NUMBER]:[$gotSheetName]:Expected: [$expData] Got: [$gotData] Diff [$difference].\n";
                                    _dump_error($error);
                                    return 0;
                                }
                                print "[PASS]\n" if $DEBUG;
                            }
                            else
                            {
                                print "INFO: [NUMBER]:[$gotSheetName]:[N/A][".($row+1)."][".($col+1)."] ... "
                                    unless $DEBUG;
                                if ($expData != $gotData)
                                {
                                    $error = "ERROR: [NUMBER]:[$gotSheetName]:Expected: [$expData] Got: [$gotData].\n";
                                    _dump_error($error);
                                    return 0;
                                }
                                print "[PASS]\n" if $DEBUG;
                            }
                        }
                    }
                    else
                    {
                        if (uc($gotData) ne uc($expData))
                        {
                            $error = "ERROR: [STRING]:[$gotSheetName]: Expected [$expData] Got [$gotData].\n";
                            _dump_error($error);
                            return 0;
                        }
                        else
                        {
                            print "INFO: [STRING]:[$gotSheetName]:[STD][".($row+1)."][".($col+1)."] ... [PASS]\n"
                                if $DEBUG;
                        }
                    }
                }
            } # col
        } # row
    } # sheet

    return 1;
}

=head2 parse()

This method parse spec file provided by the user. It expects spec file to be 
in a format mentioned below:

   Sheet  Sheet1
   Range  A3:B14
   Range  B5:C5
   Sheet  Sheet2
   Range  A1:B2

=cut

sub parse
{
    my $spec = shift;
    croak("ERROR: Spec file [$spec] doesn't exist.\n")
        unless (-f $spec);

    my ($handle, $row, $sheet, $cells, $data);    
    $handle = IO::File->new($spec)
        || croak("ERROR: Couldn't open file [$spec][$!].\n");

    $sheet = undef;        
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

    # You should expect these values:    
    # $col => 'A'
    # $row => 23

=cut

sub column_row
{
    my $cell = shift;
    return unless defined $cell;

    croak("ERROR: Invalid cell address [$cell].\n")
        unless ($cell =~ /([A-Za-z]+)(\d+)/);

    return($1, $2);
}

=head2 letter_to_number()

This method accepts a letter and returns back its equivalent number.
This simply wraps around Spreadsheet::ParseExcel::Utility::col2int().

    use strict; use warnings;
    use Test::Excel;

    my $number = Test::Excel::letter_to_number('AB');

    # You should expect $number to be 27.

=cut

sub letter_to_number
{
    my $letter = shift;
    return col2int($letter);
}

=head2 number_to_letter()

This number accepts a number and returns its equivalent letter.
This simply wraps around Spreadsheet::ParseExcel::Utility::int2col().

    use strict; use warnings;
    use Test::Excel;

    my $letter = Test::Excel::number_to_letter(27);

    # You should expect $letter to be 'AB'.

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

    # $cells would have something like below:
    # [ {row => 1, col => 0},
    #   {row => 1, col => 1},
    #   {row => 2, col => 0},
    #   {row => 2, col => 1},
    #   {row => 3, col => 0},
    #   {row => 3, col => 1} ]

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

    for($row = $min_row; $row <= $max_row; $row++)
    {
        for($col = $min_col; $col <= $max_col; $col++)
        {
            push @{$cells}, { col => $col, row => $row };
        }
    }

    return $cells;
}

=head2 _dump_error()

This is an internal method that dumps the message to STDOUT.

=cut

sub _dump_error
{
    my $message = shift;
    return unless defined($message);

    print {*STDOUT} "\n".$message;
}

=head2 Important Disclaimer

It should be clearly noted that this module does not claim to provide a 
fool-proof comparison of generated Excels. In fact there are still a number 
of ways in which I want to expand the existing comparison functionality. 
This module I<is> actively being developed for a number of projects I am 
currently working on, so expect many changes to happen. If you have any 
suggestions/comments/questions please feel free to contact me.

=head1 CAVEATS

=head2 Testing Large Excels

Testing of large Excels can take a long time, this is because, well, we are 
doing a lot of computation. In fact, this module test suite includes tests 
against several large Excels, however I am not including those in this distibution 
for obvious reasons.

=head1 TO DO

=over 4

=item More functions for more testing

=item Testing of font data

=item Testing of embedded image data

=back

=head1 BUGS

None that I am aware of. Of course, if you find a bug, let me know, and I will be 
sure to fix it. This is still a very early version, so it is always possible that 
I have just "gotten it wrong" in some places. 

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

Copyright 2010 by Mohammad S Anwar.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself. 

=cut

1;
__END__