package Test::Excel;

use strict; use warnings;

our $VERSION = '0.03';

use Carp;
use Readonly;
use Data::Dumper;
use Test::Builder ();
use Scalar::Util 'blessed';
use Spreadsheet::ParseExcel;

require Exporter;

our @ISA    = qw(Exporter);
our @EXPORT = qw(cmp_excel compare_excel);

Readonly my $ALMOST_ZERO => 10**-16;

my $Test = Test::Builder->new;

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
    
    my (@gotWorkSheets, @expWorkSheets, $error);
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
                        ($expData=~ /^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$/))
                    {
                        if (($gotData < $ALMOST_ZERO) && ($expData < $ALMOST_ZERO))
                        {
                            # Can be treated as the same.
                        }
                        else
                        {
                            if (defined($rule) && ref($rule) eq 'HASH')
                            {
                                my ($compare_with, $sheet, $difference);
                                
                                $sheet = $rule->{sheet};
                                $difference = abs($expData - $gotData) / abs($expData);
                                if ($gotSheetName =~ /$sheet/)
                                {
                                    $compare_with = $rule->{sheet_tolerance};
                                }
                                else
                                {
                                    $compare_with = $rule->{tolerance};
                                }
                                if ($compare_with < $difference)
                                {
                                    $Test->ok(0, $message);
                                    return;
                                }    
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

sub compare_excel
{
    my $got  = shift;
    my $exp  = shift;
    my $rule = shift;
	print Dumper($got);
	print Dumper($exp);
	print Dumper($rule);
    
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

	if (defined($rule) && !((ref($rule) eq 'HASH') && (exists $rule->{sheet}) && (exists $rule->{tolerance}) && (exists $rule->{sheet_tolerance})))
	{
		die("ERROR: Invalid RULE definition. Rule should be passed in as reference to a HASH with keys sheet, tolerance and sheet_tolerance.\n");
	};
		
    my (@gotWorkSheets, @expWorkSheets, $error);
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
                        ($expData=~ /^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$/))
                    {
                        if (($gotData < $ALMOST_ZERO) && ($expData < $ALMOST_ZERO))
                        {
                            # Can be treated as the same.
                        }
                        else
                        {
                            if (defined($rule) && ref($rule) eq 'HASH')
                            {
                                my ($compare_with, $sheet, $difference);
                                
                                $sheet = $rule->{sheet};
                                $difference = abs($expData - $gotData) / abs($expData);
                                if ($gotSheetName =~ /$sheet/)
                                {
                                    $compare_with = $rule->{sheet_tolerance};
                                }
                                else
                                {
                                    $compare_with = $rule->{tolerance};
                                }
                                if ($compare_with < $difference)
                                {
                                    $error = "ERROR: [NUMBER]:[$gotSheetName]:Expected: [$expData] Got: [$gotData].\n";
                                    _dump_error($error);
                                    return 0;
                                }    
                            }
                            else
                            {
                                if ($expData != $gotData)
                                {
                                    $error = "ERROR: [NUMBER]:[$gotSheetName]:Expected: [$expData] Got: [$gotData].\n";
                                    _dump_error($error);
                                    return 0;
                                }
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
                    }
                }
            } # col
        } # row    
    } # sheet
    
    return 1;
}

sub _dump_error
{
    my $message = shift;
    return unless defined($message);
    
    print {*STDOUT} $message;
}

1;

__END__

=head1 NAME

Test::Excel - A module for testing and comparing Excel files

=head1 VERSION

Version 0.3

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

=head1 Definition of Rule

The new paramter has been added to both method cmp_excel() and method compare_excel() 
called rule. This is optional, however, this would allow to apply your own rule for
comparison. This should be passed in as reference to a HASH with the keys 'sheet',
'tolerance' and 'sheet_tolerance'.

=over 3

=item sheet: "|" seperated sheet name.

The tolerance defined by the key sheet_tolerance would apply on these sheets

=item tolerance: Number in the form of 10**-12.

This would apply to all the sheets in the excel when comparing numbers.

=item sheet_tolerance: Something like 0.20.

These rule would be applied to all the sheets defined by the key sheet.

=back

=head2 What is "Visually" Similar?

This module uses the C<Spreadsheet::ParseExcel> module to parse Excel files, 
then compares the parsed data structure for differences. We ignore cetain 
components of the Excel file, such as embedded fonts, images, forms and 
annotations, and focus entirely on the layout of each Excel page instead. 
Future versions will likely support font and image comparisons, but not 
in this initial release.

=head2 Important Disclaimer

It should be clearly noted that this module does not claim to provide a 
fool-proof comparison of generated Excels. In fact there are still a number 
of ways in which I want to expand the existing comparison functionality. 
This module I<is> actively being developed for a number of projects I am 
currently working on, so expect many changes to happen. If you have any 
suggestions/comments/questions please feel free to contact me.

=head1 FUNCTIONS

=over 4

=item C<cmp_excel($got, $expected, $rule, $message)>

This function will tell you whether the two Excel files are "visually" 
different, ignoring differences in embedded fonts/images and metadata.

Both $got and $expected can be either instances of Spreadsheet::ParseExcel 
or a file path (which is in turn passed to the Spreadsheet::ParseExcel constructor).

=item C<compare_excel($got, $expected, $rule)>

This function will tell you whether the two Excel files are "visually" 
different, ignoring differences in embedded fonts/images and metadata in standalone mode.

Both $got and $expected can be either instances of Spreadsheet::ParseExcel 
or a file path (which is in turn passed to the Spreadsheet::ParseExcel constructor).

=back

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

None that I am aware of. Of course, if you find a bug, let me know, and I will be sure to fix it. This is still 
a very early version, so it is always possible that I have just "gotten it wrong" in some places. 

=head1 SEE ALSO

=over 4

=item C<Spreadsheet::ParseExcel> - I could not have written this without this module. 

=back

=head1 ACKNOWLEDGEMENTS

=over 4

=item John McNamara (author of Spreadsheet::ParseExcel).

=item Stevan Little (author of Test::PDF).

=back

=head1 AUTHOR

Mohammad S Anwar, E<lt>mohammad.anwar@yahoo.comE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2010 by Mohammad S Anwar.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself. 

=cut