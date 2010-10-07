package Test::Excel;

use strict; use warnings;

our $VERSION = '0.01';

use Carp;
use Test::Builder ();
use Scalar::Util 'blessed';
use Spreadsheet::ParseExcel;

require Exporter;

our @ISA    = qw(Exporter);
our @EXPORT = qw(cmp_excel compare_excel);

my $Test = Test::Builder->new;

sub cmp_excel($$;$) 
{
    my $got      = shift;
	my $expected = shift;
	my $message  = shift;

    unless (blessed($got) && $got->isa('Spreadsheet::ParseExcel::WorkBook')) 
	{
        $got = Spreadsheet::ParseExcel::Workbook->Parse($got) 
			|| croak("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$got]\n");
    }
    unless (blessed($expected) && $expected->isa('Spreadsheet::ParseExcel::WorkBook')) 
	{
        $expected = Spreadsheet::ParseExcel::Workbook->Parse($expected) 
			|| croak("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$expected]\n");
    }
	
	my ($i, $row, $col);
	my (@gotWorkSheets, @expectedWorkSheets);
	@gotWorkSheets      = $got->worksheets();
	@expectedWorkSheets = $expected->worksheets();
	if (scalar(@gotWorkSheets) != scalar(@expectedWorkSheets))
	{
		$Test->ok(0, $message);
        return;
	}
	
	for ($i=0; $i<scalar(@gotWorkSheets); $i++)
	{
		my ($gotWorkSheet, $expectedWorkSheet);
		my ($gotRowMin, $gotRowMax, $gotColMin, $gotColMax);
		my ($expectedRowMin, $expectedRowMax, $expectedColMin, $expectedColMax);
		
		$gotWorkSheet                      = $gotWorkSheets[$i];
		$expectedWorkSheet                 = $expectedWorkSheets[$i];
		($gotRowMin, $gotRowMax)           = $gotWorkSheet->row_range();
		($gotColMin, $gotColMax)           = $gotWorkSheet->col_range();
		($expectedRowMin, $expectedRowMax) = $expectedWorkSheet->row_range();
		($expectedColMin, $expectedColMax) = $expectedWorkSheet->col_range();
		if (defined($gotRowMax) 
			&& 
			defined($gotColMax) 
			&& 
			defined($expectedRowMax) 
			&& 
			defined($expectedColMax)
			&&
			($gotRowMax != $expectedRowMax) || ($gotColMax != $expectedColMax))
		{
			$Test->ok(0, $message);
			return;
		}
	
		for ($row=$gotRowMin; $row<=$gotRowMax; $row++) 
		{
			for ($col=$gotColMin; $col<=$gotColMax; $col++) 
			{
				my ($gotData, $expectedData);
				$gotData      = $gotWorkSheet->{Cells}[$row][$col]->{Val};
				$expectedData = $expectedWorkSheet->{Cells}[$row][$col]->{Val};
				
				if (defined($gotData) 
					&& 
					defined($expectedData) 
					&& 
					($gotData !~ m/$expectedData/i))
				{
					$Test->ok(0, $message);
					return;
				};
			}
		}
	}	
	$Test->ok(1, $message);
}

sub compare_excel($$) 
{
    my $got      = shift;
	my $expected = shift;

    unless (blessed($got) && $got->isa('Spreadsheet::ParseExcel::WorkBook')) 
	{
        $got = Spreadsheet::ParseExcel::Workbook->Parse($got) 
			|| croak("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$got]\n");
    }
    unless (blessed($expected) && $expected->isa('Spreadsheet::ParseExcel::WorkBook')) 
	{
        $expected = Spreadsheet::ParseExcel::Workbook->Parse($expected) 
			|| croak("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$expected]\n");
    }
	
	my (@gotWorkSheets, @expectedWorkSheets);
	@gotWorkSheets      = $got->worksheets();
	@expectedWorkSheets = $expected->worksheets();
	
	if (scalar(@gotWorkSheets) != scalar(@expectedWorkSheets))
	{
        return 0;
	}

	my ($i);	
	for ($i=0; $i<scalar(@gotWorkSheets); $i++)
	{
		my ($gotWorkSheet, $expectedWorkSheet);
		my ($gotRowMin, $gotRowMax, $gotColMin, $gotColMax);
		my ($expectedRowMin, $expectedRowMax, $expectedColMin, $expectedColMax);
		
		$gotWorkSheet                      = $gotWorkSheets[$i];
		$expectedWorkSheet                 = $expectedWorkSheets[$i];
		($gotRowMin, $gotRowMax)           = $gotWorkSheet->row_range();
		($gotColMin, $gotColMax)           = $gotWorkSheet->col_range();
		($expectedRowMin, $expectedRowMax) = $expectedWorkSheet->row_range();
		($expectedColMin, $expectedColMax) = $expectedWorkSheet->col_range();
		
		if (defined($gotRowMax) 
			&& 
			defined($gotColMax) 
			&& 
			defined($expectedRowMax) 
			&& 
			defined($expectedColMax)
			&&
			($gotRowMax != $expectedRowMax) || ($gotColMax != $expectedColMax))
		{
			return 0;
		}
	
		my ($row, $col);	
		for ($row=$gotRowMin; $row<=$gotRowMax; $row++) 
		{
			for ($col=$gotColMin; $col<=$gotColMax; $col++) 
			{
				my ($gotData, $expectedData);
				$gotData      = $gotWorkSheet->{Cells}[$row][$col]->{Val};
				$expectedData = $expectedWorkSheet->{Cells}[$row][$col]->{Val};
				
				if (defined($gotData) 
					&& 
					defined($expectedData) 
					&& 
					($gotData !~ m/$expectedData/i))
				{
					return 0;
				};
			}
		}
	}	
	return 1;
}

1;

__END__

=head1 NAME

Test::Excel - A module for testing and comparing Excel files

=head1 SYNOPSIS

  use Test::More no_plan => 1;
  use Test::Excel;
  
  cmp_excel('foo.xls', 'bar.xls', 'EXCELSs are identical.');
  
  # or
  
  my $foo = Spreadsheet::ParseExcel::Workbook->Parse('foo.xls');
  my $bar = Spreadsheet::ParseExcel::Workbook->Parse('bar.xls');
  cmp_excel($foo, $bar, 'EXCELs are identical.');
  
  # or even in standalone mode:
  
  use Test::Excel;
  print "EXCELs are identical.\n"
	  if compare_excel("foo.xls", "bar.xls");

=head1 DESCRIPTION

This module is meant to be used for testing custom generated Excel files, it provides two 
functions at the moment, which is C<cmp_excel> and C<compare_excel>. These can be used to compare two Excel files 
to see if they are I<visually> similar. The function C<cmp_excel> is for testing purpose where function C<compare_excel>
can be used as standalone. Future versions may include other testing functions.

=head2 What is "Visually" Similar?

This module uses the C<Spreadsheet::ParseExcel> module to parse Excel files, then compares the parsed data 
structure for differences. We ignore cetain components of the Excel file, such as embedded fonts, 
images, forms and annotations, and focus entirely on the layout of each Excel page instead. Future 
versions will likely support font and image comparisons, but not in this initial release.

=head2 Important Disclaimer

It should be clearly noted that this module does not claim to provide a fool-proof comparison of 
generated Excels. In fact there are still a number of ways in which I want to expand the existing 
comparison functionality. This module I<is> actively being developed for a number of projects I am 
currently working on, so expect many changes to happen. If you have any suggestions/comments/questions 
please feel free to contact me.

=head1 FUNCTIONS

=over 4

=item C<cmp_excel($got, $expected, ?$message)>

This function will tell you whether the two Excel files are "visually" different, ignoring differences 
in embedded fonts/images and metadata.

Both $got and $expected can be either instances of Spreadsheet::ParseExcel or a file path (which is in turn passed 
to the Spreadsheet::ParseExcel constructor).

=item C<compare_excel($got, $expected)>

This function will tell you whether the two Excel files are "visually" different, ignoring differences 
in embedded fonts/images and metadata in standalone mode.

Both $got and $expected can be either instances of Spreadsheet::ParseExcel or a file path (which is in turn passed 
to the Spreadsheet::ParseExcel constructor).

=back

=head1 CAVEATS

=head2 Testing Large Excels

Testing of large Excels can take a long time, this is because, well, we are doing a lot of
computation. In fact, this module test suite includes tests against several large Excels, 
however I am not including those in this distibution for obvious reasons.

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