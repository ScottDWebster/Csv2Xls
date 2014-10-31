#!/usr/bin/env perl 
#===============================================================================
#
#         FILE: cvs2xls2excel.pl
#
#        USAGE: ./cvs2xls.pl FILENAME 
#
#  DESCRIPTION: Generates an Excel spreadsheet from a csv (comma separated value)
#  				file, removing leading and trailing whitespace but retaining
#  				leading zeros, then opens the resultant file in Excel
#
#      OPTIONS: ---
# REQUIREMENTS: Spreadsheet::WriteExcel
#         BUGS: ---
#        NOTES: ---
#       AUTHOR: Scott Webster (sdw), ScottW@ChefCentral.com
# ORGANIZATION: Chef Central, Inc.
#      VERSION: 1.0
#      CREATED: 6/2/2014 4:17:56 PM
#     REVISION: 001
#===============================================================================

use strict;
use warnings;
use utf8;
use Spreadsheet::WriteExcel;

# Variable Declarations
my ($debug, $CSV_FileName, $XLS_FileName, $LineNum, @LineArray, $LineRef);
my ($element, $WorkSheeti, $Excel);

$debug = 0;	# 0 = off, 1 = on

$Excel = "C:/PROGRA~1/MICROS~2/OFFICE11/EXCEL.EXE";

# debug print
if($debug){print ("\$ARGV[0] = $ARGV[0]\n");}
# get filename from Argument vector zero
$CSV_FileName = $ARGV[0];
if($debug){print("\$CSV_FileName = $CSV_FileName\n");}
# assign $CSV_FileName to $XLS_FileName
$XLS_FileName = $CSV_FileName;
# Change file extension from csv to xls
$XLS_FileName =~ s/.[Cc][Ss][Vv]$/.xls/;
if($debug){print("\$XLS_FileName = $XLS_FileName\n");}

open (CSVFILE, "$CSV_FileName");
# Create workbook object
my $WorkBook = Spreadsheet::WriteExcel->new("$XLS_FileName");
# Create worksheet object
my $WorkSheet1 = $WorkBook->add_worksheet("Sheet1");
# Set option to keep leading zeros
$WorkSheet1->keep_leading_zeros();

while (<CSVFILE>)
{
#	Remove EOL (end of line) character(s)
	chomp;
	@LineArray = split(',', $_);
#	Remove leading and trailing whitespace from each element in the array
	foreach $element (@LineArray)
	{
		if($debug){print("BEFORE \$element = [$element]\n");}
		$element =~ s/^\s*(.+?)\s*$/$1/;
		if($debug){print("AFTER  \$element = [$element]\n");}
	}
#	Create a reference to the array
	$LineRef = \@LineArray;
#	print the row and increment the line number counter
	$WorkSheet1->write_row($LineNum++, 0, $LineRef);
}

$WorkBook->close();
exec ($Excel, $XLS_FileName) or die "Can't exec \$Excel [$Excel]";
