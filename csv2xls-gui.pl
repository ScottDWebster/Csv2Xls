#!/usr/bin/env perl 
#===============================================================================
#
#         FILE: csv2xls-gui.pl
#
#        USAGE: ./csv2xls-gui.pl  
#
#  DESCRIPTION: Converts a comma separated values (.csv) file to and Excel
#  				spreadsheet (.xls), leaving leading zeros in place (Excel
#  				removes leading zeros when importing csv files).  This version
#  				uses a GUI to choose the input and output files.
#
#      OPTIONS: ---
# REQUIREMENTS: Win32::GUI, Spreadsheet::WriteExcel
#         BUGS: ---
#        NOTES: ---
#       AUTHOR: Scott Webster (sdw), ScottW@ChefCentral.com
# ORGANIZATION: Chef Central, Inc.
#      VERSION: 1.0
#      CREATED: 7/5/2014 1:27:46 PM
#     REVISION: ---
#===============================================================================

use strict;
use warnings;
use utf8;
use Spreadsheet::WriteExcel;
use Win32 qw( CSIDL_PERSONAL );
use Win32::GUI();

# Variable Declarations
my ($debug, $CSV_FileName, $XLS_FileName, $XLS_DefaultFileName, $LineNum,
	@LineArray, $LineRef, $element, $WorkSheet, @InFileFilter, @OutFileFilter,
	$MyDocsPath);

# Look up the "My Documents" path 
$MyDocsPath = Win32::GetFolderPath(CSIDL_PERSONAL);
@InFileFilter = (
				"Comma Separated Value", "*.csv",
				"All Files", "*"
				);
@OutFileFilter = ("Excel Spreadsheet", "*.xls");

$debug = 1;	# 0 = off, 1 = on, 2 = verbose (show element modifications)

# Get file name from GUI dialog
$CSV_FileName = Win32::GUI::GetOpenFileName(
					-title => "Input CSV File to Open",
					-directory => $MyDocsPath,
					-filemustexiest => 1,
					-defaultextension => "csv",
					-filter => \@InFileFilter 
					);

if($debug){print("\$CSV_FileName = $CSV_FileName\n");}
# assign $CSV_FileName to $XLS_FileName
$XLS_DefaultFileName = $CSV_FileName;
# Change file extension from csv to xls
$XLS_DefaultFileName =~ s/.[Cc][Ss][Vv]$/.xls/;
if($debug){print("\$XLS_DefaultFileName = $XLS_DefaultFileName\n");}
# Get file name from GUI dialog
$XLS_FileName = Win32::GUI::GetSaveFileName(
					-title => "Output Excel File to Save",
					-file => $XLS_DefaultFileName,
					-filter => \@OutFileFilter
					);
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
		if($debug > 1){print("BEFORE \$element = [$element]\t");}
		$element =~ s/^\s*(.+?)\s*$/$1/;
		if($debug > 1){print("AFTER  \$element = [$element]\n");}
	}
#	Create a reference to the array
	$LineRef = \@LineArray;
#	print the row and increment the line number counter
	$WorkSheet1->write_row($LineNum++, 0, $LineRef);
}
