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

#use strict;
use warnings;
use utf8;
use Spreadsheet::WriteExcel;
use Win32 qw( CSIDL_PERSONAL );
use Win32::GUI();

# Variable Declarations
my ($debug, $lines, $CSV_FileName, $XLS_FileName, $XLS_DefaultFileName, $LineNum,
	@LineArray, $LineRef, $element, $WorkSheet, @InFileFilter, @OutFileFilter,
	$MyDocsPath, $Win);

# Look up the "My Documents" path 
$MyDocsPath = Win32::GetFolderPath(CSIDL_PERSONAL);
@InFileFilter = (
				"Comma Separated Value", "*.csv",
				"All Files", "*"
				);
@OutFileFilter = ("Excel Spreadsheet", "*.xls");

$debug = 0;	# 0 = off, 1 = on, 2 = verbose (show element modifications)

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

$lines = 0;
# Open csv file fo count lines
open (CSVFILE, "$CSV_FileName");
$lines += tr/\n/\n/ while sysread(CSVFILE, $_, 2 ** 16);
#while (<CSVFILE>)
#{ $lines++ }
close (CSVFILE);

# Open filhandle to csv file for processing
open (CSVFILE, "$CSV_FileName");
# Create workbook object
my $WorkBook = Spreadsheet::WriteExcel->new("$XLS_FileName");
# Create worksheet object
my $WorkSheet1 = $WorkBook->add_worksheet("Sheet1");
# Set option to keep leading zeros
$WorkSheet1->keep_leading_zeros();


$Win = new Win32::GUI::Window(
	-title	=>	"Csv2Xls-gui",
	-width	=>	300,
	-height	=>	150,
	-name	=>	"Progress"
) or print_and_die("new Window");

$Win->AddLabel(
	-name	=>	"PRCSNG",
	-text	=>	"Processing...",
	-left	=>	0, #$Win->Width / 8,
	-top	=>	50
);

$Win->AddProgressBar(
	-name	=>	"PB",
	-top	=>	$Win->Height / 4,
	-left	=>	$Win->Width / 8,
	-width	=>	($Win->Width / 4) * 3,
	-height =>	$Win->Height / 4,
	-smooth	=>	1
);
$Win->PB->SetPos(0);

$Win->Show;
$LineNum = 0;

while (<CSVFILE>)
{
	$PctDone = 100 * ($LineNum / $lines);
	if($debug){print("$PctDone = $LineNum / $lines\n");}
	$Win->PB->SetPos( $PctDone );
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

sub print_and_die {
    my($text) = @_;
    my $err = Win32::GetLastError();
    die "$text: Error $err\n";
}
