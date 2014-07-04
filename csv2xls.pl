#!/usr/bin/env perl 
#===============================================================================
#
#         FILE: cvs2xls.pl
#
#        USAGE: ./cvs2xls.pl  
#
#  DESCRIPTION: 
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
use GetOpt::Long;
use File::Spec;

# Variable Declaration
my ($debug, $CSV_FileName, $XLS_FileName, $LineNum, @LineArray, $LineRef, $WorkSheet);
my ($element);

$debug = 1;

# debug print
if($debug){print ("\$ARGV[0] = $ARGV[0]\n");}
# get filename from Arguemant vector zero
$CSV_FileName = $ARGV[0];
if($debug){print("\$CSV_FileName = $CSV_FileName\n");}
# assign $CSV_FileName to $XLS_FileName
$XLS_FileName = $CSV_FileName;
# Change file extension from csv to xls
$XLS_FileName =~ s/.csv$/.xls/;
if($debug){print("\$XLS_FileName = $XLS_FileName\n");}

open (CSVFILE, "$CSV_FileName");
my $WorkBook = Spreadsheet::WriteExcel->new("$XLS_FileName");
my $WorkSheet1 = $WorkBook->add_worksheet("Sheet1");
$WorkSheet1->keep_leading_zeros();

while (<CSVFILE>)
{
	chomp;
	@LineArray = split(',', $_);
#	Remove leading and trailing whitespace from each element in the array
	foreach $element (@LineArray)
	{
		if($debug){print("BEFORE \$element = [$element]\n");}
		$element =~ s/^\s*(.+?)\s*$/$1/;
		if($debug){print("AFTER  \$element = [$element]\n");}
	}
	$LineRef = \@LineArray;
	$WorkSheet1->write_row($LineNum++, 0, $LineRef);
}
