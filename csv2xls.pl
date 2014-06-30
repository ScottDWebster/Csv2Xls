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

#use strict;
use warnings;
use utf8;

use Spreadsheet::WriteExcel;
use GetOpt::Long;
use File::Spec;

open (CSVFILE, 'On_Order_Report_By_Sku.csv');
my $workbook = Spreadsheet::WriteExcel->new('On_Order_Report_By_Sku.xls');
$worksheet1 = $workbook->add_worksheet("Sheet1");

$LineNum = 1;
while (<CSVFILE>)
{
	chomp;
	@LineArray = split(',', $_);
	$LineRef = \@LineArray;
	$worksheet1->write_row($LineNum++, 0, $LineRef);
}
