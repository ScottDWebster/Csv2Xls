#!/usr/bin/env perl 
#===============================================================================
#
#         FILE: FileExtConv.pl
#
#        USAGE: ./FileExtConv.pl  
#
#  DESCRIPTION: Test file extention conversion
#
#      OPTIONS: ---
# REQUIREMENTS: ---
#         BUGS: ---
#        NOTES: ---
#       AUTHOR: Scott Webster (sdw), ScottW@ChefCentral.com
# ORGANIZATION: Chef Central, Inc.
#      VERSION: 1.0
#      CREATED: 6/30/2014 1:05:11 PM
#     REVISION: ---
#===============================================================================

use strict;
use warnings;
use utf8;

my ($debug, $FileName, $NewFileName);

$debug = 1;

# debug print
if($debug){print ("\$ARGV[0] = $ARGV[0]\n");}
# get filename from Arguemant vector zero
$FileName = $ARGV[0];
if($debug){print("\$FileName = $FileName\n");}
# assign $FileName to $NewFileName
$NewFileName = $FileName;
# Change file extension from csv to xls
$NewFileName =~ s/.csv$/.xls/;
if($debug){print("\$NewFileName = $NewFileName\n");}
