#!/usr/bin/env perl 
#===============================================================================
#
#         FILE: progressbar.pl
#
#        USAGE: ./progressbar.pl  
#
#  DESCRIPTION: From https://code.google.com/p/bsdz-perl/source/browse/trunk/win32-gui-xmlbuilder/Win32-GUI-665-samples/progressbar.pl?r=2 
#
#      OPTIONS: ---
# REQUIREMENTS: ---
#         BUGS: ---
#        NOTES: ---
#       AUTHOR: Scott Webster (sdw), ScottW@ChefCentral.com
# ORGANIZATION: Chef Central, Inc.
#      VERSION: 1.0
#      CREATED: 9/1/2014 3:05:20 PM
#     REVISION: ---
#===============================================================================

#use strict;
use warnings;
use utf8;

use Win32::GUI();
use Win32::Sound();
my ($W, $F, $tW, $tX, $tY, $sX1, $sX2, $sX3, $sX4);

$W = new Win32::GUI::Window(
    -title    => "Win32::GUI::ProgressBar test",
    -left     => 100,
    -top      => 100,
    -width    => 400,
    -height   => 150,
    -font     => $F,
    -name     => "Window",
) or print_and_die("new Window");

$tX = 5; # Left edge of first label
$W->AddLabel(
    -name => "MIN_L",
    -text => "Min.:",
    -left => $tX,
    -top  => 5,
);
$tX += $W->MIN_L->Width + 10;
$sX1 = $tX; # Left edge of first text field
$W->AddTextfield(
    -name   => "MIN",
    -left   => $tX,
    -top    => 5,
    -width  => 80,
    -height => 20,
    -text   => "0",
);
$tX += $W->MIN->Width + 10;
$sX2 = $tX; # Right edge of second label
$W->AddLabel(
    -name => "MAX_L",
    -text => "Max.:",
    -left => $tX,
    -top  => 5,
);
$tX += $W->MAX_L->Width + 10;
$sX3 = $tX; # Left edge of second text field
$W->AddTextfield(
    -name   => "MAX",
    -left   => $tX,
    -top    => 5,
    -width  => 80,
    -height => 20,
    -text   => "100",
);
$tX += $W->MAX->Width + 10;
$sX4 = $tX; # Left edge of button
$W->AddButton(
    -name => "SET",
    -text => "Set range",
    -left => $tX,
    -top  => 5,
);
$tY = 5 + $W->MAX->Height * 2; # Top edge of second row
$tX = 5;
$W->AddLabel(
    -name => "ACT_L",
    -text => "Actual:",
    -left => $tX,
    -top  => $tY,
);
$tX = $sX1;
$W->AddTextfield(-name => "ACT",
    -text   => "10",
    -left   => $tX,
    -top    => $tY,
    -width  => 80,
    -height => 20,
);
$tX = $sX2;
$W->AddLabel(
    -name => "INC_L",
    -text => "Inc.:",
    -left => $tX,
    -top  => $tY,
);
$tX = $sX3;
$W->AddTextfield(
    -name   => "INC",
    -text   => "10",
    -left   => $tX,
    -top    => $tY,
    -width  => 80,
    -height => 20,
);
$tX = $sX4;
$W->AddButton(
    -name => "POS",
    -text => "Set position",
    -left => $tX,
    -top  => $tY,
);
$tX += $W->POS->Width + 10;
$W->AddButton(
    -name => "UP",
    -text => "+",
    -left => $tX,
    -top  => $tY,
);
$tX += $W->UP->Width + 10;
$W->AddButton(
    -name => "DN",
    -text => "-",
    -left => $tX,
    -top  => $tY,
);

$tX += $W->DN->Width + 10;

$W->Resize($tX, $W->Height);

$tY = 5 + $W->MAX->Height * 2 + $W->ACT->Height * 2;
$tW = $W->Width / 2;

$W->AddProgressBar(
    -name   => "PB",
    -left   => ($W->Width - $tW) / 2,
    -top    => $tY,
    -width  => $tW,
    -height => 20,
    -smooth => 1,
);

$W->PB->SetPos(10);

$W->Show;

$return = $W->Dialog();
print "Dialog: $return\n";

sub POS_Click {
    $W->PB->SetStep($W->INC->Text);
    $W->PB->SetPos($W->ACT->Text);
}

sub UP_Click {
    if($W->ACT->Text == $W->MAX->Text) {
        Win32::Sound::Play("SystemDefault", SND_ASYNC);
        return 1;
    } else {
        my $pos = $W->PB->SetPos(0);
        $W->PB->SetPos($pos);
        my $new = $pos + $W->INC->Text;
        $new = $W->MAX->Text if $new > $W->MAX->Text;
        $W->PB->SetPos($new);
                $W->ACT->{-text} = $new;
    }
}

sub DN_Click {
    if($W->ACT->Text == $W->MIN->Text) {
        Win32::Sound::Play("SystemDefault", SND_ASYNC);
        return 1;
    } else {
        my $pos = $W->PB->SetPos(0);
        $W->PB->SetPos($pos);
        my $new = $pos - $W->INC->Text;
        $new = $W->MIN->Text if $new < $W->MIN->Text;
        $W->PB->SetPos($new);
        $W->ACT->{-text} = $new;
    }
}


sub SET_Click {
    $W->PB->SetRange($W->MIN->Text, $W->MAX->Text);
}

sub print_and_die {
    my($text) = @_;
    my $err = Win32::GetLastError();
    die "$text: Error $err\n";
}

sub Window_Resize {
    my $tW = $W->ScaleWidth / 2;
    my $tY = ($W->MAX->Height * 2) + ($W->ACT->Height * 2) + 5;
    if($W->PB) {
        $W->PB->Move(($W->ScaleWidth-$tW)/2, $tY);
        $W->PB->Resize($tW, $W->ScaleHeight-$tY-20);
    }
}
