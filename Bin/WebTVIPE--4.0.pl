#!/usr/bin/perl

#################################
#   WebTV IPE (In-place Edit)   #
#                               #
# By: Eric MacDonald            #
# Date: November 11, 2004       #
#                               #
# This is a patcher tool        #
# for any SuperViewer template  #
#################################

use Digest::MD5 qw(md5_hex);

my $CONFIGDIR = "../Config";
my $CONFIGFILE = "Config.ini";

&main;
exit;

sub addHObjs ($) {
	open(HEADER,"< $CONFIGDIR/".shift) or die "Can't open template file";
	
	
	while(<HEADER>){
		next if (/^\#/ or /^$/);
		
		/^(\S*)\:\s*\"(.*?)\"\=?(.*)$/;
		
		$examps = $3;
		
		$hderObjs{$1}[0] = $2;
		$hderObjs{$1}[1] = [split(/\,/,$examps)] if ($examps);
	
	}
	
	close(TEMP);
	
	1;
}

sub readPresets ($)
{

	open(TEMP,"< $CONFIGDIR/".shift) or die "Can't open template file";
	while(<TEMP>){
		next if(/$\#/ or /^$/);
		
		# The extra "fluff" may be useful in later versions.
	
		if(/^\s*\}/){
			
			--$notlevel1;
			$notlevel1 = 0 if($notlevel1 < 0);
			pop(@inlevel);
	
		}elsif(/^(\S*)\s*(\S*)\:\s*(.*)$/){
			
			$variab  = $1;
			$subvar  = $2;
			$flatval = $3;
		
			if($notlevel1){
	
				if($inlevel[$notlevel1 - 1] eq "define" && $notlevel1 == 1){
	
					if($variab eq "description" && $flatval=~/\"(.*?)\"/){
	
						$boxPresets{$defining}{'description'} = $1;
	
					}else{
	
						$boxPresets{$defining}{'header'}{$variab} = $flatval;
	
					}
			}
	
		}else{
	
			if($variab eq "define" && ($flatval=~/^\{/)){
				
				$notlevel1 = 1;
				$inlevel[$notlevel1 - 1] = $variab;
				$defining = $subvar;
			
			}
	
		}
	
	}



}
	
	close(TEMP);
}


sub readTemp ($) {

open(TEMP,"< $CONFIGDIR/".shift) or die "Can't open template file";
while(<TEMP>){
next if(/$\#/ or /^$/);

if(/^\-\>(.*)$/){

$temptitle = $1;

}elsif(/^\s*\}/){
	
--$notlevel1;

$notlevel1 = 0 if($notlevel1 < 0);

pop(@inlevel);

}elsif(/^(\S*)\s*(\S*)\:\s*(.*)$/){

$variab  = $1;
$subvar  = $2;
$flatval = $3;


if($notlevel1){

if($inlevel[$notlevel1 - 1] eq "define" && $notlevel1 == 1){

if($variab eq "block-offset"&& ($flatval=~/^[A-Fa-f0-9]*$/) ){

$blockVars{$defining}{'block-offset'} = $flatval;
}elsif($variab eq "block-size" && ($flatval=~/^\d*$/)){

$blockVars{$defining}{'block-size'} = $flatval;

}elsif($variab eq "description" && ($flatval=~/^\"(.*?)\"$/)){

$blockVars{$defining}{'description'} = $1;

}elsif($variab eq "multiple"){

$blockVars{$defining}{'multiple'} = $flatval;

}elsif($variab eq "views"){

$blockVars{$defining}{'views'} = $flatval;

}elsif($variab eq "noblanks"){

$blockVars{$defining}{'noblanks'} = $flatval;

}elsif($variab eq "forward" && $flatval=~/^\"(.*)\"$/){

push(@{$blockVars{$defining}{'forward'}{'names'}},$1);
$blockVars{$defining}{'forward'}{'packed'} .= $subvar;

}elsif($variab eq "isspecial"){

$blockVars{$defining}{'isspecial'} = $flatval;

}elsif($variab eq "write-end"){

$blockVars{$defining}{'write-end'} = $flatval;

}elsif($variab eq "headers"){
$notlevel1++;
$inlevel[$notlevel1 - 1] = "headers";
}

}elsif($inlevel[$notlevel1 - 1] eq "headers" && $notlevel1 == 2){
push(@{$blockVars{$defining}{'headers'}},$variab) if($variab);
}
}else{

if($variab eq "hex-code" && ($subvar=~/^[A-Fa-f0-9]*$/) && !($editCodes{$subvar})){

foreach $hexC (split(/(..)/,$flatval)) {
next if(!($hexC =~ /^[A-Fa-f0-9][A-Fa-f0-9]$/));

$editCodes{$subvar} .= chr(eval("0x$hexC"));

}

}elsif($variab eq "file-code" && ($subvar=~/^[A-Fa-f0-9]*$/) && !($editCodes{$subvar})){
open(FILE,"< $CONFIGDIR/$flatval") or die "Couldn't open file to get code.";
binmode(FILE);
sysread(FILE,$editCodes{$subvar},-s FILE);
close(FILE);
}elsif($variab eq "string-code" && ($subvar=~/^[A-Fa-f0-9]*$/) && ($flatval=~/^\"(.*?)\"$/) && !($editCodes{$subvar})){
$flatval = $1;

$editCodes{$subvar} = $flatval . "\x00";

}elsif($variab eq "path" && $flatval=~/^\"(.*)\"$/){

$pathto = $1;

}elsif($variab eq "vers" && $flatval=~/^\"(.*)\"$/){

$viewervers = $1;

}elsif($variab eq "unedited-hash"){

$thehash = $flatval;

}elsif($variab eq "define" && ($flatval=~/^\{/)){

$notlevel1 = 1;
$inlevel[$notlevel1 - 1] = $variab;
$defining = $subvar;
}
}
}



}
close(TEMP);

}

sub editHeader {
system('cls');

open(SV,"+< $pathto") or die "Couldn't open SV executable.";
binmode(SV);
my @options = [];
my $count = 0;
%currentSV = ();
foreach $blockn (keys %blockVars) {
next if(!(defined($blockVars{$blockn}{'headers'})));
sysseek(SV,eval("0x" . $blockVars{$blockn}{'block-offset'}),0);
sysread(SV,$block,$blockVars{$blockn}{'block-size'});
$block = substr($block,0,index($block,"\x00"));
foreach $headO (split(/\x0D\x0A/,$block)) {
$headO=~/^(\S*?)\:\s*(.*)/;
$currentSV{$1} = $2;
}
}
close(SV);

print "Please pick an action:\n\n";


foreach $blockn (keys %blockVars) {
next if(!(defined($blockVars{$blockn}{'headers'})));
my $count2 = 0;

foreach $headers (@{$blockVars{$blockn}{'headers'}}) {
++$count;
push(@options,[$count2,$blockn]);
++$count2;

print "[$count] Edit $headers" .
	  (($currentSV{$headers}) ? " (" . $currentSV{$headers} . ")\n" : "\n");
}

}

++$count;
print "[$count] EXIT\n";

REP1:
print "\nChoose an action: "; chomp($choice = <STDIN>);
goto REP1 if(!($choice) || ($choice > $count) || ($choice < 1));
return if($count == $choice);

my $headcos = $blockVars{$options[$choice][1]}{'headers'}[$options[$choice][0]];


print  "\n--------------------\n" . 
	   $hderObjs{$headcos}[0] . "\n";

foreach $examh (@{$hderObjs{$headcos}[1]}) {
print "Example: $examh\n"
}
print "\nValue: "; chomp($ans = <STDIN>);

$currentSV{$headcos} = $ans;

writeHeader(\%currentSV);

}

sub writeHeader (%) {

%currentSV = (%{$_[0]});


open(SV,"+< $pathto");
foreach $blockn (keys %blockVars) {
next if(!(defined($blockVars{$blockn}{'headers'})));
$block = "";

foreach $headers (@{$blockVars{$blockn}{'headers'}}) {

$block .= "$headers: $currentSV{$headers}\x0D\x0A" if($currentSV{$headers});


}
$block .= $blockn if($blockVars{$blockn}{'write-end'});


$block = pack("a" . $blockVars{$blockn}{'block-size'},$block);
sysseek(SV,eval("0x" . $blockVars{$blockn}{'block-offset'}),0);
syswrite(SV,$block,$blockVars{$blockn}{'block-size'});


}
close(SV);


}

sub loadConfig($) {

open(CONFIG,"< $CONFIGDIR/".shift) or die "Can't open config file";
while (<CONFIG>){

next if (/^\#/ or /^$/);

/^(\S*)\s*\=\s*(\"|\'|)(.*?)(\"|\'|)$/;

$variable = $1;
$value    = $3;

if($variable eq "template"){

readTemp($value);

}elsif($variable eq "headers"){

addHObjs($value);


}elsif($variable eq "presets"){

readPresets($value);

}

}
close(FILE);

1;
}

sub printSVHOs {
system("cls");

open(SV,"< $pathto") or die "Couldn't open SV executable.";
binmode(SV);

$blockstr = "";
foreach $blockn (keys %blockVars) {
next if(!(defined($blockVars{$blockn}{'headers'})));

sysseek(SV,eval("0x" . $blockVars{$blockn}{'block-offset'}),0);
sysread(SV,$block,$blockVars{$blockn}{'block-size'});
$block = substr($block,0,index($block,"\x00"));
$block .= ": VALUE\n" if(rindex($block,$blockn) == (length($block) - length($blockn)));
$blockstr .= $block;
}

close(SV);
print "Here's the header:\n\n" .
      "$blockstr\n\n";

print "Press <ENTER> to return.\n"; <STDIN>;
}

sub printHOs {
system("cls");

foreach  (keys %hderObjs) {

print "[Header]: $_\n" .
	  "[Description]: " . ($hderObjs{$_}[0]) . " \n\n";


}

print "Press <ENTER> to return.\n"; <STDIN>;

}
sub modifyToSpecs ($){

my $checkme = shift;

foreach $subvar (keys %editCodes) {
my $subsit  = $editCodes{$subvar};
my $lenofsv = length($subsit);
$subvar = eval("0x$subvar");
$checkme = (substr($checkme,0,($subvar)) . ($subsit) . substr($checkme,(($subvar + $lenofsv))));
}

return $checkme;
}

sub startWizzard {
system("cls");

print <<ERIC;
The Super WebTV Viewer is a creation of Eric MacDonald.  It is different than a regular viewer in the sence that it is more relitive to the client protocol of the ever-changing WebTV Protocol (WTVP).  What this means is it allows people who have a WebTV Viewer that isn't capabile of connecting to newer WTVP servers to do so.  This is only important to WebTV hackers interested in connecting to WebTV services on their computer, and of course, on the defensive side- it's important to WebTV employees.  For this reason, it's important to keep this tool and the use of the SV secret from WebTV empoyees so they may not find a way to make this tool useless.

The first step to creating a SV is to have a generic viewer (one that was or is available at WebTV's developer website).

This tool parses a template to discover how to create and modify a SV.  The viewer that this tool knows how to edit is the "$viewervers" and that viewer executable should be located at "$pathto".  If your computer does not fit these credentials then please make adjustments before continuing.
	
If you are ready, press <ENTER> and this tool will patch this viewer with the correct instructions to operate as a SV.
ERIC
	;
<STDIN>;

if(buildSV(1)){
system("cls");

print <<ERIC;
At this point in time you're generic viewer is not set to reap a header.  A header is the section of the WTVP that educates the WTVP sever of how it should handle a request.  The header is the most important and frequently changed part of the WTVP and it's important to setup one on the SV.  Each individual box has a different header and thoes difference are most reflective on the box type, the build on the box and the Silocon Serial Identification number.  Before you continue, I urge you to gather information from a box's 411 (info) page or to packet sniff the box's outgoing messages to see the header in full form.

When you are ready please press <ENTER> and this tool will send you to a menu in which you would have to select your box type and then enter specific data to give the SV more of a name.
ERIC
	;
<STDIN>;

if(&presetHeader){
system("cls");
print "Success!  What needs to be done now is for you to open your new viewer and connect to a server.  Press <ENTER> to return to the main menu.\n";
<STDIN>;
}else{
system("cls");
print "It seams as if this tool can't create a header.  Opps.  Thank you, come again.  Press <ENTER> to return to the main menu.\n";
<STDIN>;
}

}else{
system("cls");
print "It seams as if this tool can't create a SV.  Opps.  Thank you, come again.  Press <ENTER> to return to the main menu.\n";
<STDIN>
}

}

sub presetHeader {
system('cls');
print "Choose a preset\n\n";

my $count = 0;
my @options = [];

foreach  (keys %boxPresets) {
++$count;

print "[$count] $boxPresets{$_}{'description'}\n";
$options[$count] = $_;

}
++$count;
print "[$count] EXIT\n\n";
REP1:
print "Choice: "; chomp($choice = <STDIN>);

goto REP1 if(!($choice) || ($choice > $count) || ($choice < 1));
return if($count == $choice);

foreach $headcos (keys %{$boxPresets{$options[$choice]}{'header'}}) {

my $ans = $boxPresets{$options[$choice]}{'header'}{$headcos};

if($ans eq "=GET(?)"){

print  "\n--------------------\n" . 
	   $hderObjs{$headcos}[0] . "\n";

foreach $examh (@{$hderObjs{$headcos}[1]}) {
print "Example: $examh\n"
}
print "\nValue: "; chomp($ans = <STDIN>);

$currentSV{$headcos} = $ans;
print "\n";
}else{
$currentSV{$headcos} = $ans;
}


}

writeHeader(\%currentSV);


}

sub buildSV ($) {
system("cls");

my $uans = shift;

print "Are you sure you want me to write to the file '$pathto' [Y/N]?: "; chomp($ans = <STDIN>);

return if(!($ans=~/^Y/i));

open(SV,"< $pathto") or die "Couldn't open SV executable.";
binmode(SV);
sysread(SV,$checkme,-s SV,0);
if($uans == 1){
$hash = md5_hex($checkme);
if($thehash ne $hash){
print "\n\nThe file that I've opened is not unedited.\n\n" .
      "ERROR: $hash ne $thehash\nFILELEN: ". length($checkme) ."\n\n";
print "Press <ENTER> to return.\n"; <STDIN>;
return;
}
}
close(SV);


$checkme = modifyToSpecs($checkme);

open(SV,"> $pathto") or die "Couldn't open SV executable.";
binmode(SV);
print SV $checkme;
close(SV);

if($uans == 1){
print "\n\nThe WebTV Viewer is now edited.  There are no headers set, however, so please set them.\n";
print "Press <ENTER> to return.\n"; <STDIN>;
}

return 1;
}

sub editSpecial {
system('cls');
my $count = 0;

print "Choose a block to edit or view:\n\n";

foreach  (keys %blockVars) { 
if($blockVars{$_}{'isspecial'} eq "yes"){
++$count;
print "[$count] View item(s) in the block named \"$blockVars{$_}{'description'}\"\n";
$options[$count] = $_;
}
}
++$count;
print "[$count] EXIT\n\n";
REP1:
print "Choice: "; chomp($choice = <STDIN>);

goto REP1 if(!($choice) || ($choice > $count) || ($choice < 1));
return if($count == $choice);


my $block = $options[$choice];
my $buf = "";
open(SV,"+< $pathto") or die "Couldn't open SV executable.";
binmode(SV);

sysseek(SV,eval('0x' . $blockVars{$block}{'block-offset'}),0);
sysread(SV,$buf,$blockVars{$block}{'block-size'});

$tempbool = ($blockVars{$block}{'multiple'} eq "no" or $buf eq "");

$blockVars{$block}{'views'} = 2 if($tempbool);

BAK:
system('cls');
print "$blockVars{$block}{'description'}:\n";

my $tempviews = "";
my $count2 = 1; 
for($i = 0; $blockVars{$block}{'views'} > $i; $i++){
my $count = 0; 
my @array1 = unpack($blockVars{$block}{'forward'}{'packed'},$buf);
$length = length(pack($blockVars{$block}{'forward'}{'packed'},@array1));
$buf = substr($buf,$length);
print "\n" if(!($tempviews) && $#array1 > 1 || $i == 0);

foreach  (@array1) {
if($blockVars{$block}{'noblanks'} eq "yes" and $_ eq ""){
$tempviews = $blockVars{$block}{'views'} if(!($tempviews));
++$blockVars{$block}{'views'};
next;
}

print "[$count2] [" . $blockVars{$block}{'forward'}{'names'}[$count] . "]: $_\n";

if(($tempviews && ($count2 == $tempviews))){
$blockVars{$block}{'views'} = $tempviews;
}

++$count2;
++$count;
}

last if($buf eq "");
}
$tempbool = ($blockVars{$block}{'multiple'} eq "no" or $buf eq "");

print "\nIf you wish to exit press";
print " 1 then <ENTER>\nTo view the next $blockVars{$block}{'views'} entries press" if(!($tempbool));
print " <ENTER>.\n\nChoice: "; chomp($ans = <STDIN>);

goto BAK if($ans eq "" && !($tempbool));

close(SV);

}

sub main {
system("cls");
reset 'a-z';

loadConfig($CONFIGFILE);

  while ($#ARGV >= 0) {
    $_ = shift @ARGV;
    if (m/--?f/) { $nofrontend = 1; }
    elsif (m/--?i/) { $dataout = 1; }
  }

if(!$nofrontend){
print "The Super WebTV Viewer Pather (In-Place Editor) Tool 3.0\n" .
	  "Copyright (c) 1999-2004 Eric MacDonald\n\n" .
	  "Template: \"$temptitle\"\n\n";

my @options = (
               ["Show allowable headers",\&printHOs],
               ["Show the SV's header",\&printSVHOs],
               ["Build SV's header.",\&editHeader],
               ["Build SV's header using preset.",\&presetHeader],
               ["Build a new base SV.",\&buildSV,1],
               ["Rebuild base SV (don't destroy header).",\&buildSV,2],
               ["New SV Wizard",\&startWizzard]
               );

foreach  (keys %blockVars) { if($blockVars{$_}{'isspecial'} eq "yes"){
push(@options,["View special viewer data.",\&editSpecial]);
last;
} }


open(SV,"< $pathto") or die "Couldn't open SV executable.";
binmode(SV);
sysread(SV,$checkme,-s SV,0);
$hash = md5_hex($checkme);
close(SV);
print "*$hash*";

my $count = 0;
for($count = 1; $count < ($#options + 2); ++$count){
print "[$count] ". $options[$count - 1][0] ."\n";
}
print "[$count] EXIT\n\n";

REP2:
print "Choose an action: "; chomp($choice = <STDIN>);
goto REP2 if(!($choice) || ($choice > $count) || ($choice < 1));

return 1 if($choice eq $count);

--$choice;
$options[$choice][1]->($options[$choice][2]);

&main;
}else{





}

}