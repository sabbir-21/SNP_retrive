#!/usr/bin/perl -w

###########################################################################################################################
#Given a mutaion, predict how it changes the stability of protein structure
#Input: svm_seq predictor, model, input file, method(classification or regression).
#Input file format: name, sequence, pos, original aa, substibute aa
#Author: Jianlinc Cheng
#Start Date: 9/14/2005
###########################################################################################################################

$amino_acids = "ACDEFGHIKLMNPQRSTVWY";

if (@ARGV != 4)
{
	die "need 4 params:svm_predictor, svm_model_seq,  input mutation file, method(1:classification, 2:regression)\n"; 
}
$svm_pre = shift @ARGV;
$svm_seq_model = shift @ARGV; 
$input = shift @ARGV;
$method = shift @ARGV; 

-f $svm_pre || die "can't find svm seq predictor.\n"; 
-f $svm_seq_model || die "can't find svm seq model.\n"; 


open(INPUT, "$input") || die "can't open input file.\n";


$name = <INPUT>; 
chomp $name;
$seq = <INPUT>;
chomp $seq;
$pos = <INPUT>;
chomp $pos;
$org_aa = <INPUT>;
chomp $org_aa;
$rep_aa = <INPUT>; 
chomp $rep_aa;
close INPUT; 

$error = ""; 

if ($seq eq "")
{
	$error = "sequence is empty, can't make prediction."; 
}

if ($pos < 0 || $pos > length($seq))
{
	$error = "mutation position is out of boundary."; 
}

if (substr($seq, $pos-1, 1) ne $org_aa)
{
	$error = "the amino acid at position: $pos is not $org_aa"; 
}
if (index($amino_acids, $org_aa) < 0)
{
	$error = "the original amino acid is not standard amino acid.";
}
if (index($amino_acids, $rep_aa) < 0)
{
	$error = "the substitute amino acid is not standard amino acid.";
}

if ($error ne "")
{
	print "Input error: $error\n"; 
	die; 
}

open(TMP_SEQ, ">$input.svm.seq") || die "can't create tmp file.\n"; 
print TMP_SEQ "0"; 
@vec = ();
for ($i = 0; $i < 20; $i++)
{
	$vec[$i] = 0; 
}
$org_i = index($amino_acids, $org_aa);
$rep_i = index($amino_acids, $rep_aa); 
$vec[$org_i] = -1;
$vec[$rep_i] = 1; 
for ($i = 0; $i < 20; $i++)
{
	print TMP_SEQ " ", $i+1, ":", $vec[$i]; 
}

$start = 21;
$win_extent = 3; 
@windows = (); 
for ($i = 0; $i < 2 * $win_extent * 20; $i++)
{
	$windows[$i] = 0; 
}
$relative = 0; 
for ($i = -$win_extent; $i <= $win_extent; $i++)
{
	if ($i == 0)
	{
		next; 
	}
	$cur_idx = $pos + $i - 1; 
	if ($cur_idx >= 0 && $cur_idx < length($seq))
	{
		$amino = substr($seq, $cur_idx, 1); 
		$ord = index($amino_acids, $amino);
		if ($ord >= 0)
		{
			$windows[$relative*20 + $ord] = 1; 
		}
	}
	$relative++; 
}
for ($i = 0; $i < 2 * $win_extent * 20; $i++)
{
	print TMP_SEQ " ", $start+$i, ":", $windows[$i]; 
}
print TMP_SEQ "\n"; 
close TMP_SEQ; 


system("$svm_pre $input.svm.seq $svm_seq_model $input.svm.pre >/dev/null");

open(RES, "$input.svm.pre") || die "can't open svm seq prediction file\n"; 
$res = <RES>; 
chomp $res; 
$score = $res; 
if ($score > 0)
{
	print "The mutation INCREASE the stability of the protein.\n";
	if ($method == 0)
	{
		print "Confidence score: $score (bigger -> more confident)\n";
	}
	else
	{
		print "Energy change (delta G) = $score\n";
	}
}
else
{
	print "The mutation DECREASE the stability of the protein.\n";
	if ($method == 0)
	{
		print "Confidence score: $score (smaller -> more confident)\n";
	}
	else
	{
		print "Energy change (delta G) = $score\n";
	}
}
close RES; 

`rm $input.svm.seq`; 
`rm $input.svm.pre`; 


