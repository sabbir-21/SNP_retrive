#!/usr/bin/perl -w
###########################################################
#
#Mupro: Protein Mutation Stability Prediction Program
#configure.pl: to configure the installation of Mupro 
#
#Author: Jianlin cheng
#Date: Sept. 14, 2005 
#Institute for Genomics and Bioinformatics
#School of Information and Computer Science
#University of California, Irvine
#
##########################################################

#######Customize settings here##############
#
#set installation directory of Mupro
$install_dir = "/home/nurul/Documents/win_apps/imutant/muPro1.1/mupro1.1";

######End of Customization##################


################Don't Change the code below##############
if (! -d $install_dir)
{
	die "can't find installation directory.\n";
}
if ( substr($install_dir, length($install_dir) - 1, 1) ne "/" )
{
	$install_dir .= "/"; 
}

-f "$install_dir/server/svm_classify" || die "Please copy svm_clasify of SVM-light to $install_dir/server.\n";

#check if the installation directory is right
#the configuration file must run in the installation directory
$cur_dir = `pwd`;  
chomp $cur_dir; 
$configure_file = "$cur_dir/configure.pl";
if (! -f $configure_file || $install_dir ne "$cur_dir/")
{
	die "Please check the installation directory setting and run the configure program there.\n";
}

$bin_dir = "${install_dir}bin/";
$model_dir = "${install_dir}model/";
$server_dir = "${install_dir}server/";
$script_dir = "${install_dir}script/";
$test_dir = "${install_dir}test/";

if ( ! -d $bin_dir || ! -d $model_dir 
   || ! -d $server_dir || ! -d $test_dir )
{
	die "some sub directories don't exist. check the installation tar ball.\n";
}

$svm_exe = "${server_dir}svm_classify.sh";

#generate svm classify script
open(SERVER_SH, ">$svm_exe") || die "can't write svm shell script.\n";
print SERVER_SH "#!/bin/sh\n"; 
print SERVER_SH "if [ \$# -ne 3 ]\n"; 
print SERVER_SH "then\n";
print SERVER_SH "\techo \"need three parameters: data set, model, output file.\"\n";
print SERVER_SH "\texit 1\n";
print SERVER_SH "fi\n"; 
print SERVER_SH "${server_dir}svm_classify \$1 \$2 \$3\n"; 
close SERVER_SH; 


$mupro_class_sh = "${bin_dir}predict_class.sh";
$class_model = "${model_dir}class/model_class_final";
#$class_model = "${model_dir}class/model_class_big";
print "generate mupro classification script using SVM only...\n";
open(SERVER_SH, ">$mupro_class_sh") || die "can't write mupro shell script.\n";
print SERVER_SH "#!/bin/sh\n#predict mutation stability for one mutation.\n";
print SERVER_SH "if [ \$# -ne 1 ]\n";
print SERVER_SH "then\n\techo \"need mutation input file.\"\n\texit 1\nfi\n";
print SERVER_SH "${script_dir}predict_mut.pl $svm_exe $class_model \$1 0 \n"; 
close SERVER_SH;

$mupro_small_sh = "${bin_dir}predict_class_small.sh";
$class_model = "${model_dir}svm_seq/model";
$svm_exe5 = "${server_dir}svm_classify5";
print "generate mupro classification script (small window) using SVM only...\n";
open(SERVER_SH, ">$mupro_small_sh") || die "can't write mupro shell script.\n";
print SERVER_SH "#!/bin/sh\n#predict mutation stability for one mutation.\n";
print SERVER_SH "if [ \$# -ne 1 ]\n";
print SERVER_SH "then\n\techo \"need mutation input file.\"\n\texit 1\nfi\n";
print SERVER_SH "${script_dir}predict_mut_small.pl $svm_exe5 $class_model \$1 0 \n"; 
close SERVER_SH;

$mupro_regr_sh = "${bin_dir}predict_regr.sh";
$regr_model = "${model_dir}regression/model_regr_final";
print "generate mupro regression script using SVM only...\n";
open(SERVER_SH, ">$mupro_regr_sh") || die "can't write mupro shell script.\n";
print SERVER_SH "#!/bin/sh\n#predict mutation stability for one mutation.\n";
print SERVER_SH "if [ \$# -ne 1 ]\n";
print SERVER_SH "then\n\techo \"need mutation input file.\"\n\texit 1\nfi\n";
print SERVER_SH "${script_dir}predict_mut.pl $svm_exe $regr_model \$1 1 \n"; 
close SERVER_SH;

$mupro_regr_all = "${bin_dir}predict_regr_all.sh";
$regr_model = "${model_dir}regression/model_regr_final";
print "generate mupro regression script using SVM only...\n";
open(SERVER_SH, ">$mupro_regr_all") || die "can't write mupro shell script.\n";
print SERVER_SH "#!/bin/sh\n#predict mutation stability for one mutation or all 19 possible mutations.\n";
print SERVER_SH "if [ \$# -ne 1 ]\n";
print SERVER_SH "then\n\techo \"need mutation input file.\"\n\texit 1\nfi\n";
print SERVER_SH "${script_dir}predict_mut_all.pl $svm_exe $regr_model \$1 1 \n"; 
close SERVER_SH;


`chmod 755 $mupro_class_sh $mupro_small_sh $mupro_regr_sh`; 
`chmod 755 ${script_dir}*.pl`; 
`chmod 755 ${server_dir}*.sh`; 
`chmod 755 ${bin_dir}*.sh`; 


