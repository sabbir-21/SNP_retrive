#!/bin/sh
#predict mutation stability for one mutation.
if [ $# -ne 1 ]
then
	echo "need mutation input file."
	exit 1
fi
/usr/local/httpd/htdocs/test/mupro1.1/script/predict_mut_small.pl /usr/local/httpd/htdocs/test/mupro1.1/server/svm_classify5 /usr/local/httpd/htdocs/test/mupro1.1/model/svm_seq/model $1 0 
