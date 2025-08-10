#!/bin/sh
#predict mutation stability for one mutation or all 19 possible mutations.
if [ $# -ne 1 ]
then
	echo "need mutation input file."
	exit 1
fi
/usr/local/httpd/htdocs/test/mupro1.1/script/predict_mut_all.pl /usr/local/httpd/htdocs/test/mupro1.1/server/svm_classify.sh /usr/local/httpd/htdocs/test/mupro1.1/model/regression/model_regr_final $1 1 
