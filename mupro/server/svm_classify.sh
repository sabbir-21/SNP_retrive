#!/bin/sh
if [ $# -ne 3 ]
then
	echo "need three parameters: data set, model, output file."
	exit 1
fi
/usr/local/httpd/htdocs/test/mupro1.1/server/svm_classify $1 $2 $3
