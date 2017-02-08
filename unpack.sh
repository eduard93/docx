#!/bin/bash

rm -rf $2
unzip -qq $1 -d $2
OLDDIR=$(pwd)
cd $2
IFS=$'\n'
for f in $(find -name '*.xml' -or -name '*.rels')
do
    xmllint --format -o $f $f
done
cd $OLDDIR