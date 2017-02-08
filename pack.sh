#!/bin/bash

rm -f $2
OLDDIR=$(pwd)
cd $1
zip -r -qq $OLDDIR/$2 ./
cd $OLDDIR