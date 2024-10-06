#!/bin/bash

gcc -o txtToExcel src/txtToExcel.c -L/usr/local/lib64 -lxlsxwriter -lz $(pkg-config --cflags --libs gtk+-3.0)

sleep 1;

./txtToExcel