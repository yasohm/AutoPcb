CC = gcc
CFLAGS = -Wall -Wextra -std=c99
LIBS = -lxlsxio_read -lxlsxwriter -lz -llzma -lbz2 -lzstd


modif: modif.c
	$(CC) $(CFLAGS) -o modif modif.c $(LIBS)


main: main.c modif.c
	$(CC) $(CFLAGS) -o main main.c


db-import:
	python3 import_xlsx_to_sqlite.py


db-export:
	python3 export_sqlite_to_xlsx.py

clean:
	rm -f modif main *.o


install-deps:
	sudo apt-get update
	sudo apt-get install -y libxlsxio-dev libxlsxwriter-dev

run-main: main
	./main

.PHONY: clean install-deps db-import db-export run-main