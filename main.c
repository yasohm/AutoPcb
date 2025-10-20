#include <stdio.h>
#include <stdlib.h>

int main(void) {
    int rc;

    printf("Exporting tables to input/*.xlsx...\n");
    rc = system("python3 export_sqlite_to_xlsx.py");
    if (rc != 0) {
        fprintf(stderr, "export_sqlite_to_xlsx.py failed\n");
        return 1;
    }

    printf("Running modif...\n");
    rc = system("./modif");
    if (rc != 0) {
        fprintf(stderr, "modif failed\n");
        return 1;
    }

    printf("Done.\n");
    return 0;
}