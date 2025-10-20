#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <xlsxio_read.h>
#include <xlsxwriter.h>

// Provide a portable strdup for C99 without POSIX prototype
static char* safe_strdup(const char* source) {
    if (source == NULL) {
        return NULL;
    }
    size_t length = strlen(source) + 1;
    char* copy = (char*)malloc(length);
    if (copy == NULL) {
        return NULL;
    }
    memcpy(copy, source, length);
    return copy;
}

#define strdup safe_strdup

/*
 * This program extracts specific columns from PCB.xlsx and creates Excel formulas
 * for FB and MAX columns instead of pre-calculating values.
 * 
 * FB column: Uses VLOOKUP to get values from FB.xlsx based on WIDF (REF) and current week
 * WLOM column: Uses lookup to get values from ABC.xlsx WKQCO column based on WKIDF matching WIDF
 * MAX column: Uses Excel's MAX function to compare FB and WCMJ values
*/


const char* wanted_cols[] = {
    "WSTB","WIDF","WFOR","WGES","WPIV","WDES","WCOF","WLOM",
    "WCMJ","WSTKG","FB","MAX"
};
#define NUM_COLS (sizeof(wanted_cols)/sizeof(wanted_cols[0]))

int col_indices[NUM_COLS]; // Store matching indices

// Function to get current week number

int get_current_week() {
    time_t now = time(NULL);
    struct tm *tm_info = localtime(&now);
    
    // Calculate week number (ISO 8601 week)
    int day_of_year = tm_info->tm_yday + 1;
    
    // Simple week calculation (you might want to use ISO 8601 for more accuracy)
    int week = (day_of_year + 6) / 7;
    
    // Return calculated week - will be automatically detected in FB.xlsx
    return week;
}


// Global flag to show warning only once
static int fb_warning_shown = 0;
static int abc_warning_shown = 0;


// Simple hash table for FB data
#define HASH_SIZE 10000

typedef struct fb_hash_entry {
    char* ref;
    char* value;
    struct fb_hash_entry* next;
} fb_hash_entry;

static fb_hash_entry* fb_hash_table[HASH_SIZE] = {NULL};
static int fb_cache_loaded = 0;

// Simple hash table for ABC data
typedef struct abc_hash_entry {
    char* wkidf;
    char* wlom_value;
    struct abc_hash_entry* next;
} abc_hash_entry;

static abc_hash_entry* abc_hash_table[HASH_SIZE] = {NULL};
static int abc_cache_loaded = 0;

// Simple hash function
unsigned int hash_string(const char* str) {
    unsigned int hash = 5381;
    int c;
    while ((c = *str++)) {
        hash = ((hash << 5) + hash) + c;
    }
    return hash % HASH_SIZE;
}

// Function to load ABC data into hash table
int load_abc_hash_table() {
    const char* abc_file = "input/ABC.xlsx";
    xlsxioreader reader;
    xlsxioreadersheet sheet;
    
    if ((reader = xlsxioread_open(abc_file)) == NULL) {
        if (!abc_warning_shown) {
            printf("Warning: Could not open ABC.xlsx file\n");
            abc_warning_shown = 1;
        }
        return 0;
    }

int load_db(){
    const char* backup_FB = "FB.xlsx.backup";
    
}
    
    if ((sheet = xlsxioread_sheet_open(reader, NULL, XLSXIOREAD_SKIP_EMPTY_ROWS)) == NULL) {
        printf("Warning: Could not read ABC.xlsx sheet\n");
        xlsxioread_close(reader);
        return 0;
    }
    
    char* value;
    int wkidf_col = -1;
    int wlom_col = -1; // Will be set to WKQCO column
    
    // Read header row to find WKIDF and WKQCO columns
    if (xlsxioread_sheet_next_row(sheet)) {
        int col = 0;
        printf("ABC.xlsx header columns:\n");
        while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
            printf("  [%d] %s\n", col, value);
            if (strcmp(value, "WKIDF") == 0) {
                wkidf_col = col;
                printf("    -> Found WKIDF at column %d\n", col);
            } else if (strcmp(value, "WKQCO") == 0) {
                wlom_col = col;
                printf("    -> Found WKQCO at column %d (using for WLOM)\n", col);
            } else if (strcmp(value, "WLOM") == 0) {
                // Keep WLOM as fallback if WKQCO not found
                if (wlom_col == -1) {
                    wlom_col = col;
                    printf("    -> Found WLOM at column %d (fallback)\n", col);
                }
            }
            free(value);
            col++;
        }
        printf("ABC.xlsx has %d columns\n", col);
    }
    
    // Load data into hash table
    if (wkidf_col >= 0 && wlom_col >= 0) {
        int loaded_count = 0;
        while (xlsxioread_sheet_next_row(sheet)) {
            char* wkidf_val = NULL;
            char* wlom_val = NULL;
            int col = 0;
            
            while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
                if (col == wkidf_col) {
                    wkidf_val = strdup(value);
                } else if (col == wlom_col) {
                    wlom_val = strdup(value);
                }
                free(value);
                col++;
            }
            
            if (wkidf_val && wlom_val) {
                // Create hash entry
                abc_hash_entry* entry = malloc(sizeof(abc_hash_entry));
                if (entry) {
                    entry->wkidf = wkidf_val;
                    entry->wlom_value = wlom_val;
                    
                    // Insert into hash table
                    unsigned int hash = hash_string(wkidf_val);
                    entry->next = abc_hash_table[hash];
                    abc_hash_table[hash] = entry;
                    loaded_count++;
                } else {
                    free(wkidf_val);
                    free(wlom_val);
                }
            } else {
                if (wkidf_val) free(wkidf_val);
                if (wlom_val) free(wlom_val);
            }
        }
        
        abc_cache_loaded = 1;
        printf("Loaded %d ABC entries into hash table\n", loaded_count);
        
        // Show some sample WKIDF values for debugging
        if (loaded_count > 0) {
            printf("Sample WKIDF values from 'ABC.xlsx':\n");
            int sample_count = 0;
            for (int i = 0; i < HASH_SIZE && sample_count < 5; i++) {
                abc_hash_entry* entry = abc_hash_table[i];
                while (entry && sample_count < 5) {
                    printf("  %s\n", entry->wkidf);
                    sample_count++;
                    entry = entry->next;
                }
            }
        }
    }
    
    xlsxioread_sheet_close(sheet);
    xlsxioread_close(reader);
    return abc_cache_loaded;
}

// Function to get WLOM value from ABC hash table
char* get_wlom_value_by_widf(const char* widf_value) {
    // Load hash table if not loaded
    if (!abc_cache_loaded) {
        if (!load_abc_hash_table()) {
            return NULL;
        }
    }
    
    // Search in hash table
    unsigned int hash = hash_string(widf_value);
    abc_hash_entry* entry = abc_hash_table[hash];
    
    while (entry) {
        if (strcmp(entry->wkidf, widf_value) == 0) {
            return strdup(entry->wlom_value);
        }
        entry = entry->next;
    }
    
    return NULL;
}

// Function to clear ABC hash table (for reloading data)
void clear_abc_hash_table() {
    for (int i = 0; i < HASH_SIZE; i++) {
        abc_hash_entry* entry = abc_hash_table[i];
        while (entry) {
            abc_hash_entry* next = entry->next;
            free(entry->wkidf);
            free(entry->wlom_value);
            free(entry);
            entry = next;
        }
        abc_hash_table[i] = NULL;
    }
    abc_cache_loaded = 0;
}

// Function to load FB data into hash table
int load_fb_hash_table(int week) {
    const char* fb_file = "input/FB.xlsx";
    xlsxioreader reader;
    xlsxioreadersheet sheet;
    
    if ((reader = xlsxioread_open(fb_file)) == NULL) {
        if (!fb_warning_shown) {
            printf("Warning: Could not open FB.xlsx file\n");
            fb_warning_shown = 1;
        }
        return 0;
    }
    
    if ((sheet = xlsxioread_sheet_open(reader, NULL, XLSXIOREAD_SKIP_EMPTY_ROWS)) == NULL) {
        printf("Warning: Could not read FB.xlsx sheet\n");
        xlsxioread_close(reader);
        return 0;
    }
    
    char* value;
    int ref_col = 0; // Use row labels (column 0) as REF
    int week_col = -1;
    
    // Read header row to find REF column and week column
    if (xlsxioread_sheet_next_row(sheet)) {
        int col = 0;
        int first_available_week = -1;
        int first_available_week_col = -1;
        int target_week_found = 0;
        
        printf("FB.xlsx header columns:\n");
        while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
            printf("  [%d] %s\n", col, value);
            if (strcmp(value, "REF") == 0) {
                ref_col = col;
                printf("    -> Found REF at column %d\n", col);
            } else if (strcmp(value, "Étiquettes de lignes") == 0) {
                // Handle French column name - treat it as REF column
                ref_col = col;
                printf("    -> Found 'Étiquettes de lignes' at column %d (treating as REF)\n", col);
            } else {
                // Check if this is a week number
                int week_num = atoi(value);
                if (week_num >= 1 && week_num <= 52) {
                    // Store first available week for fallback
                    if (first_available_week == -1) {
                        first_available_week = week_num;
                        first_available_week_col = col;
                        printf("    -> First available week: %d at column %d\n", week_num, col);
                    }
                    
                    // Check if this is our target week
                    if (week_num == week) {
                        week_col = col;
                        target_week_found = 1;
                        printf("    -> Found target week %d at column %d\n", week_num, col);
                    }
                }
            }
            free(value);
            col++;
        }
        printf("FB.xlsx has %d columns\n", col);
        
        // If target week not found, use first available week
        if (!target_week_found && first_available_week != -1) {
            week_col = first_available_week_col;
            printf("    -> Target week %d not found, using first available week %d at column %d\n", 
                   week, first_available_week, week_col);
        }
    }
    
    // Load data into hash table
    if (ref_col >= 0 && week_col >= 0) {
        int loaded_count = 0;
        while (xlsxioread_sheet_next_row(sheet)) {
            char* ref_val = NULL;
            char* week_val = NULL;
            int col = 0;
            
            while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
                if (col == ref_col) {
                    ref_val = strdup(value);
                } else if (col == week_col) {
                    week_val = strdup(value);
                }
                free(value);
                col++;
            }
            
            if (ref_val && week_val) {
                // Create hash entry
                fb_hash_entry* entry = malloc(sizeof(fb_hash_entry));
                if (entry) {
                    entry->ref = ref_val;
                    entry->value = week_val;
                    
                    // Insert into hash table
                    unsigned int hash = hash_string(ref_val);
                    entry->next = fb_hash_table[hash];
                    fb_hash_table[hash] = entry;
                    loaded_count++;
                } else {
                    free(ref_val);
                    free(week_val);
                }
            } else {
                if (ref_val) free(ref_val);
                if (week_val) free(week_val);
            }
        }
        
        fb_cache_loaded = 1;
        printf("Loaded %d FB entries into hash table for week %d\n", loaded_count, week);
        
        // Show some sample REF values for debugging
        if (loaded_count > 0) {
            printf("Sample REF values from FB.xlsx:\n");
            int sample_count = 0;
            for (int i = 0; i < HASH_SIZE && sample_count < 5; i++) {
                fb_hash_entry* entry = fb_hash_table[i];
                while (entry && sample_count < 5) {
                    printf("  %s -> %s\n", entry->ref, entry->value);
                    sample_count++;
                    entry = entry->next;
                }
            }
        }
    }
    
    xlsxioread_sheet_close(sheet);
    xlsxioread_close(reader);
    return fb_cache_loaded;
}

// Function to get FB value from hash table
char* get_fb_value_by_widf(const char* widf_value, int week) {
    // Load hash table if not loaded
    if (!fb_cache_loaded) {
        if (!load_fb_hash_table(week)) {
            return NULL;
        }
    }
    
    // Search in hash table
    unsigned int hash = hash_string(widf_value);
    fb_hash_entry* entry = fb_hash_table[hash];
    
    while (entry) {
        if (strcmp(entry->ref, widf_value) == 0) {
            return strdup(entry->value);
        }
        entry = entry->next;
    }
    
    return NULL;
}

// Function to clear hash table (for reloading data)
void clear_fb_hash_table() {
    for (int i = 0; i < HASH_SIZE; i++) {
        fb_hash_entry* entry = fb_hash_table[i];
        while (entry) {
            fb_hash_entry* next = entry->next;
            free(entry->ref);
            free(entry->value);
            free(entry);
            entry = next;
        }
        fb_hash_table[i] = NULL;
    }
    fb_cache_loaded = 0;
}

// Function to get the maximum value between FB and WCMJ
char* get_max_value(const char* fb_value, const char* wcmj_value) {
    if (!fb_value && !wcmj_value) {
        return NULL;
    }
    
    if (!fb_value) {
        return strdup(wcmj_value);
    }
    
    if (!wcmj_value) {
        return strdup(fb_value);
    }
    
    // Convert strings to double for comparison
    double fb_num = atof(fb_value);
    double wcmj_num = atof(wcmj_value);
    
    if (fb_num >= wcmj_num) {
        return strdup(fb_value);
    } else {
        return strdup(wcmj_value);
    }
}

int find_column_index(const char* name, const char** header, int header_count) {
    for (int i = 0; i < header_count; i++) {
        if (strcmp(name, header[i]) == 0)
            return i;
    }
    return -1;
}

///////////////////////// the main function /////////////////////////

int main(int argc, char* argv[]) {
    const char* input_file = "PCB.xlsx";
    const char* output_file = "output.xlsx";
    
    // Check for reload flag
    int force_reload = 0;
    if (argc > 1 && strcmp(argv[1], "--reload") == 0) {
        force_reload = 1;
        printf("Force reload mode: will reload FB and ABC data\n");
    }
    
    // Check for file preprocessing flag
    int preprocess_files = 0;
    if (argc > 1 && strcmp(argv[1], "--preprocess") == 0) {
        preprocess_files = 1;
        printf("Preprocessing mode: will check and fix file issues\n");
    }
    
    // Get current week
    int current_week = get_current_week();
    printf("Current week: %d\n", current_week);
    
    // Check if required files exist
    printf("Checking required files...\n");
    
    // Check PCB file (try both extensions)
    FILE* pcb_test = fopen("input/PCB.xlsx", "r");
    if (pcb_test) {
        fclose(pcb_test);
        input_file = "input/PCB.xlsx";
        printf("Found PCB.xlsx in input folder\n");
    } else {
        pcb_test = fopen("input/PCB.xls", "r");
        if (pcb_test) {
            fclose(pcb_test);
            input_file = "input/PCB.xls";
            printf("Found PCB.xls in input folder\n");
        } else {
            printf("Error: PCB file not found (input/PCB.xls or input/PCB.xlsx)\n");
            printf("Please ensure PCB.xls or PCB.xlsx exists in the input folder\n");
            return 1;
        }
    }
    
    // Check ABC file
    FILE* abc_test = fopen("input/ABC.xlsx", "r");
    if (!abc_test) {
        printf("Error: ABC.xlsx file not found in input folder\n");
        printf("Please ensure ABC.xlsx exists in the input folder\n");
        return 1;
    }

    fclose(abc_test);
    printf("Found ABC.xlsx in input folder\n");
    
    // Check FB file
    FILE* fb_test = fopen("input/FB.xlsx", "r");
    if (!fb_test) {
        printf("Error: FB.xlsx file not found in input folder\n");
        printf("Please ensure FB.xlsx exists in the input folder\n");
        return 1;
    }
    fclose(fb_test);
    printf("Found FB.xlsx in input folder\n");

    // Open XLSX for reading
    xlsxioreader reader;
    printf("Opening %s file...\n", input_file);
    if ((reader = xlsxioread_open(input_file)) == NULL) {
        fprintf(stderr, "Error opening %s\n", input_file);
        return 1;
    }
    printf("%s opened successfully\n", input_file);


    // Prepare XLSX writer
    lxw_workbook  *workbook  = workbook_new(output_file);
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);


    // Open first sheet
    xlsxioreadersheet sheet;
    printf("Opening sheet...\n");
    if ((sheet = xlsxioread_sheet_open(reader, NULL, XLSXIOREAD_SKIP_EMPTY_ROWS)) == NULL) {
        fprintf(stderr, "Error reading sheet\n");
        return 1;
    }

    printf("Sheet opened successfully\n");

    char* value;
    int row = 0;

    // Read header row
    const char* header[500];
    int header_count = 0;

    printf("Attempting to read header row...\n");
    if (xlsxioread_sheet_next_row(sheet)) {
        printf("Header row found!\n");
        int col = 0;
        while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
            header[header_count++] = strdup(value);
            free(value);
            col++;
        }
        printf("Found %d columns in header\n", header_count);
        
        // Match indices
        printf("\nLooking for required columns:\n");
        for (size_t i = 0; i < NUM_COLS; i++) {
            col_indices[i] = find_column_index(wanted_cols[i], header, header_count);
            printf("Column '%s' found at index %d\n", wanted_cols[i], col_indices[i]);
        }
        
        printf("\nAvailable columns in PCB.xls:\n");
        for (int i = 0; i < header_count; i++) {
            printf("[%d] %s\n", i, header[i]);
        }

        // Write header to output
        for (size_t i = 0; i < NUM_COLS; i++)
            worksheet_write_string(worksheet, row, i, wanted_cols[i], NULL );
        // Extra columns
        worksheet_write_string(worksheet, row, NUM_COLS, "Inventaire", NULL);
        worksheet_write_string(worksheet, row, NUM_COLS + 1, "couv", NULL);
        row++;
        printf("Header written to output, starting data rows...\n");
    } else {
        printf("No header row found!\n");
    }

    // Read and write data rows
    int data_row_count = 0;
    while (xlsxioread_sheet_next_row(sheet)) {
        int col = 0;
        char* row_values[500] = {0};

        while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
            row_values[col++] = strdup(value);
            free(value);
        }
        
        data_row_count++;
        if (data_row_count <= 3) {
            printf("Processing row %d with %d columns\n", data_row_count, col);
        }


        
        for (size_t i = 0; i < NUM_COLS; i++) {
            if (strcmp(wanted_cols[i], "WLOM") == 0) {
                // Get WIDF value from the current row
                char* widf_value = NULL;
                if (col_indices[1] >= 0 && col_indices[1] < col && row_values[col_indices[1]]) {
                    widf_value = strdup(row_values[col_indices[1]]);
                }
                
                // Use C function to get WLOM value by checking WIDF against WKIDF in ABC.xlsx
                char* wlom_value = NULL;
                if (widf_value) {
                    // Force reload if requested
                    if (force_reload && abc_cache_loaded) {
                        clear_abc_hash_table();
                    }
                    wlom_value = get_wlom_value_by_widf(widf_value);
                    if (wlom_value) {
                        worksheet_write_string(worksheet, row, i, wlom_value, NULL);
                        printf("Found WLOM value for WIDF %s: %s\n", widf_value, wlom_value);
                        free(wlom_value);
                    } else {
                        printf("No WLOM value found for WIDF: %s\n", widf_value);
                    }
                    free(widf_value);
                }
                
            } else if (strcmp(wanted_cols[i], "FB") == 0) {
                // Get WIDF value from the current row
                char* widf_value = NULL;
                if (col_indices[1] >= 0 && col_indices[1] < col && row_values[col_indices[1]]) {
                    widf_value = strdup(row_values[col_indices[1]]);
                }
                
                // Use C function to get FB value by checking WIDF against REF in FB.xlsx
                char* fb_value = NULL;
                if (widf_value) {
                    // Force reload if requested
                    if (force_reload && fb_cache_loaded) {
                        clear_fb_hash_table();
                    }
                    fb_value = get_fb_value_by_widf(widf_value, current_week);
                    if (fb_value) {
                        worksheet_write_string(worksheet, row, i, fb_value, NULL);
                        printf("Found FB value for WIDF %s: %s\n", widf_value, fb_value);
                        free(fb_value);
                    } else {
                        printf("No FB value found for WIDF: %s\n", widf_value);
                    }
                    free(widf_value);
                }
                
            } else if (strcmp(wanted_cols[i], "MAX") == 0) {
                // Get WCMJ value from the current row
                char* wcmj_value = NULL;
                if (col_indices[8] >= 0 && col_indices[8] < col && row_values[col_indices[8]]) {
                    wcmj_value = strdup(row_values[col_indices[8]]);
                }
                
                // Get FB value for this row (we need to calculate it again for MAX)
                char* fb_value = NULL;
                char* widf_value = NULL;
                if (col_indices[1] >= 0 && col_indices[1] < col && row_values[col_indices[1]]) {
                    widf_value = strdup(row_values[col_indices[1]]);
                    fb_value = get_fb_value_by_widf(widf_value, current_week);
                }
                
                // Calculate MAX value using C function
                char* max_value = get_max_value(fb_value, wcmj_value);
                if (max_value) {
                    worksheet_write_string(worksheet, row, i, max_value, NULL);
                    printf("MAX value: %s (FB: %s, WCMJ: %s)\n", max_value, fb_value ? fb_value : "NULL", wcmj_value ? wcmj_value : "NULL");
                    free(max_value);
                }
///////////////////////////////the divel is here !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!               
                // Clean up
                if (fb_value) free(fb_value);
                if (wcmj_value) free(wcmj_value);
                if (widf_value) free(widf_value);
                
            } else if (col_indices[i] >= 0 && col_indices[i] < col && row_values[col_indices[i]]) {
                worksheet_write_string(worksheet, row, i, row_values[col_indices[i]], NULL);
            }
        }




        // Write extra columns for this row
        const char* wstkg_val_str = NULL;
        const char* max_val_str = NULL;
        if (col_indices[9] >= 0 && col_indices[9] < col && row_values[col_indices[9]])
            wstkg_val_str = row_values[col_indices[9]];
        if (col_indices[11] >= 0 && col_indices[11] < col) {
            // MAX may have been calculated and written above; try to read from output where possible is complex,
            // so recompute using helper already used for MAX calculation when available
            // Fall back to parsed string from row if present
            max_val_str = NULL; // we will recompute via get_max_value inputs above already wrote string value
        }

        // Leave Inventaire empty
        // (no write to column NUM_COLS)

        // Compute couv = WSTKG / MAX using values we wrote: parse numeric strings
        double wstkg_num = 0.0;
        double max_num = 0.0;
        int have_wstkg = 0, have_max = 0;
        if (wstkg_val_str && strlen(wstkg_val_str) > 0) { wstkg_num = atof(wstkg_val_str); have_wstkg = 1; }
        // For MAX, attempt to reuse the computed string by recalculating again
        // Recompute MAX: need WCMJ and FB
        const char* wcmj_str = NULL;
        const char* widf_for_row = NULL;
        char* fb_val_tmp = NULL;
        if (col_indices[8] >= 0 && col_indices[8] < col && row_values[col_indices[8]])
            wcmj_str = row_values[col_indices[8]];
        if (col_indices[1] >= 0 && col_indices[1] < col && row_values[col_indices[1]])
            widf_for_row = row_values[col_indices[1]];
        if (widf_for_row) {
            char* fb_calc = get_fb_value_by_widf(widf_for_row, current_week);
            fb_val_tmp = fb_calc;
        }
        char* max_calc = get_max_value(fb_val_tmp, (char*)wcmj_str);
        if (max_calc && strlen(max_calc) > 0) { max_num = atof(max_calc); have_max = 1; }
        if (fb_val_tmp) free(fb_val_tmp);
        if (max_calc) free(max_calc);

        if (have_wstkg && have_max && max_num != 0.0) {
            worksheet_write_number(worksheet, row, NUM_COLS + 1, wstkg_num / max_num, NULL);
        }

        for (int i = 0; i < col; i++)
            free(row_values[i]);
        row++;
    }

    // Cleanup
    xlsxioread_sheet_close(sheet);
    xlsxioread_close(reader);
    workbook_close(workbook);
    clear_fb_hash_table(); // Clear the FB hash table
    clear_abc_hash_table(); // Clear the ABC hash table

    printf("Processed %d data rows\n", data_row_count);
    printf("Extraction complete: %s\n", output_file);
    
    return 0;
}