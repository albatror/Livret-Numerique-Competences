#include "data.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static char *trim_quotes_and_unescape(char *str) {
    if (str[0] == '"') {
        str++;
        size_t len = strlen(str);
        if (len > 0 && str[len-1] == '"') {
            str[len-1] = '\0';
        }
    }
    char *res = strdup(str);
    if (!res) return NULL;

    char *src = res, *dst = res;
    while (*src) {
        if (*src == '"' && *(src+1) == '"') {
            *dst++ = '"';
            src += 2;
        } else {
            *dst++ = *src++;
        }
    }
    *dst = '\0';
    return res;
}

static char **parse_csv_line(FILE *f, int *num_cols) {
    char **cols = NULL;
    int count = 0;
    int in_quotes = 0;
    size_t buffer_size = 1024;
    char *buffer = malloc(buffer_size);
    if (!buffer) {
        *num_cols = 0;
        return NULL;
    }
    size_t pos = 0;
    int c;

    while (1) {
        c = fgetc(f);
        if (c == EOF) break;

        if (pos + 1 >= buffer_size) {
            size_t new_size = buffer_size * 2;
            char *new_buffer = realloc(buffer, new_size);
            if (!new_buffer) break; // Should probably handle better
            buffer = new_buffer;
            buffer_size = new_size;
        }

        if (c == '"') {
            in_quotes = !in_quotes;
            buffer[pos++] = c;
        } else if (c == ',' && !in_quotes) {
            buffer[pos] = '\0';
            char **new_cols = realloc(cols, sizeof(char *) * (count + 1));
            if (new_cols) {
                cols = new_cols;
                cols[count++] = trim_quotes_and_unescape(buffer);
            }
            pos = 0;
        } else if (c == '\n' && !in_quotes) {
            break;
        } else if (c == '\r' && !in_quotes) {
            // skip
        } else {
            buffer[pos++] = c;
        }
    }

    if (pos > 0 || count > 0) {
        buffer[pos] = '\0';
        char **new_cols = realloc(cols, sizeof(char *) * (count + 1));
        if (new_cols) {
            cols = new_cols;
            cols[count++] = trim_quotes_and_unescape(buffer);
        }
    } else if (c == EOF && count == 0) {
        free(buffer);
        *num_cols = 0;
        return NULL;
    }

    free(buffer);
    *num_cols = count;
    return cols;
}

CompetenceList load_competences(const char *filename) {
    FILE *f = fopen(filename, "r");
    if (!f) return (CompetenceList){NULL, 0};

    int num_cols;
    char **header = parse_csv_line(f, &num_cols);
    if(header) {
        for(int i=0; i<num_cols; i++) free(header[i]);
        free(header);
    }

    Competence *items = NULL;
    int count = 0;

    while (1) {
        char **cols = parse_csv_line(f, &num_cols);
        if (cols == NULL) break;
        if (num_cols >= 3) {
            Competence *new_items = realloc(items, sizeof(Competence) * (count + 1));
            if (new_items) {
                items = new_items;
                items[count].domain = cols[0];
                items[count].subdomain = cols[1];
                items[count].text = cols[2];
                count++;
                for (int i = 3; i < num_cols; i++) free(cols[i]);
            } else {
                for (int i = 0; i < num_cols; i++) free(cols[i]);
            }
        } else {
            for (int i = 0; i < num_cols; i++) free(cols[i]);
        }
        free(cols);
    }

    fclose(f);
    return (CompetenceList){items, count};
}

void free_competences(CompetenceList list) {
    if (!list.items) return;
    for (int i = 0; i < list.count; i++) {
        free(list.items[i].domain);
        free(list.items[i].subdomain);
        free(list.items[i].text);
    }
    free(list.items);
}

DomaineList load_domaines(const char *filename) {
    FILE *f = fopen(filename, "r");
    if (!f) return (DomaineList){NULL, 0};
    int num_cols;
    char **header = parse_csv_line(f, &num_cols);
    if(header) { for(int i=0; i<num_cols; i++) free(header[i]); free(header); }
    Domaine *items = NULL;
    int count = 0;
    while (1) {
        char **cols = parse_csv_line(f, &num_cols);
        if (cols == NULL) break;
        if (num_cols >= 3) {
            Domaine *new_items = realloc(items, sizeof(Domaine) * (count + 1));
            if (new_items) {
                items = new_items;
                items[count].domain = cols[0];
                items[count].subdomain = cols[1];
                items[count].description = cols[2];
                count++;
                for (int i = 3; i < num_cols; i++) free(cols[i]);
            } else {
                for (int i = 0; i < num_cols; i++) free(cols[i]);
            }
        } else {
            for (int i = 0; i < num_cols; i++) free(cols[i]);
        }
        free(cols);
    }
    fclose(f);
    return (DomaineList){items, count};
}

void free_domaines(DomaineList list) {
    if (!list.items) return;
    for (int i = 0; i < list.count; i++) {
        free(list.items[i].domain);
        free(list.items[i].subdomain);
        free(list.items[i].description);
    }
    free(list.items);
}

DomainColorList load_couleurs(const char *filename) {
    FILE *f = fopen(filename, "r");
    if (!f) return (DomainColorList){NULL, 0};
    int num_cols;
    char **header = parse_csv_line(f, &num_cols);
    if(header) { for(int i=0; i<num_cols; i++) free(header[i]); free(header); }
    DomainColor *items = NULL;
    int count = 0;
    while (1) {
        char **cols = parse_csv_line(f, &num_cols);
        if (cols == NULL) break;
        if (num_cols >= 2) {
            DomainColor *new_items = realloc(items, sizeof(DomainColor) * (count + 1));
            if (new_items) {
                items = new_items;
                items[count].domain = cols[0];
                items[count].color = cols[1];
                count++;
                for (int i = 2; i < num_cols; i++) free(cols[i]);
            } else {
                for (int i = 0; i < num_cols; i++) free(cols[i]);
            }
        } else {
            for (int i = 0; i < num_cols; i++) free(cols[i]);
        }
        free(cols);
    }
    fclose(f);
    return (DomainColorList){items, count};
}

void free_couleurs(DomainColorList list) {
    if (!list.items) return;
    for (int i = 0; i < list.count; i++) {
        free(list.items[i].domain);
        free(list.items[i].color);
    }
    free(list.items);
}
