#ifndef PPTX_GEN_H
#define PPTX_GEN_H

#include <zip.h>
#include <libxml/parser.h>
#include <libxml/tree.h>
#include "data/data.h"

typedef struct {
    char *name;
    char *first_name;
    SelectedCompetence *selected;
    int selected_count;
    char *photo_path;
    double photo_x;
    double photo_y;
} PPTXData;

void generate_pptx(const char *filename, PPTXData *data);

#endif
