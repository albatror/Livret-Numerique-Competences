#include "pptx_gen.h"
#include <stdio.h>
#include <string.h>

static void add_file_to_zip(struct zip *archive, const char *name, const char *content) {
    struct zip_source *s = zip_source_buffer(archive, content, strlen(content), 0);
    zip_file_add(archive, name, s, ZIP_FL_ENC_UTF_8);
}

void generate_pptx(const char *filename) {
    int error = 0;
    struct zip *archive = zip_open(filename, ZIP_CREATE | ZIP_TRUNCATE, &error);
    if (!archive) {
        fprintf(stderr, "Failed to create zip archive: %s\n", filename);
        return;
    }

    add_file_to_zip(archive, "[Content_Types].xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
        "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>\n"
        "<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>\n"
        "</Types>");

    add_file_to_zip(archive, "_rels/.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>\n"
        "</Relationships>");

    add_file_to_zip(archive, "ppt/presentation.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n"
        "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/></p:sldIdLst>\n"
        "<p:notesSz x=\"6858000\" y=\"9144000\"/><p:sldSz x=\"9144000\" y=\"6858000\"/>\n"
        "</p:presentation>");

    add_file_to_zip(archive, "ppt/_rels/presentation.xml.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>\n"
        "</Relationships>");

    add_file_to_zip(archive, "ppt/slides/slide1.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n"
        "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>\n"
        "</p:sld>");

    zip_close(archive);
    printf("Generated basic openable PPTX: %s\n", filename);
}
