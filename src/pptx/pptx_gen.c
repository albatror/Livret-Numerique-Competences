#include "pptx_gen.h"
#include <stdio.h>
#include <string.h>
#include <stdlib.h>

static void add_file_to_zip(struct zip *archive, const char *name, const char *content) {
    struct zip_source *s = zip_source_buffer(archive, content, strlen(content), 0);
    zip_file_add(archive, name, s, ZIP_FL_ENC_UTF_8);
}

static void add_image_to_zip(struct zip *archive, const char *zip_path, const char *file_path) {
    struct zip_source *s = zip_source_file(archive, file_path, 0, 0);
    if (s) zip_file_add(archive, zip_path, s, ZIP_FL_ENC_UTF_8);
}

static char* xml_escape(const char* input) {
    if (!input) return strdup("");
    size_t len = strlen(input);
    char* output = malloc(len * 6 + 1);
    char* ptr = output;
    while (*input) {
        switch (*input) {
            case '&':  strcpy(ptr, "&amp;");  ptr += 5; break;
            case '<':  strcpy(ptr, "&lt;");   ptr += 4; break;
            case '>':  strcpy(ptr, "&gt;");   ptr += 4; break;
            case '"':  strcpy(ptr, "&quot;"); ptr += 6; break;
            case '\'': strcpy(ptr, "&apos;"); ptr += 6; break;
            default:   *ptr++ = *input; break;
        }
        input++;
    }
    *ptr = '\0';
    return output;
}

void generate_pptx(const char *filename, PPTXData *data) {
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
        "<Default Extension=\"png\" ContentType=\"image/png\"/>\n"
        "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>\n"
        "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>\n"
        "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>\n"
        "<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>\n"
        "<Override PartName=\"/ppt/slides/slide2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>\n"
        "</Types>");

    add_file_to_zip(archive, "_rels/.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>\n"
        "</Relationships>");

    add_file_to_zip(archive, "ppt/presentation.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n"
        "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/></p:sldIdLst>\n"
        "<p:notesSz x=\"6858000\" y=\"9144000\"/><p:sldSz x=\"9144000\" y=\"6858000\"/>\n"
        "</p:presentation>");

    add_file_to_zip(archive, "ppt/_rels/presentation.xml.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>\n"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide2.xml\"/>\n"
        "</Relationships>");

    if (data->photo_path) {
        add_image_to_zip(archive, "ppt/media/photo.jpg", data->photo_path);
        add_file_to_zip(archive, "ppt/slides/_rels/slide1.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/photo.jpg\"/>\n"
            "</Relationships>");
    }

    char *esc_first = xml_escape(data->first_name);
    char *esc_last = xml_escape(data->name);

    // Cover Slide (Slide 1)
    char slide1_xml[8192];
    sprintf(slide1_xml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n"
        "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>\n"
        "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Title\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>\n"
        "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>%s %s</a:t></a:r></a:p></p:txBody></p:sp>\n", esc_first, esc_last);

    if (data->photo_path) {
        // Convert screen coordinates to EMUs (approx 1 px = 9525 EMUs)
        long lx = (long)(data->photo_x * 9525);
        long ly = (long)(data->photo_y * 9525);
        char img_xml[1024];
        sprintf(img_xml, "<p:pic><p:nvPicPr><p:cNvPr id=\"4\" name=\"Student Photo\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>\n"
            "<p:blipFill><a:blip r:embed=\"rId1\"/></p:blipFill>\n"
            "<p:spPr><a:xfrm><a:off x=\"%ld\" y=\"%ld\"/><a:ext cx=\"1000000\" cy=\"1000000\"/></a:xfrm><a:prstGeom prst=\"rect\"/></p:spPr></p:pic>\n", lx, ly);
        strcat(slide1_xml, img_xml);
    }

    strcat(slide1_xml, "</p:spTree></p:cSld></p:sld>");
    add_file_to_zip(archive, "ppt/slides/slide1.xml", slide1_xml);
    free(esc_first); free(esc_last);

    // Competencies Slide (Slide 2)
    char *slide2_content = malloc(131072);
    strcpy(slide2_content, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n"
        "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>");

    int id = 10;
    int y_off = 500000;
    for (int i = 0; i < data->selected_count && i < 15; i++) {
        char *esc_text = xml_escape(data->selected[i].text);
        char item_xml[2048];
        sprintf(item_xml, "<p:sp><p:nvSpPr><p:cNvPr id=\"%d\" name=\"Item %d\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>\n"
            "<p:spPr><a:xfrm><a:off x=\"500000\" y=\"%d\"/><a:ext cx=\"8000000\" cy=\"300000\"/></a:xfrm></p:spPr>\n"
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>• %s</a:t></a:r></a:p></p:txBody></p:sp>", id++, i, y_off, esc_text);
        strcat(slide2_content, item_xml);
        free(esc_text);
        y_off += 400000;
    }
    strcat(slide2_content, "</p:spTree></p:cSld></p:sld>");
    add_file_to_zip(archive, "ppt/slides/slide2.xml", slide2_content);
    free(slide2_content);

    zip_close(archive);
    printf("Generated enhanced PPTX with photo and %d competencies: %s\n", data->selected_count, filename);
}
