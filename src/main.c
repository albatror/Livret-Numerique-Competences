#include <gtk/gtk.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "data/data.h"
#include "pptx/pptx_gen.h"

typedef struct {
    GtkWidget *domain_tree;
    GtkWidget *comp_list;
    GtkWidget *selected_tree;
    CompetenceList available_comps;
    DomaineList domaines;
    DomainColorList colors;

    GtkWidget *entry_nom;
    GtkWidget *entry_prenom;
    GtkWidget *entry_date;
    GtkWidget *entry_month;
    GtkWidget *entry_year;

    GtkWidget *preview_canvas;
    GtkWidget *page_label;
    int current_page;
    int total_pages;

    StudentInfo student;
    SectionData sections[4];
    SelectedCompetence *selected;
    int selected_count;

    // Image handling
    double img_x, img_y;
    double drag_start_x, drag_start_y;
    int is_dragging;
} AppState;

static void update_preview(AppState *state) {
    state->total_pages = (state->selected_count > 0) ? (state->selected_count + 4) / 5 : 1;
    if (state->current_page >= state->total_pages) state->current_page = state->total_pages - 1;
    if (state->current_page < 0) state->current_page = 0;

    char buf[128];
    sprintf(buf, "Page %d/%d", state->current_page + 1, state->total_pages);
    gtk_label_set_text(GTK_LABEL(state->page_label), buf);
    gtk_widget_queue_draw(state->preview_canvas);
}

static void on_import_photo_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    GtkWidget *dialog = gtk_file_chooser_dialog_new("Choisir une photo", NULL, GTK_FILE_CHOOSER_ACTION_OPEN, "_Annuler", GTK_RESPONSE_CANCEL, "_Ouvrir", GTK_RESPONSE_ACCEPT, NULL);
    if (gtk_dialog_run(GTK_DIALOG(dialog)) == GTK_RESPONSE_ACCEPT) {
        char *filename = gtk_file_chooser_get_filename(GTK_FILE_CHOOSER(dialog));
        if (state->student.photo_path) free(state->student.photo_path);
        state->student.photo_path = strdup(filename);
        g_free(filename);
        update_preview(state);
    }
    gtk_widget_destroy(dialog);
}

static void on_export_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    PPTXData data;
    data.name = (char *)gtk_entry_get_text(GTK_ENTRY(state->entry_nom));
    data.first_name = (char *)gtk_entry_get_text(GTK_ENTRY(state->entry_prenom));
    data.selected = state->selected;
    data.selected_count = state->selected_count;
    data.photo_path = state->student.photo_path;
    data.photo_x = state->img_x;
    data.photo_y = state->img_y;

    generate_pptx("presentation.pptx", &data);

    GtkWidget *dialog = gtk_message_dialog_new(NULL, GTK_DIALOG_MODAL, GTK_MESSAGE_INFO, GTK_BUTTONS_OK, "PowerPoint exporté avec succès !");
    gtk_dialog_run(GTK_DIALOG(dialog));
    gtk_widget_destroy(dialog);
}

static void on_add_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    GList *selected_rows = gtk_list_box_get_selected_rows(GTK_LIST_BOX(state->comp_list));
    GtkTreeModel *model = gtk_tree_view_get_model(GTK_TREE_VIEW(state->selected_tree));
    GtkListStore *store = GTK_LIST_STORE(model);
    const char *month = gtk_entry_get_text(GTK_ENTRY(state->entry_month));
    const char *year = gtk_entry_get_text(GTK_ENTRY(state->entry_year));
    char ts[128];
    sprintf(ts, "%s %s", month, year);
    for (GList *l = selected_rows; l != NULL; l = l->next) {
        GtkListBoxRow *row = GTK_LIST_BOX_ROW(l->data);
        GtkWidget *label = gtk_bin_get_child(GTK_BIN(row));
        const char *text = gtk_label_get_text(GTK_LABEL(label));
        state->selected = realloc(state->selected, sizeof(SelectedCompetence) * (state->selected_count + 1));
        state->selected[state->selected_count].text = strdup(text);
        state->selected[state->selected_count].timestamp = strdup(ts);
        state->selected_count++;
        GtkTreeIter iter;
        gtk_list_store_append(store, &iter);
        gtk_list_store_set(store, &iter, 0, ts, 1, text, -1);
    }
    g_list_free(selected_rows);
    update_preview(state);
}

static void on_remove_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    GtkTreeSelection *selection = gtk_tree_view_get_selection(GTK_TREE_VIEW(state->selected_tree));
    GtkTreeIter iter;
    GtkTreeModel *model;
    if (gtk_tree_selection_get_selected(selection, &model, &iter)) {
        GtkTreePath *path = gtk_tree_model_get_path(model, &iter);
        int idx = gtk_tree_path_get_indices(path)[0];
        free(state->selected[idx].text);
        free(state->selected[idx].timestamp);
        for(int i=idx; i<state->selected_count-1; i++) state->selected[i] = state->selected[i+1];
        state->selected_count--;
        gtk_list_store_remove(GTK_LIST_STORE(model), &iter);
        gtk_tree_path_free(path);
    }
    update_preview(state);
}

static void on_prev_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    if (state->current_page > 0) { state->current_page--; update_preview(state); }
}

static void on_next_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    if (state->current_page < state->total_pages - 1) { state->current_page++; update_preview(state); }
}

static gboolean on_button_press(GtkWidget *widget, GdkEventButton *event, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    if (event->button == 1) {
        if (event->x >= state->img_x && event->x <= state->img_x + 100 && event->y >= state->img_y && event->y <= state->img_y + 100) {
            state->is_dragging = 1;
            state->drag_start_x = event->x - state->img_x;
            state->drag_start_y = event->y - state->img_y;
        }
    }
    return TRUE;
}

static gboolean on_button_release(GtkWidget *widget, GdkEventButton *event, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    state->is_dragging = 0;
    return TRUE;
}

static gboolean on_motion_notify(GtkWidget *widget, GdkEventMotion *event, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    if (state->is_dragging) {
        state->img_x = event->x - state->drag_start_x;
        state->img_y = event->y - state->drag_start_y;
        gtk_widget_queue_draw(state->preview_canvas);
    }
    return TRUE;
}

static gboolean on_preview_draw(GtkWidget *widget, cairo_t *cr, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    int width = gtk_widget_get_allocated_width(widget);
    int height = gtk_widget_get_allocated_height(widget);
    cairo_set_source_rgb(cr, 1, 1, 1);
    cairo_rectangle(cr, 0, 0, width, height);
    cairo_fill(cr);

    if (state->student.photo_path && state->current_page == 0) {
        GdkPixbuf *pixbuf = gdk_pixbuf_new_from_file_at_scale(state->student.photo_path, 100, 100, TRUE, NULL);
        if (pixbuf) {
            gdk_cairo_set_source_pixbuf(cr, pixbuf, state->img_x, state->img_y);
            cairo_paint(cr);
            g_object_unref(pixbuf);
        }
    }

    if (state->selected_count == 0) {
        cairo_set_source_rgb(cr, 0.5, 0.5, 0.5);
        cairo_select_font_face(cr, "Sans", CAIRO_FONT_SLANT_NORMAL, CAIRO_FONT_WEIGHT_NORMAL);
        cairo_set_font_size(cr, 20);
        cairo_move_to(cr, width / 2 - 120, height / 2);
        cairo_show_text(cr, "Aucune compétence sélectionnée");
    } else {
        cairo_set_source_rgb(cr, 0.2, 0.5, 0.8);
        cairo_rectangle(cr, 0, 0, width, 40);
        cairo_fill(cr);
        cairo_set_source_rgb(cr, 1, 1, 1);
        cairo_set_font_size(cr, 16);
        cairo_move_to(cr, 10, 25);
        cairo_show_text(cr, "Page de garde / Compétences");

        int start = state->current_page * 5;
        int end = start + 5;
        if (end > state->selected_count) end = state->selected_count;
        cairo_set_source_rgb(cr, 0, 0, 0);
        cairo_set_font_size(cr, 14);
        int y = 70;
        for (int i = start; i < end; i++) {
            cairo_move_to(cr, 20, y);
            cairo_show_text(cr, "• ");
            cairo_show_text(cr, state->selected[i].text);
            y += 20;
            cairo_set_source_rgb(cr, 0.4, 0.4, 0.4);
            cairo_set_font_size(cr, 10);
            cairo_move_to(cr, 35, y);
            cairo_show_text(cr, state->selected[i].timestamp);
            cairo_set_source_rgb(cr, 0, 0, 0);
            cairo_set_font_size(cr, 14);
            y += 25;
        }
    }
    return FALSE;
}

static void on_domain_selected(GtkTreeSelection *selection, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    GtkTreeIter iter;
    GtkTreeModel *model;
    char *name;
    if (gtk_tree_selection_get_selected(selection, &model, &iter)) {
        gtk_tree_model_get(model, &iter, 0, &name, -1);
        GList *children = gtk_container_get_children(GTK_CONTAINER(state->comp_list));
        for (GList *l = children; l != NULL; l = l->next) gtk_widget_destroy(GTK_WIDGET(l->data));
        g_list_free(children);
        for (int i = 0; i < state->available_comps.count; i++) {
            Competence *c = &state->available_comps.items[i];
            if (strcmp(c->domain, name) == 0 || (c->subdomain && strcmp(c->subdomain, name) == 0)) {
                GtkWidget *row = gtk_list_box_row_new();
                GtkWidget *label = gtk_label_new(c->text);
                gtk_label_set_xalign(GTK_LABEL(label), 0);
                gtk_container_add(GTK_CONTAINER(row), label);
                gtk_list_box_insert(GTK_LIST_BOX(state->comp_list), row, -1);
            }
        }
        gtk_widget_show_all(state->comp_list);
        g_free(name);
    }
}

static void activate(GtkApplication *app, gpointer user_data) {
    GtkWidget *window = gtk_application_window_new(app);
    gtk_window_set_title(GTK_WINDOW(window), "Livret Numérique des Compétences Pro (C Version)");
    gtk_window_set_default_size(GTK_WINDOW(window), 1200, 900);
    GtkWidget *main_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 0);
    gtk_container_add(GTK_CONTAINER(window), main_box);

    AppState *state = g_malloc0(sizeof(AppState));
    state->available_comps = load_competences("COMPETENCES.csv");
    state->domaines = load_domaines("DOMAINES.csv");
    state->colors = load_couleurs("COULEURS_DOMAINES.csv");
    state->total_pages = 1;
    state->img_x = 100; state->img_y = 100;

    // Personal info area
    GtkWidget *personal_frame = gtk_frame_new("Informations personnelles & Horodatage (obligatoire)");
    gtk_box_pack_start(GTK_BOX(main_box), personal_frame, FALSE, FALSE, 5);
    GtkWidget *personal_box = gtk_box_new(GTK_ORIENTATION_HORIZONTAL, 10);
    gtk_container_add(GTK_CONTAINER(personal_frame), personal_box);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_label_new("Nom:"), FALSE, FALSE, 5);
    state->entry_nom = gtk_entry_new(); gtk_box_pack_start(GTK_BOX(personal_box), state->entry_nom, TRUE, TRUE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_label_new("Prénom:"), FALSE, FALSE, 5);
    state->entry_prenom = gtk_entry_new(); gtk_box_pack_start(GTK_BOX(personal_box), state->entry_prenom, TRUE, TRUE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_label_new("Mois:"), FALSE, FALSE, 5);
    state->entry_month = gtk_entry_new(); gtk_box_pack_start(GTK_BOX(personal_box), state->entry_month, TRUE, TRUE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_label_new("Année:"), FALSE, FALSE, 5);
    state->entry_year = gtk_entry_new(); gtk_box_pack_start(GTK_BOX(personal_box), state->entry_year, TRUE, TRUE, 5);
    GtkWidget *import_btn = gtk_button_new_with_label("Importer Photo"); g_signal_connect(import_btn, "clicked", G_CALLBACK(on_import_photo_clicked), state); gtk_box_pack_start(GTK_BOX(personal_box), import_btn, FALSE, FALSE, 5);

    // Columns area
    GtkWidget *columns_box = gtk_box_new(GTK_ORIENTATION_HORIZONTAL, 10);
    gtk_box_pack_start(GTK_BOX(main_box), columns_box, TRUE, TRUE, 5);

    // Column 1
    GtkWidget *col1_frame = gtk_frame_new("Compétences disponibles");
    gtk_box_pack_start(GTK_BOX(columns_box), col1_frame, TRUE, TRUE, 0);
    GtkTreeStore *store = gtk_tree_store_new(1, G_TYPE_STRING);
    char *last_domain = NULL;
    GtkTreeIter domain_iter, sub_iter;
    for(int i=0; i<state->available_comps.count; i++) {
        Competence *c = &state->available_comps.items[i];
        if (!last_domain || strcmp(c->domain, last_domain) != 0) {
            gtk_tree_store_append(store, &domain_iter, NULL); gtk_tree_store_set(store, &domain_iter, 0, c->domain, -1);
            last_domain = c->domain;
        }
        if (c->subdomain && strlen(c->subdomain) > 0) {
            GtkTreeIter child; gboolean found = FALSE;
            if (gtk_tree_model_iter_children(GTK_TREE_MODEL(store), &child, &domain_iter)) {
                do { char *n; gtk_tree_model_get(GTK_TREE_MODEL(store), &child, 0, &n, -1); if (strcmp(n, c->subdomain) == 0) found = TRUE; g_free(n); } while (!found && gtk_tree_model_iter_next(GTK_TREE_MODEL(store), &child));
            }
            if (!found) { gtk_tree_store_append(store, &sub_iter, &domain_iter); gtk_tree_store_set(store, &sub_iter, 0, c->subdomain, -1); }
        }
    }
    state->domain_tree = gtk_tree_view_new_with_model(GTK_TREE_MODEL(store));
    gtk_tree_view_append_column(GTK_TREE_VIEW(state->domain_tree), gtk_tree_view_column_new_with_attributes("Domaine", gtk_cell_renderer_text_new(), "text", 0, NULL));
    g_signal_connect(gtk_tree_view_get_selection(GTK_TREE_VIEW(state->domain_tree)), "changed", G_CALLBACK(on_domain_selected), state);
    GtkWidget *s1 = gtk_scrolled_window_new(NULL, NULL); gtk_container_add(GTK_CONTAINER(s1), state->domain_tree); gtk_container_add(GTK_CONTAINER(col1_frame), s1);

    // Column 2
    GtkWidget *col2_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 5); gtk_box_pack_start(GTK_BOX(columns_box), col2_box, TRUE, TRUE, 0);
    GtkWidget *col2_frame = gtk_frame_new("Compétences du sous-domaine"); gtk_box_pack_start(GTK_BOX(col2_box), col2_frame, TRUE, TRUE, 0);
    state->comp_list = gtk_list_box_new(); gtk_list_box_set_selection_mode(GTK_LIST_BOX(state->comp_list), GTK_SELECTION_MULTIPLE);
    GtkWidget *s2 = gtk_scrolled_window_new(NULL, NULL); gtk_container_add(GTK_CONTAINER(s2), state->comp_list); gtk_container_add(GTK_CONTAINER(col2_frame), s2);
    GtkWidget *add_btn = gtk_button_new_with_label("Ajouter ->"); g_signal_connect(add_btn, "clicked", G_CALLBACK(on_add_clicked), state); gtk_box_pack_start(GTK_BOX(col2_box), add_btn, FALSE, FALSE, 5);

    // Column 3
    GtkWidget *col3_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 5); gtk_box_pack_start(GTK_BOX(columns_box), col3_box, TRUE, TRUE, 0);
    GtkWidget *col3_frame = gtk_frame_new("Sélectionnées (dans le PPT)"); gtk_box_pack_start(GTK_BOX(col3_box), col3_frame, TRUE, TRUE, 0);
    GtkListStore *sel_store = gtk_list_store_new(2, G_TYPE_STRING, G_TYPE_STRING);
    state->selected_tree = gtk_tree_view_new_with_model(GTK_TREE_MODEL(sel_store));
    GtkCellRenderer *r = gtk_cell_renderer_text_new();
    gtk_tree_view_append_column(GTK_TREE_VIEW(state->selected_tree), gtk_tree_view_column_new_with_attributes("TS", r, "text", 0, NULL));
    gtk_tree_view_append_column(GTK_TREE_VIEW(state->selected_tree), gtk_tree_view_column_new_with_attributes("Compétence", r, "text", 1, NULL));
    GtkWidget *s3 = gtk_scrolled_window_new(NULL, NULL); gtk_container_add(GTK_CONTAINER(s3), state->selected_tree); gtk_container_add(GTK_CONTAINER(col3_frame), s3);
    GtkWidget *col3_btns = gtk_box_new(GTK_ORIENTATION_HORIZONTAL, 5); gtk_box_pack_start(GTK_BOX(col3_box), col3_btns, FALSE, FALSE, 5);
    GtkWidget *rem_btn = gtk_button_new_with_label("Retirer <-"); g_signal_connect(rem_btn, "clicked", G_CALLBACK(on_remove_clicked), state); gtk_box_pack_start(GTK_BOX(col3_btns), rem_btn, TRUE, TRUE, 0);
    GtkWidget *exp_btn = gtk_button_new_with_label("Exporter PPTX"); g_signal_connect(exp_btn, "clicked", G_CALLBACK(on_export_clicked), state); gtk_box_pack_start(GTK_BOX(col3_btns), exp_btn, TRUE, TRUE, 0);

    // Preview area
    GtkWidget *preview_frame = gtk_frame_new("Aperçu de la page (Glisser la photo pour la placer)");
    gtk_box_pack_start(GTK_BOX(main_box), preview_frame, TRUE, TRUE, 5);
    GtkWidget *preview_vbox = gtk_box_new(GTK_ORIENTATION_VERTICAL, 5);
    gtk_container_add(GTK_CONTAINER(preview_frame), preview_vbox);
    GtkWidget *pager_box = gtk_box_new(GTK_ORIENTATION_HORIZONTAL, 10);
    gtk_box_pack_start(GTK_BOX(preview_vbox), pager_box, FALSE, FALSE, 0);
    GtkWidget *prev_btn = gtk_button_new_with_label("◀ Précédent"); g_signal_connect(prev_btn, "clicked", G_CALLBACK(on_prev_clicked), state); gtk_box_pack_start(GTK_BOX(pager_box), prev_btn, FALSE, FALSE, 0);
    state->page_label = gtk_label_new("Page 1/1"); gtk_box_pack_start(GTK_BOX(pager_box), state->page_label, TRUE, TRUE, 0);
    GtkWidget *next_btn = gtk_button_new_with_label("Suivant ▶"); g_signal_connect(next_btn, "clicked", G_CALLBACK(on_next_clicked), state); gtk_box_pack_start(GTK_BOX(pager_box), next_btn, FALSE, FALSE, 0);
    state->preview_canvas = gtk_drawing_area_new();
    gtk_widget_set_size_request(state->preview_canvas, -1, 400);
    gtk_widget_add_events(state->preview_canvas, GDK_BUTTON_PRESS_MASK | GDK_BUTTON_RELEASE_MASK | GDK_POINTER_MOTION_MASK);
    g_signal_connect(state->preview_canvas, "draw", G_CALLBACK(on_preview_draw), state);
    g_signal_connect(state->preview_canvas, "button-press-event", G_CALLBACK(on_button_press), state);
    g_signal_connect(state->preview_canvas, "button-release-event", G_CALLBACK(on_button_release), state);
    g_signal_connect(state->preview_canvas, "motion-notify-event", G_CALLBACK(on_motion_notify), state);
    gtk_box_pack_start(GTK_BOX(preview_vbox), state->preview_canvas, TRUE, TRUE, 0);

    gtk_widget_show_all(window);
}

int main(int argc, char **argv) {
    GtkApplication *app = gtk_application_new("org.albatror.livret", G_APPLICATION_DEFAULT_FLAGS);
    g_signal_connect(app, "activate", G_CALLBACK(activate), NULL);
    int status = g_application_run(G_APPLICATION(app), argc, argv);
    g_object_unref(app);
    return status;
}
