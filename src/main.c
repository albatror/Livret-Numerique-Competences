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
} AppState;

static void on_add_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    GList *selected_rows = gtk_list_box_get_selected_rows(GTK_LIST_BOX(state->comp_list));
    GtkTreeModel *model = gtk_tree_view_get_model(GTK_TREE_VIEW(state->selected_tree));
    GtkListStore *store = GTK_LIST_STORE(model);

    for (GList *l = selected_rows; l != NULL; l = l->next) {
        GtkListBoxRow *row = GTK_LIST_BOX_ROW(l->data);
        GtkWidget *label = gtk_bin_get_child(GTK_BIN(row));
        const char *text = gtk_label_get_text(GTK_LABEL(label));

        GtkTreeIter iter;
        gtk_list_store_append(store, &iter);
        gtk_list_store_set(store, &iter, 0, "", 1, text, -1);
    }
    g_list_free(selected_rows);
}

static void on_remove_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    GtkTreeSelection *selection = gtk_tree_view_get_selection(GTK_TREE_VIEW(state->selected_tree));
    GtkTreeIter iter;
    GtkTreeModel *model;

    if (gtk_tree_selection_get_selected(selection, &model, &iter)) {
        gtk_list_store_remove(GTK_LIST_STORE(model), &iter);
    }
}

static void on_export_clicked(GtkButton *button, gpointer user_data) {
    generate_pptx("presentation.pptx");
    GtkWidget *dialog = gtk_message_dialog_new(NULL, GTK_DIALOG_MODAL, GTK_MESSAGE_INFO, GTK_BUTTONS_OK, "PowerPoint exporté avec succès !");
    gtk_dialog_run(GTK_DIALOG(dialog));
    gtk_widget_destroy(dialog);
}

static void on_domain_selected(GtkTreeSelection *selection, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    GtkTreeIter iter;
    GtkTreeModel *model;
    char *name;

    if (gtk_tree_selection_get_selected(selection, &model, &iter)) {
        gtk_tree_model_get(model, &iter, 0, &name, -1);

        GList *children, *l;
        children = gtk_container_get_children(GTK_CONTAINER(state->comp_list));
        for (l = children; l != NULL; l = l->next)
            gtk_widget_destroy(GTK_WIDGET(l->data));
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
    GtkWidget *window;
    GtkWidget *main_box;

    window = gtk_application_window_new(app);
    gtk_window_set_title(GTK_WINDOW(window), "Livret Numérique des Compétences Pro (C Version)");
    gtk_window_set_default_size(GTK_WINDOW(window), 1200, 800);

    main_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 0);
    gtk_container_add(GTK_CONTAINER(window), main_box);

    // Personal info area
    GtkWidget *personal_frame = gtk_frame_new("Informations personnelles & Horodatage (obligatoire)");
    gtk_box_pack_start(GTK_BOX(main_box), personal_frame, FALSE, FALSE, 5);
    GtkWidget *personal_box = gtk_box_new(GTK_ORIENTATION_HORIZONTAL, 10);
    gtk_container_add(GTK_CONTAINER(personal_frame), personal_box);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_label_new("Nom:"), FALSE, FALSE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_entry_new(), TRUE, TRUE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_label_new("Prénom:"), FALSE, FALSE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_entry_new(), TRUE, TRUE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_label_new("Date de naissance:"), FALSE, FALSE, 5);
    gtk_box_pack_start(GTK_BOX(personal_box), gtk_entry_new(), TRUE, TRUE, 5);

    // Sections (Notebook)
    GtkWidget *sections_frame = gtk_frame_new("Sections de cycle");
    gtk_box_pack_start(GTK_BOX(main_box), sections_frame, FALSE, FALSE, 5);
    GtkWidget *notebook = gtk_notebook_new();
    gtk_container_add(GTK_CONTAINER(sections_frame), notebook);
    const char *section_labels[] = {"TPS", "PS", "MS", "GS"};
    for (int i = 0; i < 4; i++) {
        GtkWidget *vbox = gtk_box_new(GTK_ORIENTATION_VERTICAL, 5);
        gtk_box_pack_start(GTK_BOX(vbox), gtk_check_button_new_with_label("Marquer cette section comme complétée"), FALSE, FALSE, 0);
        GtkWidget *grid = gtk_grid_new();
        gtk_grid_set_column_spacing(GTK_GRID(grid), 10);
        gtk_grid_set_row_spacing(GTK_GRID(grid), 5);
        gtk_grid_attach(GTK_GRID(grid), gtk_label_new("Année scolaire:"), 0, 0, 1, 1);
        gtk_grid_attach(GTK_GRID(grid), gtk_entry_new(), 1, 0, 1, 1);
        gtk_grid_attach(GTK_GRID(grid), gtk_label_new("École:"), 0, 1, 1, 1);
        gtk_grid_attach(GTK_GRID(grid), gtk_entry_new(), 1, 1, 1, 1);
        gtk_grid_attach(GTK_GRID(grid), gtk_label_new("Enseignant(s):"), 0, 2, 1, 1);
        gtk_grid_attach(GTK_GRID(grid), gtk_entry_new(), 1, 2, 1, 1);
        gtk_box_pack_start(GTK_BOX(vbox), grid, FALSE, FALSE, 0);
        gtk_notebook_append_page(GTK_NOTEBOOK(notebook), vbox, gtk_label_new(section_labels[i]));
    }

    // Main area with columns
    GtkWidget *columns_box = gtk_box_new(GTK_ORIENTATION_HORIZONTAL, 10);
    gtk_box_pack_start(GTK_BOX(main_box), columns_box, TRUE, TRUE, 5);

    AppState *state = g_malloc0(sizeof(AppState));
    state->available_comps = load_competences("COMPETENCES.csv");
    state->domaines = load_domaines("DOMAINES.csv");
    state->colors = load_couleurs("COULEURS_DOMAINES.csv");

    // Column 1: Available domains
    GtkWidget *col1_frame = gtk_frame_new("Compétences disponibles");
    gtk_box_pack_start(GTK_BOX(columns_box), col1_frame, TRUE, TRUE, 0);
    gtk_widget_set_size_request(col1_frame, 300, -1);
    GtkTreeStore *store = gtk_tree_store_new(1, G_TYPE_STRING);
    char *last_domain = NULL;
    GtkTreeIter domain_iter, sub_iter;
    for(int i=0; i<state->available_comps.count; i++) {
        Competence *c = &state->available_comps.items[i];
        if (!last_domain || strcmp(c->domain, last_domain) != 0) {
            gtk_tree_store_append(store, &domain_iter, NULL);
            gtk_tree_store_set(store, &domain_iter, 0, c->domain, -1);
            last_domain = c->domain;
        }
        if (c->subdomain && strlen(c->subdomain) > 0) {
             GtkTreeIter child;
             gboolean found = FALSE;
             if (gtk_tree_model_iter_children(GTK_TREE_MODEL(store), &child, &domain_iter)) {
                 do {
                     char *sub_name;
                     gtk_tree_model_get(GTK_TREE_MODEL(store), &child, 0, &sub_name, -1);
                     if (strcmp(sub_name, c->subdomain) == 0) found = TRUE;
                     g_free(sub_name);
                 } while (!found && gtk_tree_model_iter_next(GTK_TREE_MODEL(store), &child));
             }
             if (!found) {
                 gtk_tree_store_append(store, &sub_iter, &domain_iter);
                 gtk_tree_store_set(store, &sub_iter, 0, c->subdomain, -1);
             }
        }
    }
    state->domain_tree = gtk_tree_view_new_with_model(GTK_TREE_MODEL(store));
    gtk_tree_view_append_column(GTK_TREE_VIEW(state->domain_tree), gtk_tree_view_column_new_with_attributes("Domaine", gtk_cell_renderer_text_new(), "text", 0, NULL));
    GtkTreeSelection *selection = gtk_tree_view_get_selection(GTK_TREE_VIEW(state->domain_tree));
    g_signal_connect(selection, "changed", G_CALLBACK(on_domain_selected), state);
    GtkWidget *scroll1 = gtk_scrolled_window_new(NULL, NULL);
    gtk_container_add(GTK_CONTAINER(scroll1), state->domain_tree);
    gtk_container_add(GTK_CONTAINER(col1_frame), scroll1);

    // Column 2: Competences list
    GtkWidget *col2_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 5);
    gtk_box_pack_start(GTK_BOX(columns_box), col2_box, TRUE, TRUE, 0);
    GtkWidget *col2_frame = gtk_frame_new("Compétences du sous-domaine");
    gtk_box_pack_start(GTK_BOX(col2_box), col2_frame, TRUE, TRUE, 0);
    state->comp_list = gtk_list_box_new();
    gtk_list_box_set_selection_mode(GTK_LIST_BOX(state->comp_list), GTK_SELECTION_MULTIPLE);
    GtkWidget *scroll2 = gtk_scrolled_window_new(NULL, NULL);
    gtk_container_add(GTK_CONTAINER(scroll2), state->comp_list);
    gtk_container_add(GTK_CONTAINER(col2_frame), scroll2);
    GtkWidget *add_btn = gtk_button_new_with_label("Ajouter ->");
    g_signal_connect(add_btn, "clicked", G_CALLBACK(on_add_clicked), state);
    gtk_box_pack_start(GTK_BOX(col2_box), add_btn, FALSE, FALSE, 5);

    // Column 3: Selected competences
    GtkWidget *col3_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 5);
    gtk_box_pack_start(GTK_BOX(columns_box), col3_box, TRUE, TRUE, 0);
    GtkWidget *col3_frame = gtk_frame_new("Sélectionnées (dans le PPT)");
    gtk_box_pack_start(GTK_BOX(col3_box), col3_frame, TRUE, TRUE, 0);
    GtkListStore *sel_store = gtk_list_store_new(2, G_TYPE_STRING, G_TYPE_STRING);
    state->selected_tree = gtk_tree_view_new_with_model(GTK_TREE_MODEL(sel_store));
    GtkCellRenderer *sel_rend = gtk_cell_renderer_text_new();
    gtk_tree_view_append_column(GTK_TREE_VIEW(state->selected_tree), gtk_tree_view_column_new_with_attributes("Sous-domaine", sel_rend, "text", 0, NULL));
    gtk_tree_view_append_column(GTK_TREE_VIEW(state->selected_tree), gtk_tree_view_column_new_with_attributes("Compétence", sel_rend, "text", 1, NULL));
    GtkWidget *scroll3 = gtk_scrolled_window_new(NULL, NULL);
    gtk_container_add(GTK_CONTAINER(scroll3), state->selected_tree);
    gtk_container_add(GTK_CONTAINER(col3_frame), scroll3);

    GtkWidget *col3_btns = gtk_box_new(GTK_ORIENTATION_HORIZONTAL, 5);
    gtk_box_pack_start(GTK_BOX(col3_box), col3_btns, FALSE, FALSE, 5);
    GtkWidget *remove_btn = gtk_button_new_with_label("Retirer <-");
    g_signal_connect(remove_btn, "clicked", G_CALLBACK(on_remove_clicked), state);
    gtk_box_pack_start(GTK_BOX(col3_btns), remove_btn, TRUE, TRUE, 0);
    GtkWidget *export_btn = gtk_button_new_with_label("Exporter PowerPoint");
    g_signal_connect(export_btn, "clicked", G_CALLBACK(on_export_clicked), state);
    gtk_box_pack_start(GTK_BOX(col3_btns), export_btn, TRUE, TRUE, 0);

    gtk_widget_show_all(window);
}

int main(int argc, char **argv) {
    GtkApplication *app;
    int status;
    app = gtk_application_new("org.albatror.livret", G_APPLICATION_DEFAULT_FLAGS);
    g_signal_connect(app, "activate", G_CALLBACK(activate), NULL);
    status = g_application_run(G_APPLICATION(app), argc, argv);
    g_object_unref(app);
    return status;
}
