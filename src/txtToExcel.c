#include <gtk/gtk.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "xlsxwriter.h"

#define MAX_NOME_ARQUIVO 255
#define MAX_TITULO 255

GtkWidget *file_entry, *excel_name_entry, *title_entry;

// Função para criar o Excel
void create_excel(const char *txt_file_path, const char *excel_file_name, const char *title) {
    FILE *arquivo;
    char buffer[255];
    char nome_arquivo[MAX_NOME_ARQUIVO];
    char *linhas[100];
    int i = 0;

    // Cria o nome do arquivo Excel
    snprintf(nome_arquivo, sizeof(nome_arquivo), "%s.xlsx", excel_file_name);

    // Abre o arquivo para leitura
    arquivo = fopen(txt_file_path, "r");
    if (arquivo == NULL) {
        perror("Erro ao abrir o arquivo");
        return;
    }

    // Lê as linhas do arquivo
    while (i < 100 && fgets(buffer, sizeof(buffer), arquivo)) {
        linhas[i] = malloc(strlen(buffer) + 1);
        if (linhas[i] == NULL) {
            perror("Erro ao alocar memória");
            return;
        }
        strcpy(linhas[i], buffer);
        i++;
    }
    fclose(arquivo);

    // Cria um novo arquivo Excel
    lxw_workbook  *workbook = workbook_new(nome_arquivo);
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    // Escreve o título na primeira linha
    worksheet_write_string(worksheet, 0, 0, title, NULL);
    worksheet_write_string(worksheet, 1, 0, "PIS", NULL);
    worksheet_write_string(worksheet, 1, 1, "Data", NULL);
    worksheet_write_string(worksheet, 1, 2, "Hora", NULL);

    // Escreve os dados
    for (int j = 0; j < i; j++) {
        char pis[11] = {0};
        char data[9] = {0};
        char hora[255] = {0};

        strncpy(pis, linhas[j], 10);
        strncpy(data, linhas[j] + 10, 8);
        strcpy(hora, linhas[j] + 18);

        worksheet_write_string(worksheet, j + 2, 0, pis, NULL);
        worksheet_write_string(worksheet, j + 2, 1, data, NULL);
        worksheet_write_string(worksheet, j + 2, 2, hora, NULL);
        
        free(linhas[j]);
    }

    // Fecha o arquivo Excel
    workbook_close(workbook);
    
    // Notificação de sucesso
    GtkWidget *dialog = gtk_message_dialog_new(NULL, GTK_DIALOG_MODAL, GTK_MESSAGE_INFO, GTK_BUTTONS_OK, "Arquivo Excel '%s' criado com sucesso!", nome_arquivo);
    gtk_dialog_run(GTK_DIALOG(dialog));
    gtk_widget_destroy(dialog);
}

// Callback para o botão de criar Excel
void on_create_excel_clicked(GtkWidget *widget, gpointer data) {
    const char *txt_file_path = gtk_entry_get_text(GTK_ENTRY(file_entry));
    const char *excel_file_name = gtk_entry_get_text(GTK_ENTRY(excel_name_entry));
    const char *title = gtk_entry_get_text(GTK_ENTRY(title_entry));

    create_excel(txt_file_path, excel_file_name, title);
}

// Callback para abrir o seletor de arquivos
void on_file_select_clicked(GtkWidget *widget, gpointer data) {
    GtkWidget *dialog;
    dialog = gtk_file_chooser_dialog_new("Selecionar Arquivo de Texto",
                                         GTK_WINDOW(data),
                                         GTK_FILE_CHOOSER_ACTION_OPEN,
                                         "_Cancelar", GTK_RESPONSE_CANCEL,
                                         "_Selecionar", GTK_RESPONSE_ACCEPT,
                                         NULL);

    if (gtk_dialog_run(GTK_DIALOG(dialog)) == GTK_RESPONSE_ACCEPT) {
        gchar *filename = gtk_file_chooser_get_filename(GTK_FILE_CHOOSER(dialog));
        gtk_entry_set_text(GTK_ENTRY(file_entry), filename);
        g_free(filename);  // Libera a memória alocada pelo GTK
    }

    gtk_widget_destroy(dialog);
}

int main(int argc, char *argv[]) {
    gtk_init(&argc, &argv);

    GtkWidget *window;
    GtkWidget *grid;
    GtkWidget *file_select_button;
    GtkWidget *create_excel_button;

    window = gtk_window_new(GTK_WINDOW_TOPLEVEL);
    gtk_window_set_title(GTK_WINDOW(window), "Converter .txt para excel formatado");
    gtk_window_set_default_size(GTK_WINDOW(window), 400, 200);
    
    g_signal_connect(window, "destroy", G_CALLBACK(gtk_main_quit), NULL);

    grid = gtk_grid_new();
    gtk_container_add(GTK_CONTAINER(window), grid);
    
    // Adiciona padding à grid
    gtk_grid_set_column_homogeneous(GTK_GRID(grid), TRUE);
    gtk_grid_set_row_homogeneous(GTK_GRID(grid), TRUE);
    gtk_widget_set_margin_top(grid, 16);
    gtk_widget_set_margin_bottom(grid, 16);
    gtk_widget_set_margin_start(grid, 16);
    gtk_widget_set_margin_end(grid, 16);

    // Campo para caminho do arquivo
    GtkWidget *file_label = gtk_label_new("Caminho do Arquivo:");
    gtk_grid_attach(GTK_GRID(grid), file_label, 0, 0, 1, 1);
    file_entry = gtk_entry_new();
    gtk_grid_attach(GTK_GRID(grid), file_entry, 0, 1, 1, 1);

    // Botão para selecionar arquivo
    file_select_button = gtk_button_new_with_label("Selecionar");
    g_signal_connect(file_select_button, "clicked", G_CALLBACK(on_file_select_clicked), window);
    gtk_grid_attach(GTK_GRID(grid), file_select_button, 1, 1, 1, 1);

    // Adiciona uma margem superior ao botão
    gtk_widget_set_margin_start(file_select_button, 8);

    // Campo para nome do arquivo Excel
    GtkWidget *excel_label = gtk_label_new("Nome do Arquivo Excel:");
    gtk_grid_attach(GTK_GRID(grid), excel_label, 0, 2, 2, 1);
    excel_name_entry = gtk_entry_new();
    gtk_grid_attach(GTK_GRID(grid), excel_name_entry, 0, 3, 2, 1);

    // Campo para título
    GtkWidget *title_label = gtk_label_new("Título:");
    gtk_grid_attach(GTK_GRID(grid), title_label, 0, 4, 2, 1);
    title_entry = gtk_entry_new();
    gtk_grid_attach(GTK_GRID(grid), title_entry, 0, 5, 2, 1);

    // Botão para criar Excel
    create_excel_button = gtk_button_new_with_label("Gerar Excel");
    g_signal_connect(create_excel_button, "clicked", G_CALLBACK(on_create_excel_clicked), NULL);
    gtk_grid_attach(GTK_GRID(grid), create_excel_button, 0, 6, 2, 2);

    // Adiciona uma margem superior ao botão
    gtk_widget_set_margin_top(create_excel_button, 30);

    gtk_widget_show_all(window);
    gtk_main();

    return 0;
}
