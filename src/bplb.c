// bplb.c - Implementação das funções BPLB
#include "bplb.h"

// Utilitários
int bplb_mm_to_points(float mm) {
    return (int)(mm * BPLB_UNIT_MM);
}

float bplb_points_to_mm(int points) {
    return (float)points / BPLB_UNIT_MM;
}

// Configuração básica
void bplb_init(bplb_config_t *config) {
    config->density = 10;
    config->speed = 3;
    config->label_height = 0;
    config->label_gap = 0;
    config->label_width = 832; // 104mm padrão
    config->direction = DIRECTION_TOP;
}

void bplb_clear_buffer(FILE *fp) {
    fprintf(fp, "N" BPLB_LF);
}

void bplb_set_density(FILE *fp, int density) {
    if (density >= 0 && density <= 15) {
        fprintf(fp, "D%d" BPLB_LF, density);
    }
}

void bplb_set_speed(FILE *fp, int speed) {
    if (speed >= 1 && speed <= 3) {
        fprintf(fp, "S%d" BPLB_LF, speed);
    }
}

void bplb_set_label_size(FILE *fp, int height, int gap) {
    fprintf(fp, "Q%d,%d" BPLB_LF, height, gap);
}

void bplb_set_label_width(FILE *fp, int width) {
    fprintf(fp, "q%d" BPLB_LF, width);
}

void bplb_set_direction(FILE *fp, bplb_direction_t direction) {
    fprintf(fp, "%c" BPLB_LF, direction);
}

void bplb_enable_backfeed(FILE *fp) {
    fprintf(fp, "a1" BPLB_LF);
}

void bplb_disable_backfeed(FILE *fp) {
    fprintf(fp, "a0" BPLB_LF);
}

// Comandos de texto
void bplb_text(FILE *fp, int x, int y, bplb_rotation_t rotation, 
               char font, int h_mult, int v_mult, 
               bplb_reverse_t reverse, const char *text) {
    fprintf(fp, "A%d,%d,%d,%c,%d,%d,%c,\"%s\"" BPLB_LF,
            x, y, rotation, font, h_mult, v_mult, reverse, text);
}

// Comandos de código de barras 1D
void bplb_barcode_1d(FILE *fp, int x, int y, bplb_rotation_t rotation,
                     const char *symbology, int thin_width, int thick_width,
                     int height, bplb_human_readable_t human_readable,
                     const char *data) {
    fprintf(fp, "B%d,%d,%d,%s,%d,%d,%d,%c,\"%s\"" BPLB_LF,
            x, y, rotation, symbology, thin_width, thick_width, 
            height, human_readable, data);
}

// Comandos de código de barras 2D PDF417
// Comandos de código de barras 2D PDF417
void bplb_barcode_2d_pdf417(FILE *fp, int x, int y, int max_width, int max_height,
                           int error_level, int compression, int module_width,
                           int module_height, int max_rows, int max_cols,
                           int truncate, bplb_rotation_t rotation,
                           const char *data) {
    fprintf(fp, "b%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,\"%s\"" BPLB_LF,
            x, y, max_width, max_height, error_level, compression,
            module_width, module_height, max_rows, max_cols, truncate,
            rotation, data);
}

// Comando de impressão
void bplb_print(FILE *fp, int copies, int identical_copies) {
    fprintf(fp, "P%d,%d" BPLB_LF, copies, identical_copies);
}

// Comandos de formulários
void bplb_form_save_begin(FILE *fp, const char *form_name) {
    fprintf(fp, BPLB_STX "F1,\"%s\"" BPLB_LF, form_name);
}

void bplb_form_save_end(FILE *fp) {
    fprintf(fp, BPLB_STX "F0" BPLB_LF);
}

void bplb_form_load(FILE *fp, const char *form_name) {
    fprintf(fp, BPLB_STX "FL,\"%s\"" BPLB_LF, form_name);
}

void bplb_form_delete(FILE *fp, const char *form_name) {
    fprintf(fp, BPLB_STX "FD,\"%s\"" BPLB_LF, form_name);
}

void bplb_form_delete_all(FILE *fp) {
    fprintf(fp, BPLB_STX "F*" BPLB_LF);
}

// Comandos de contador e variáveis
void bplb_counter(FILE *fp, int counter_id, int max_digits, 
                  bplb_align_t align, const char *increment, 
                  const char *message) {
    fprintf(fp, BPLB_STX "C%d,%d,%c,%s,\"%s\"" BPLB_LF,
            counter_id, max_digits, align, increment, message);
}

void bplb_variable(FILE *fp, int var_id, int max_chars, 
                   bplb_align_t align, const char *message) {
    fprintf(fp, BPLB_STX "V%d,%d,%c,\"%s\"" BPLB_LF,
            var_id, max_chars, align, message);
}

void bplb_enable_variable_substitution(FILE *fp) {
    fprintf(fp, BPLB_STX "V1" BPLB_LF);
}

// Comandos de imagem
void bplb_image_print(FILE *fp, int x, int y, const char *image_name) {
    fprintf(fp, "GG%d,%d,\"%s\"" BPLB_LF, x, y, image_name);
}

void bplb_image_store_begin(FILE *fp, const char *image_name, int size_bytes) {
    fprintf(fp, BPLB_STX "G1,\"%s\",%d" BPLB_LF, image_name, size_bytes);
}

void bplb_image_delete(FILE *fp, const char *image_name) {
    fprintf(fp, BPLB_STX "GD,\"%s\"" BPLB_LF, image_name);
}

// Comandos de desenho
void bplb_line_non_overlap(FILE *fp, int x, int y, int width, int height) {
    fprintf(fp, "LS%d,%d,%d,%d" BPLB_LF, x, y, width, height);
}

void bplb_line_overlap(FILE *fp, int x, int y, int width, int height) {
    fprintf(fp, "LE%d,%d,%d,%d" BPLB_LF, x, y, width, height);
}

void bplb_white_line(FILE *fp, int x, int y, int width, int height) {
    fprintf(fp, "X%d,%d,%d,%d" BPLB_LF, x, y, width, height);
}

void bplb_rectangle(FILE *fp, int x1, int y1, int thickness, int x2, int y2) {
    fprintf(fp, "FR%d,%d,%d,%d,%d" BPLB_LF, x1, y1, thickness, x2, y2);
}

// Comandos especiais
void bplb_calibrate(FILE *fp) {
    fprintf(fp, BPLB_STX "CAL" BPLB_LF);
}

void bplb_reset_counter(FILE *fp) {
    fprintf(fp, BPLB_STX "RC" BPLB_LF);
}

void bplb_factory_reset(FILE *fp) {
    fprintf(fp, BPLB_STX "DEFAULT" BPLB_LF);
}

void bplb_save_as_default(FILE *fp) {
    fprintf(fp, BPLB_STX "SAVE" BPLB_LF);
}