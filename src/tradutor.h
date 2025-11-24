#ifndef ZPL_TO_BPLB_H
#define ZPL_TO_BPLB_H

#include <stdio.h>
#include <string.h>
#include <stdlib.h>

// ============================================================================
// CONSTANTES E DEFINIÇÕES ZPL
// ============================================================================

// Comandos básicos de estrutura
#define ZPL_START_FORMAT "^XA"
#define ZPL_END_FORMAT "^XZ"
#define ZPL_FIELD_SEPARATOR "^FS"
#define ZPL_COMMENT "^FX"

// Comandos de configuração de etiqueta
#define ZPL_LABEL_HOME "^LH"        // ^LHx,y - Posição inicial da etiqueta
#define ZPL_PRINT_MODE "^MM"        // ^MMt - Modo de impressão (T=tear-off)
#define ZPL_MEDIA_TRACKING "^MN"    // ^MNy - Rastreamento de mídia
#define ZPL_UNIT_MEASURE "^MU"      // ^MUm - Unidade de medida (m=mm)

// Comandos de fonte e texto
#define ZPL_CHANGE_FONT "^CF"       // ^CFf,h,w - Alterar fonte padrão
#define ZPL_FONT "^A"               // ^Af,o,h,w - Especificar fonte
#define ZPL_FIELD_ORIGIN "^FO"      // ^FOx,y - Origem do campo
#define ZPL_FIELD_DATA "^FD"        // ^FDtexto - Dados do campo
#define ZPL_FIELD_NUMBER "^FN"      // ^FNx - Número do campo

// Comandos de código de barras
#define ZPL_BARCODE_128 "^BC"       // ^BCo,h,f,g,e,j - Code 128
#define ZPL_BARCODE_25 "^B2"        // ^B2o,h,f,g,e,j - Interleaved 2 of 5
#define ZPL_BARCODE_EAN "^BE"       // ^BEo,h,f,g - EAN-8/EAN-13

// Comandos gráficos
#define ZPL_GRAPHIC_BOX "^GB"       // ^GBw,h,t,c - Caixa/linha
#define ZPL_GRAPHIC_CIRCLE "^GC"    // ^GCw,t,c - Círculo
#define ZPL_GRAPHIC_ELLIPSE "^GE"   // ^GEw,h,t,c - Elipse

// Comandos de imagem
#define ZPL_IMAGE_DOWNLOAD "^DG"    // ^DGname,x,y - Download imagem
#define ZPL_IMAGE_FIELD "^GF"       // ^GFa,b,c,d,data - Campo de imagem
#define ZPL_IMAGE_RETRIEVE "^XG"    // ^XGname,x,y - Recuperar imagem

// Comandos de formato
#define ZPL_FORMAT_DOWNLOAD "^DF"   // ^DFname - Download formato
#define ZPL_FORMAT_RETRIEVE "^XF"   // ^XFname - Recuperar formato

// Comandos especiais
#define ZPL_REVERSE_PRINT "^FR"     // ^FR - Impressão reversa
#define ZPL_LABEL_REVERSE "^LR"     // ^LRy - Reverso da etiqueta
#define ZPL_SUPPRESS_BACKFEED "^XB" // ^XB - Suprimir backfeed

// ============================================================================
// ESTRUTURAS DE DADOS
// ============================================================================

typedef struct {
    int x;
    int y;
} ZPLPosition;

typedef struct {
    char font;
    char orientation;
    int height;
    int width;
} ZPLFont;

typedef struct {
    char type;          // 'B'=barcode, 'T'=text, 'I'=image, 'G'=graphic
    ZPLPosition pos;
    char data[256];
    ZPLFont font;
    int barcode_height;
    char barcode_type[10];
} ZPLField;

typedef struct {
    int width;
    int height;
    int thickness;
    char color;
    int corner_radius;
} ZPLGraphicBox;

// ============================================================================
// FUNÇÕES DE CONVERSÃO ZPL -> BPLB
// ============================================================================

/**
 * Converte comando ZPL ^FO (Field Origin) para BPLB
 */
void convert_FO_to_BPLB(int x, int y, FILE *out) {
    // Em BPLB, as coordenadas são usadas diretamente nos comandos
    // A conversão depende do contexto (texto, código de barras, etc.)
    fprintf(out, "// FO: x=%d, y=%d\n", x, y);
}

/**
 * Converte comando ZPL ^A (Font) para BPLB
 */
void convert_A_to_BPLB(char font, char orientation, int height, int width, FILE *out) {
    // Mapeamento de fontes ZPL para BPLB
    char bplb_font = '3'; // padrão
    
    switch(font) {
        case '0': bplb_font = '1'; break; // pequena
        case '1': bplb_font = '2'; break; // média  
        case '2': bplb_font = '3'; break; // grande
        default: bplb_font = '3'; break;
    }
    
    fprintf(out, "// Fonte ZPL: %c -> BPLB: %c, h=%d, w=%d\n", 
            font, bplb_font, height, width);
}

/**
 * Converte comando ZPL ^FD (Field Data) para BPLB (texto)
 */
void convert_FD_to_BPLB_text(int x, int y, const char *text, char font, FILE *out) {
    char bplb_font = '3';
    
    switch(font) {
        case '0': bplb_font = '1'; break;
        case '1': bplb_font = '2'; break;  
        case '2': bplb_font = '3'; break;
        default: bplb_font = '3'; break;
    }
    
    // Comando BPLB para texto: A<x>,<y>,<rotation>,<font>,<multiplier_x>,<multiplier_y>,N,"<text>"
    fprintf(out, "A%d,%d,0,%c,1,1,N,\"%s\"\n", x, y, bplb_font, text);
}

/**
 * Converte comando ZPL ^BC (Code 128) para BPLB
 */
void convert_BC_to_BPLB(int x, int y, int height, const char *data, FILE *out) {
    // Comando BPLB para código de barras: B<x>,<y>,<rotation>,<type>,<narrow>,<wide>,<height>,<format>,"<data>"
    // type: 1=Code128, 2=Code39, 3=Interleaved 2 of 5, etc.
    fprintf(out, "B%d,%d,0,1,2,4,%d,B,\"%s\"\n", x, y, height, data);
}

/**
 * Converte comando ZPL ^B2 (Interleaved 2 of 5) para BPLB
 */
void convert_B2_to_BPLB(int x, int y, int height, const char *data, FILE *out) {
    // Comando BPLB para Interleaved 2 of 5
    fprintf(out, "B%d,%d,0,3,2,4,%d,B,\"%s\"\n", x, y, height, data);
}

/**
 * Converte comando ZPL ^GB (Graphic Box) para BPLB
 */
void convert_GB_to_BPLB(int x, int y, int width, int height, int thickness, FILE *out) {
    // Comando BPLB para caixa: X<x>,<y>,<width>,<height>,<thickness>
    fprintf(out, "X%d,%d,%d,%d,%d\n", x, y, width, height, thickness);
}

/**
 * Converte comando ZPL ^XG (Image) para BPLB
 */
void convert_XG_to_BPLB(int x, int y, const char *image_name, FILE *out) {
    // Comando BPLB para imagem: GG<x>,<y>,"<image_name>"
    fprintf(out, "GG%d,%d,\"%s\"\n", x, y, image_name);
}

// ============================================================================
// FUNÇÕES DE PROCESSAMENTO
// ============================================================================

/**
 * Extrai coordenadas do comando ^FO
 */
ZPLPosition extract_FO_coordinates(const char *zpl_line) {
    ZPLPosition pos = {0, 0};
    if (strstr(zpl_line, "^FO") != NULL) {
        sscanf(zpl_line, "^FO%d,%d", &pos.x, &pos.y);
    }
    return pos;
}

/**
 * Extrai parâmetros de fonte do comando ^A
 */
ZPLFont extract_A_parameters(const char *zpl_line) {
    ZPLFont font = {'0', 'N', 10, 10};
    if (strstr(zpl_line, "^A") != NULL) {
        sscanf(zpl_line, "^A%c%c,%d,%d", &font.font, &font.orientation, &font.height, &font.width);
    }
    return font;
}

/**
 * Extrai dados do campo ^FD
 */
void extract_FD_data(const char *zpl_line, char *output) {
    output[0] = '\0';
    const char *fd_start = strstr(zpl_line, "^FD");
    if (fd_start) {
        fd_start += 3; // Pula "^FD"
        const char *fs_end = strstr(fd_start, "^FS");
        if (fs_end) {
            int len = fs_end - fd_start;
            if (len > 0 && len < 255) {
                strncpy(output, fd_start, len);
                output[len] = '\0';
            }
        }
    }
}

/**
 * Extrai altura do código de barras
 */
int extract_barcode_height(const char *zpl_line) {
    int height = 10; // padrão
    if (strstr(zpl_line, "^B") != NULL) {
        // Procura por padrão ^Bx,<orientation>,<height>
        const char *b_start = strstr(zpl_line, "^B");
        if (b_start) {
            char b_type[3] = {0};
            char orientation;
            sscanf(b_start, "^%2s%c,%d", b_type, &orientation, &height);
        }
    }
    return height;
}

/**
 * Processa linha ZPL e converte para BPLB
 */
void process_zpl_line(const char *zpl_line, FILE *out) {
    if (strstr(zpl_line, "^XA")) {
        fprintf(out, "N\n"); // Limpa buffer BPLB
        fprintf(out, "// Início do formato ZPL\n");
    }
    else if (strstr(zpl_line, "^XZ")) {
        fprintf(out, "P1\n"); // Imprime 1 cópia
        fprintf(out, "// Fim do formato ZPL\n");
    }
    else if (strstr(zpl_line, "^FO") && strstr(zpl_line, "^A") && strstr(zpl_line, "^FD")) {
        // Campo de texto: ^FOx,y^Af,o,h,w^FDtexto^FS
        ZPLPosition pos = extract_FO_coordinates(zpl_line);
        ZPLFont font = extract_A_parameters(zpl_line);
        char text[256];
        extract_FD_data(zpl_line, text);
        
        if (strlen(text) > 0) {
            convert_FD_to_BPLB_text(pos.x, pos.y, text, font.font, out);
        }
    }
    else if (strstr(zpl_line, "^FO") && strstr(zpl_line, "^BC") && strstr(zpl_line, "^FD")) {
        // Código de barras Code 128
        ZPLPosition pos = extract_FO_coordinates(zpl_line);
        int height = extract_barcode_height(zpl_line);
        char data[256];
        extract_FD_data(zpl_line, data);
        
        if (strlen(data) > 0) {
            convert_BC_to_BPLB(pos.x, pos.y, height, data, out);
        }
    }
    else if (strstr(zpl_line, "^FO") && strstr(zpl_line, "^B2") && strstr(zpl_line, "^FD")) {
        // Código de barras Interleaved 2 of 5
        ZPLPosition pos = extract_FO_coordinates(zpl_line);
        int height = extract_barcode_height(zpl_line);
        char data[256];
        extract_FD_data(zpl_line, data);
        
        if (strlen(data) > 0) {
            convert_B2_to_BPLB(pos.x, pos.y, height, data, out);
        }
    }
    else if (strstr(zpl_line, "^FO") && strstr(zpl_line, "^GB")) {
        // Caixa/linha gráfica
        ZPLPosition pos = extract_FO_coordinates(zpl_line);
        int width, height, thickness;
        if (strstr(zpl_line, "^GB")) {
            sscanf(zpl_line, "^GB%d,%d,%d", &width, &height, &thickness);
            convert_GB_to_BPLB(pos.x, pos.y, width, height, thickness, out);
        }
    }
    else if (strstr(zpl_line, "^XG")) {
        // Imagem
        ZPLPosition pos = extract_FO_coordinates(zpl_line);
        char image_name[128] = {0};
        sscanf(zpl_line, "^XG%[^,^ ]", image_name);
        convert_XG_to_BPLB(pos.x, pos.y, image_name, out);
    }
    else if (strstr(zpl_line, "^FX")) {
        // Comentário - preserva no BPLB
        fprintf(out, "// %s\n", zpl_line + 3);
    }
}

#endif // ZPL_TO_BPLB_H