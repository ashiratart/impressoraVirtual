/*
 grf2pcx_final.c
 Versão final (único arquivo) - Windows / MinGW
 - Converte .grf (~DG) -> .pcx 1-bit (inverte cores)
 - Lista impressoras, permite escolher
 - Menu:
    1) Converter + Enviar + Imprimir
    2) Somente Converter
    3) Somente Enviar + Imprimir
    4) Sair
 - Sempre reenvia (GM) antes de imprimir (opção B escolhida)
 - NOME enviado em MAIÚSCULAS, sem extensão, truncado a 16 chars
 Compile:
   gcc -std=c11 -O2 grf2pcx_final.c -o grf2pcx_final.exe -lwinspool
*/

#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>
#include <ctype.h>
#include <io.h>
#include <direct.h>

#define MAX_PATH_LEN 4096

/* -------------------- PCX RLE writer -------------------- */
static void pcx_write_rle(FILE *out, const uint8_t *buf, size_t len) {
    size_t i = 0;
    while (i < len) {
        uint8_t v = buf[i];
        size_t run = 1;
        while (i + run < len && buf[i + run] == v && run < 63) run++;
        if (run > 1 || (v & 0xC0) == 0xC0) {
            uint8_t code = (uint8_t)(0xC0 | (int)run);
            fputc(code, out);
            fputc(v, out);
            i += run;
        } else {
            fputc(v, out);
            i++;
        }
    }
}

/* -------------------- Parse GRF (~DG) -------------------- */
static int parse_grf_file(const char *path, uint8_t **out_bytes, long *out_total_bytes, long *out_bytes_per_row, long *out_width, long *out_height) {
    FILE *f = fopen(path, "rb");
    if (!f) return -1;
    fseek(f, 0, SEEK_END);
    long fsize = ftell(f);
    fseek(f, 0, SEEK_SET);
    char *text = malloc((size_t)fsize + 1);
    if (!text) { fclose(f); return -2; }
    if (fread(text, 1, fsize, f) != (size_t)fsize) { free(text); fclose(f); return -3; }
    text[fsize] = '\0';
    fclose(f);

    char *p = strstr(text, "~DG");
    if (!p) { free(text); return -4; }
    p += 3;
    /* skip name */
    while (*p && *p != ',') p++;
    if (!*p) { free(text); return -5; }
    p++; /* after comma */

    /* total_bytes */
    char tok[128]; int ti = 0;
    while (*p && *p != ',' && ti + 1 < (int)sizeof(tok)) tok[ti++] = *p++;
    tok[ti] = '\0';
    if (*p != ',') { free(text); return -6; }
    long total_bytes = atol(tok);
    p++; /* after comma */

    /* bytes_per_row */
    ti = 0;
    while (*p && *p != ',' && ti + 1 < (int)sizeof(tok)) tok[ti++] = *p++;
    tok[ti] = '\0';
    long bytes_per_row = atol(tok);

    /* advance to first hex digit */
    while (*p && !isxdigit((unsigned char)*p)) p++;
    if (!*p) { free(text); return -7; }

    long capacity = (total_bytes > 0) ? total_bytes : 4096;
    uint8_t *buf = malloc((size_t)capacity);
    if (!buf) { free(text); return -8; }

    long count = 0;
    while (*p && (total_bytes <= 0 || count < total_bytes)) {
        while (*p && !isxdigit((unsigned char)*p)) p++;
        if (!*p) break;
        char h1 = *p++;
        while (*p && !isxdigit((unsigned char)*p)) p++;
        if (!*p) break;
        char h2 = *p++;
        int v1 = isdigit((unsigned char)h1) ? h1 - '0' : toupper((unsigned char)h1) - 'A' + 10;
        int v2 = isdigit((unsigned char)h2) ? h2 - '0' : toupper((unsigned char)h2) - 'A' + 10;
        uint8_t val = (uint8_t)((v1 << 4) | v2);
        if (count >= capacity) {
            capacity += 4096;
            uint8_t *tmp = realloc(buf, (size_t)capacity);
            if (!tmp) { free(buf); free(text); return -9; }
            buf = tmp;
        }
        buf[count++] = val;
    }

    if (count == 0) { free(buf); free(text); return -10; }
    if (count != total_bytes) total_bytes = count;

    long height = (bytes_per_row > 0) ? (total_bytes / bytes_per_row) : 0;
    long width = (bytes_per_row > 0) ? (bytes_per_row * 8) : 0;

    *out_bytes = buf;
    *out_total_bytes = total_bytes;
    *out_bytes_per_row = bytes_per_row;
    *out_height = height;
    *out_width = width;

    free(text);
    return 0;
}

/* -------------------- Create PCX from GRF -------------------- */
static int create_pcx_from_grf(const char *grfPath, const char *pcxPath) {
    uint8_t *imgbytes = NULL;
    long total_bytes = 0;
    long bytes_per_row = 0, width = 0, height = 0;
    if (parse_grf_file(grfPath, &imgbytes, &total_bytes, &bytes_per_row, &width, &height) != 0) {
        return -1;
    }
    if (bytes_per_row <= 0 || height <= 0) { free(imgbytes); return -2; }

    /* invert bits */
    for (long i = 0; i < total_bytes; ++i) imgbytes[i] = (uint8_t)~imgbytes[i];

    int bytes_per_line = (int)((width + 7) / 8);
    if (bytes_per_line % 2) bytes_per_line++;

    FILE *out = fopen(pcxPath, "wb");
    if (!out) { free(imgbytes); return -3; }

    uint8_t header[128];
    memset(header, 0, sizeof(header));
    header[0] = 0x0A;
    header[1] = 5;
    header[2] = 1;
    header[3] = 1;

    header[4]=0; header[5]=0; header[6]=0; header[7]=0;
    unsigned short xmax = (unsigned short)(width - 1);
    unsigned short ymax = (unsigned short)(height - 1);
    header[8] = xmax & 0xFF; header[9] = (xmax >> 8) & 0xFF;
    header[10] = ymax & 0xFF; header[11] = (ymax >> 8) & 0xFF;

    header[12] = (width & 0xFF); header[13] = (width >> 8) & 0xFF;
    header[14] = (height & 0xFF); header[15] = (height >> 8) & 0xFF;

    header[16] = 0; header[17] = 0; header[18] = 0;
    header[19] = 255; header[20] = 255; header[21] = 255;

    header[65] = 1;
    header[66] = bytes_per_line & 0xFF;
    header[67] = (bytes_per_line >> 8) & 0xFF;
    header[68] = 1;

    if (fwrite(header,1,128,out) != 128) { fclose(out); free(imgbytes); return -4; }

    uint8_t *linebuf = malloc((size_t)bytes_per_line);
    if (!linebuf) { fclose(out); free(imgbytes); return -5; }

    for (long y = 0; y < height; ++y) {
        long src = y * bytes_per_row;
        if (src + bytes_per_row <= total_bytes) {
            memcpy(linebuf, imgbytes + src, bytes_per_row);
        } else {
            long remain = total_bytes - src;
            if (remain > 0) memcpy(linebuf, imgbytes + src, remain);
            if (bytes_per_line > (int)remain) memset(linebuf + (remain>0?remain:0), 0x00, bytes_per_line - (remain>0?remain:0));
        }
        if (bytes_per_row < bytes_per_line) memset(linebuf + bytes_per_row, 0x00, bytes_per_line - bytes_per_row);
        pcx_write_rle(out, linebuf, bytes_per_line);
    }

    free(linebuf);
    fclose(out);
    free(imgbytes);
    return 0;
}

/* -------------------- List printers and choose -------------------- */
static char *choose_printer(void) {
    DWORD needed = 0, returned = 0;
    EnumPrintersA(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, NULL, 0, &needed, &returned);
    if (needed == 0) {
        printf("Nenhuma impressora encontrada.\n");
        return NULL;
    }
    BYTE *buf = malloc(needed);
    if (!buf) return NULL;
    if (!EnumPrintersA(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, buf, needed, &needed, &returned)) {
        free(buf); return NULL;
    }
    PRINTER_INFO_2A *pi2 = (PRINTER_INFO_2A*)buf;
    printf("Impressoras encontradas:\n");
    for (DWORD i = 0; i < returned; ++i) {
        printf("  [%u] %s\n", (unsigned)i, pi2[i].pPrinterName ? pi2[i].pPrinterName : "(sem nome)");
    }
    int choice = 0;
    printf("Escolha a impressora (numero): ");
    if (scanf("%d", &choice) != 1) { free(buf); return NULL; }
    if (choice < 0 || (DWORD)choice >= returned) choice = 0;
    const char *sel = pi2[choice].pPrinterName;
    char *ret = _strdup(sel ? sel : "");
    free(buf);
    int c; while ((c = getchar()) != '\n' && c != EOF) {}
    return ret;
}

/* -------------------- Send PCX to printer (GM store) -------------------- */
static int send_pcx_store_to_printer(const char *printerName, const char *pcxPath, const char *nameUpper16, long *out_sent_bytes) {
    if (!printerName || !*printerName) return -1;
    HANDLE hPrinter = NULL;
    if (!OpenPrinterA((LPSTR)printerName, &hPrinter, NULL)) {
        printf("OpenPrinter falhou para %s\n", printerName);
        return -2;
    }

    FILE *pf = fopen(pcxPath, "rb");
    if (!pf) { ClosePrinter(hPrinter); return -3; }
    fseek(pf, 0, SEEK_END);
    long filesize = ftell(pf);
    fseek(pf, 0, SEEK_SET);

    char gmcmd[64];
    snprintf(gmcmd, sizeof(gmcmd), "GM %s %ld\n", nameUpper16, filesize);

    DOC_INFO_1A doc;
    doc.pDocName = "StoreImagePCX";
    doc.pOutputFile = NULL;
    doc.pDatatype = "RAW";

    DWORD dwWritten = 0;
    if (StartDocPrinterA(hPrinter, 1, (LPBYTE)&doc) == 0) {
        fclose(pf); ClosePrinter(hPrinter); return -4;
    }
    if (!StartPagePrinter(hPrinter)) {
        EndDocPrinter(hPrinter); fclose(pf); ClosePrinter(hPrinter); return -5;
    }

    if (!WritePrinter(hPrinter, gmcmd, (DWORD)strlen(gmcmd), &dwWritten)) {
        EndPagePrinter(hPrinter); EndDocPrinter(hPrinter); fclose(pf); ClosePrinter(hPrinter); return -6;
    }

    uint8_t buf[8192];
    size_t r;
    while ((r = fread(buf, 1, sizeof(buf), pf)) > 0) {
        if (!WritePrinter(hPrinter, buf, (DWORD)r, &dwWritten)) {
            EndPagePrinter(hPrinter); EndDocPrinter(hPrinter); fclose(pf); ClosePrinter(hPrinter); return -7;
        }
    }

    EndPagePrinter(hPrinter);
    EndDocPrinter(hPrinter);
    fclose(pf);
    ClosePrinter(hPrinter);

    if (out_sent_bytes) *out_sent_bytes = filesize;
    return 0;
}

/* -------------------- Send print sequence (GG) -------------------- */
static int send_print_sequence(const char *printerName, char **namesUpper, int nNames) {
    if (!printerName || !*printerName) return -1;
    HANDLE hPrinter = NULL;
    if (!OpenPrinterA((LPSTR)printerName, &hPrinter, NULL)) {
        printf("OpenPrinter falhou para %s\n", printerName);
        return -2;
    }

    DOC_INFO_1A doc;
    doc.pDocName = "PrintImages";
    doc.pOutputFile = NULL;
    doc.pDatatype = "RAW";

    DWORD dwWritten = 0;
    if (StartDocPrinterA(hPrinter, 1, (LPBYTE)&doc) == 0) {
        ClosePrinter(hPrinter); return -3;
    }

    if (!StartPagePrinter(hPrinter)) {
        EndDocPrinter(hPrinter); ClosePrinter(hPrinter); return -4;
    }

    for (int i = 0; i < nNames; ++i) {
        char block[512];
        snprintf(block, sizeof(block),
            "N\n"
            "D12\n"
            "S3\n"
            "JF\n"
            "q600\n"
            "Q260,24\n"
            "GG60 60 %s\n"
            "P1\n",
            namesUpper[i]
        );
        if (!WritePrinter(hPrinter, block, (DWORD)strlen(block), &dwWritten)) {
            printf("WritePrinter falhou ao enviar bloco de impressão para %s\n", namesUpper[i]);
            /* continuar tentando para outras imagens */
        } else {
            /* sem sleep por padrão; se precisar, adicionar Sleep(ms) */
        }
    }

    EndPagePrinter(hPrinter);
    EndDocPrinter(hPrinter);
    ClosePrinter(hPrinter);
    return 0;
}

/* -------------------- Folder processing -------------------- */

/* convert all .grf to .pcx; fill names array with uppercase names (no ext) */
static int convert_all_grf_in_folder(const char *folderPath, char ***out_names, int *out_count) {
    char pattern[MAX_PATH_LEN];
    _snprintf(pattern, sizeof(pattern), "%s\\*.grf", folderPath);
    struct _finddata_t file;
    intptr_t h = _findfirst(pattern, &file);
    if (h == -1) {
        *out_names = NULL; *out_count = 0;
        return 0;
    }
    int capacity = 32, count = 0;
    char **names = malloc(sizeof(char*) * capacity);

    do {
        char inpath[MAX_PATH_LEN], outpath[MAX_PATH_LEN];
        _snprintf(inpath, sizeof(inpath), "%s\\%s", folderPath, file.name);
        strncpy(outpath, inpath, sizeof(outpath)-1);
        char *dot = strrchr(outpath, '.');
        if (dot) strcpy(dot, ".pcx"); else strncat(outpath, ".pcx", sizeof(outpath)-strlen(outpath)-1);

        printf("Convertendo: %s -> %s ... ", file.name, strrchr(outpath,'\\')?strrchr(outpath,'\\')+1:outpath);
        int r = create_pcx_from_grf(inpath, outpath);
        if (r == 0) {
            printf("OK\n");
            /* build uppercase name without extension truncated to 16 */
            char base[MAX_PATH_LEN];
            strncpy(base, file.name, sizeof(base)-1);
            char *d = strrchr(base, '.'); if (d) *d = '\0';
            char nameUpper[17]; memset(nameUpper,0,sizeof(nameUpper));
            size_t bl = strlen(base);
            for (size_t i=0;i<16 && i<bl;i++) nameUpper[i] = (char)toupper((unsigned char)base[i]);
            nameUpper[16]='\0';
            if (count >= capacity) { capacity*=2; names = realloc(names, sizeof(char*)*capacity); }
            names[count++] = _strdup(nameUpper);
        } else {
            printf("ERRO (codigo %d)\n", r);
        }
    } while (_findnext(h, &file) == 0);
    _findclose(h);

    *out_names = names; *out_count = count;
    return count;
}

/* send all .pcx in folder (store via GM) and collect the uppercase names used */
static int send_all_pcx_in_folder_store(const char *folderPath, const char *printerName, char ***out_names, int *out_count) {
    char pattern[MAX_PATH_LEN];
    _snprintf(pattern, sizeof(pattern), "%s\\*.pcx", folderPath);
    struct _finddata_t file;
    intptr_t h = _findfirst(pattern, &file);
    if (h == -1) { *out_names=NULL; *out_count=0; return 0; }
    int capacity = 32, count = 0;
    char **names = malloc(sizeof(char*) * capacity);

    do {
        char pcxpath[MAX_PATH_LEN];
        _snprintf(pcxpath, sizeof(pcxpath), "%s\\%s", folderPath, file.name);
        printf("Enviando (store): %s ... ", file.name);

        /* build uppercase name without ext truncated 16 */
        char base[MAX_PATH_LEN];
        strncpy(base, file.name, sizeof(base)-1);
        char *d = strrchr(base, '.'); if (d) *d = '\0';
        char nameUpper[17]; memset(nameUpper,0,sizeof(nameUpper));
        size_t bl = strlen(base);
        for (size_t i=0;i<16 && i<bl;i++) nameUpper[i] = (char)toupper((unsigned char)base[i]);
        nameUpper[16]='\0';

        long sent_bytes = 0;
        int r = send_pcx_store_to_printer(printerName, pcxpath, nameUpper, &sent_bytes);
        if (r == 0) {
            printf("OK (bytes=%ld)\n", sent_bytes);
            if (count >= capacity) { capacity*=2; names = realloc(names, sizeof(char*)*capacity); }
            names[count++] = _strdup(nameUpper);
        } else {
            printf("ERRO (codigo %d)\n", r);
        }
    } while (_findnext(h, &file) == 0);
    _findclose(h);

    *out_names = names; *out_count = count;
    return count;
}

/* free names array */
static void free_names_array(char **arr, int n) {
    if (!arr) return;
    for (int i=0;i<n;i++) free(arr[i]);
    free(arr);
}

/* -------------------- main -------------------- */
int main(void) {
    printf("=== GRF -> PCX Converter + Sender + Print Sequencer ===\n\n");

    while (1) {
        printf("Opcoes:\n");
        printf("  1) Converter + Enviar + Imprimir\n");
        printf("  2) Somente Converter\n");
        printf("  3) Somente Enviar + Imprimir\n");
        printf("  4) Sair\n");
        printf("Escolha (1-4): ");

        int opt = 0;
        if (scanf("%d", &opt) != 1) { printf("Entrada invalida.\n"); return 1; }
        int c; while ((c = getchar()) != '\n' && c != EOF) {} /* flush */

        if (opt == 4) { printf("Encerrando.\n"); break; }

        char *printerName = NULL;
        if (opt == 1 || opt == 3) {
            printerName = choose_printer();
            if (!printerName) { printf("Falha ao selecionar impressora.\n"); continue; }
            printf("Impressora selecionada: %s\n", printerName);
        }

        char folder[MAX_PATH_LEN];
        printf("Digite o caminho da pasta (ex: C:\\imagens): ");
        if (!fgets(folder, sizeof(folder), stdin)) { if (printerName) free(printerName); break; }
        folder[strcspn(folder, "\r\n")] = '\0';
        if (folder[0] == '\0') { printf("Nenhuma pasta informada.\n"); if (printerName) free(printerName); continue; }

        char **namesUpper = NULL;
        int namesCount = 0;

        if (opt == 1) {
            printf("\n== Convertendo .grf -> .pcx ==\n");
            convert_all_grf_in_folder(folder, &namesUpper, &namesCount);
            printf("Convertidos: %d\n", namesCount);

            /* Sempre reenviar antes de imprimir (opção B): send all .pcx (store) */
            free_names_array(namesUpper, namesCount);
            namesUpper = NULL; namesCount = 0;
            printf("\n== Enviando (store) todos os .pcx encontrados ==\n");
            send_all_pcx_in_folder_store(folder, printerName, &namesUpper, &namesCount);
            printf("Enviados (store): %d\n", namesCount);

            /* Reenvio novamente imediatamente antes da impressão (garantia - opção B) */
            if (namesCount > 0) {
                printf("\n== Reenviando (store) antes da impressao (garantia) ==\n");
                /* simply call send_all_pcx_in_folder_store again to re-store and refresh names list */
                free_names_array(namesUpper, namesCount);
                namesUpper = NULL; namesCount = 0;
                send_all_pcx_in_folder_store(folder, printerName, &namesUpper, &namesCount);
                printf("Reenviados (store): %d\n", namesCount);
            }

            if (namesCount > 0) {
                printf("\n== Enviando blocos de impressao sequenciais ==\n");
                send_print_sequence(printerName, namesUpper, namesCount);
                printf("Sequencia de impressao enviada para %d imagens.\n", namesCount);
            } else {
                printf("Nenhuma imagem armazenada para imprimir.\n");
            }

            free_names_array(namesUpper, namesCount);
            namesUpper = NULL; namesCount = 0;
        }
        else if (opt == 2) {
            printf("\n== Somente Converter ==\n");
            int conv = convert_all_grf_in_folder(folder, &namesUpper, &namesCount);
            printf("Convertidos: %d\n", conv);
            free_names_array(namesUpper, namesCount);
            namesUpper = NULL; namesCount = 0;
        }
        else if (opt == 3) {
            printf("\n== Somente Enviar + Imprimir ==\n");
            send_all_pcx_in_folder_store(folder, printerName, &namesUpper, &namesCount);
            printf("Enviados (store): %d\n", namesCount);

            /* Reenvio novamente antes de imprimir (opção B): */
            if (namesCount > 0) {
                printf("\n== Reenviando (store) antes da impressao (garantia) ==\n");
                free_names_array(namesUpper, namesCount);
                namesUpper = NULL; namesCount = 0;
                send_all_pcx_in_folder_store(folder, printerName, &namesUpper, &namesCount);
                printf("Reenviados (store): %d\n", namesCount);
            }

            if (namesCount > 0) {
                printf("\n== Enviando blocos de impressao sequenciais ==\n");
                send_print_sequence(printerName, namesUpper, namesCount);
                printf("Sequencia de impressao enviada para %d imagens.\n", namesCount);
            } else {
                printf("Nenhuma imagem armazenada para imprimir.\n");
            }

            free_names_array(namesUpper, namesCount);
            namesUpper = NULL; namesCount = 0;
        }
        else {
            printf("Opcao invalida.\n");
        }

        if (printerName) free(printerName);
        printf("\nOperacao concluida. Voltar ao menu.\n\n");
    }

    return 0;
}
