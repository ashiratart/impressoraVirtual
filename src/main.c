#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <time.h>

#define BUF 4096

FILE *glog = NULL;

// ------------------------------------------------------------------
// LOG
// ------------------------------------------------------------------
void log_open()
{
    glog = fopen("tradutor.log", "a");
    if (!glog) {
        MessageBoxA(NULL, "Falha ao abrir tradutor.log", "Erro", MB_ICONERROR);
        exit(1);
    }
}

void log_write(const char *msg)
{
    if (!glog) return;

    time_t t = time(NULL);
    struct tm lt;
    localtime_s(&lt, &t);

    fprintf(glog, "[%04d-%02d-%02d %02d:%02d:%02d] %s\n",
            lt.tm_year+1900, lt.tm_mon+1, lt.tm_mday,
            lt.tm_hour, lt.tm_min, lt.tm_sec,
            msg);
    fflush(glog);
}

void log_printf(const char *fmt, ...)
{
    if (!glog) return;

    time_t t = time(NULL);
    struct tm lt;
    localtime_s(&lt, &t);

    fprintf(glog, "[%04d-%02d-%02d %02d:%02d:%02d] ",
            lt.tm_year+1900, lt.tm_mon+1, lt.tm_mday,
            lt.tm_hour, lt.tm_min, lt.tm_sec);

    va_list ap;
    va_start(ap, fmt);
    vfprintf(glog, fmt, ap);
    va_end(ap);

    fprintf(glog, "\n");
    fflush(glog);
}

// -----------------------------------------------------------
// Função auxiliar
// -----------------------------------------------------------
int get_num(const char *src, const char *cmd)
{
    const char *p = strstr(src, cmd);
    if (!p) return -1;
    p += strlen(cmd);
    return atoi(p);
}

// -----------------------------------------------------------
// Tradutor ZPL → BPLB Básico
// -----------------------------------------------------------
void traduz_zpl_para_bplb(char *zpl, FILE *out)
{
    fprintf(out, "N\n"); // limpa buffer
    log_write("Iniciando traducao ZPL -> BPLB");

    // MD - densidade
    int md = get_num(zpl, "^MD");
    if (md >= 0) {
        fprintf(out, "D%d\n", md);
        log_printf("MD detectado: %d", md);
    }

    // PW - largura
    int pw = get_num(zpl, "^PW");
    if (pw >= 0) {
        fprintf(out, "q%d\n", pw);
        log_printf("PW detectado: %d", pw);
    }

    fprintf(out, "S3\n");

    // processar campos FO/FD
    char *p = zpl;

    while ((p = strstr(p, "^FO")) != NULL)
    {
        int x = 0, y = 0;
        sscanf(p, "^FO%d,%d", &x, &y);
        log_printf("Encontrou ^FO em x=%d y=%d", x, y);

        char *fd = strstr(p, "^FD");
        if (!fd) { p++; continue; }

        fd += 3;

        char texto[1024] = {0};
        int i = 0;

        while (fd[i] && !(fd[i]=='^' && fd[i+1]=='F' && fd[i+2]=='S'))
            texto[i] = fd[i], i++;

        texto[i] = 0;

        log_printf("Texto FD extraído: '%s'", texto);

        // Barcode?
        if (strstr(p, "^BC") || strstr(p, "^BEN") || strstr(p, "^BCN")) {
            fprintf(out, "B%d,%d,0,1,2,4,60,B,\"%s\"\n", x, y, texto);
            log_write("Campo identificado como código de barras.");
        }
        // Imagem?
        else if (strstr(p, "^XG"))
        {
            char nomeimg[128]={0};
            sscanf(p, "^XG%[^,^ ]", nomeimg);
            fprintf(out, "GG%d,%d,\"%s\"\n", x, y, nomeimg);

            log_printf("Imagem detectada: %s", nomeimg);
        }
        else
        {
            // Texto normal
            fprintf(out, "A%d,%d,0,3,1,1,N,\"%s\"\n", x, y, texto);
            log_write("Campo identificado como texto.");
        }

        p++;
    }

    // PQ - quantidade
    int pq = get_num(zpl, "^PQ");
    if (pq < 1) pq = 1;

    fprintf(out, "P%d\n", pq);
    log_printf("PQ detectado: %d", pq);

    log_write("Tradução finalizada.");
}

// -----------------------------------------------------------
// MAIN — Lê COM4 (lado livre do com0com)
// -----------------------------------------------------------
int main()
{
    log_open();
    log_write("Tradutor iniciado.");
    log_write("Abrindo porta COM4...");

    HANDLE h = CreateFileA(
        "\\\\.\\COM4",
        GENERIC_READ | GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE,
        NULL,
        OPEN_EXISTING,
        0,
        NULL
    );

    if (h == INVALID_HANDLE_VALUE) {
        log_printf("Falha ao abrir COM4. Erro: %lu", GetLastError());
        printf("Erro abrindo COM4.\n");
        return 1;
    }

    log_write("COM4 aberta com sucesso.");

    // Configurar DCB
    DCB dcb = {0};
    dcb.DCBlength = sizeof(DCB);

    GetCommState(h, &dcb);
    dcb.BaudRate = CBR_115200;
    dcb.ByteSize = 8;
    dcb.StopBits = ONESTOPBIT;
    dcb.Parity   = NOPARITY;
    SetCommState(h, &dcb);

    // Timeouts
    COMMTIMEOUTS t = {0};
    t.ReadIntervalTimeout = 50;
    t.ReadTotalTimeoutConstant = 50;
    t.ReadTotalTimeoutMultiplier = 10;
    SetCommTimeouts(h, &t);

    printf("Aguardando ZPL na COM4...\n");
    log_write("Aguardando ZPL na COM4...");

    char buffer[BUF]={0};
    char zpl[BUF*4]={0};
    int pos = 0;

    DWORD lidos;

    while (1)
    {
        if (ReadFile(h, buffer, BUF-1, &lidos, NULL) && lidos > 0)
        {
            buffer[lidos] = 0;

            log_printf("Recebidos %lu bytes: '%s'", lidos, buffer);

            memcpy(zpl + pos, buffer, lidos);
            pos += lidos;

            if (strstr(zpl, "^XZ"))
            {
                log_write("ZPL completo detectado (^XZ). Iniciando tradução...");

                FILE *out = fopen("saida_bplb.txt", "w");
                traduz_zpl_para_bplb(zpl, out);
                fclose(out);

                printf("ZPL traduzido → saida_bplb.txt\n");
                log_write("Arquivo saida_bplb.txt gerado com sucesso.");

                memset(zpl, 0, sizeof(zpl));
                pos = 0;
            }
        }
        Sleep(10);
    }

    CloseHandle(h);
    return 0;
}
