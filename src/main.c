#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <time.h>
#include <stdarg.h>
#include <stdlib.h>

#define BUF 4096

FILE *glog = NULL;
char* impressora_selecionada = NULL;

// Variáveis para medir tempo de execução
clock_t tempo_inicio;
clock_t tempo_recebimento_zpl;
clock_t tempo_traducao;
clock_t tempo_impressao;

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
// Funções para medir tempo
// -----------------------------------------------------------
void iniciar_tempo()
{
    tempo_inicio = clock();
    log_write("=== INICIO DA EXECUCAO ===");
    log_printf("Tempo inicial: %lu ms", tempo_inicio);
}

void log_tempo_recebimento_zpl()
{
    tempo_recebimento_zpl = clock();
    double tempo_decorrido = ((double)(tempo_recebimento_zpl - tempo_inicio)) / CLOCKS_PER_SEC * 1000;
    log_printf("ZPL recebido apos: %.2f ms", tempo_decorrido);
}

void log_tempo_traducao()
{
    tempo_traducao = clock();
    double tempo_decorrido = ((double)(tempo_traducao - tempo_recebimento_zpl)) / CLOCKS_PER_SEC * 1000;
    log_printf("Traducao concluida em: %.2f ms", tempo_decorrido);
}

void log_tempo_impressao()
{
    tempo_impressao = clock();
    
    double tempo_total = ((double)(tempo_impressao - tempo_inicio)) / CLOCKS_PER_SEC * 1000;
    double tempo_traducao_only = ((double)(tempo_traducao - tempo_recebimento_zpl)) / CLOCKS_PER_SEC * 1000;
    double tempo_impressao_only = ((double)(tempo_impressao - tempo_traducao)) / CLOCKS_PER_SEC * 1000;
    
    log_printf("=== TEMPOS DE EXECUCAO ===");
    log_printf("Tempo total: %.2f ms", tempo_total);
    log_printf("Tempo traducao: %.2f ms", tempo_traducao_only);
    log_printf("Tempo impressao: %.2f ms", tempo_impressao_only);
    log_printf("Tempo desde inicio: %.2f ms", tempo_total);
    log_printf("=== FIM DA EXECUCAO ===");
}

void log_tempo_apenas_arquivo()
{
    clock_t tempo_final = clock();
    double tempo_total = ((double)(tempo_final - tempo_inicio)) / CLOCKS_PER_SEC * 1000;
    double tempo_traducao_only = ((double)(tempo_traducao - tempo_recebimento_zpl)) / CLOCKS_PER_SEC * 1000;
    
    log_printf("=== TEMPOS DE EXECUCAO (APENAS ARQUIVO) ===");
    log_printf("Tempo total: %.2f ms", tempo_total);
    log_printf("Tempo traducao: %.2f ms", tempo_traducao_only);
    log_printf("Tempo desde inicio: %.2f ms", tempo_total);
    log_printf("=== FIM DA EXECUCAO ===");
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
// Listar todas as impressoras disponíveis
// -----------------------------------------------------------
char** listar_impressoras(int* count)
{
    char** impressoras = NULL;
    *count = 0;
    
    DWORD needed = 0, returned = 0;
    
    // Primeiro obtemos o tamanho necessário
    EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, 
                 NULL, 0, &needed, &returned);
    
    if (needed > 0 && returned > 0) {
        BYTE* printer_info = (BYTE*)malloc(needed);
        
        if (EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, 
                        printer_info, needed, &needed, &returned)) {
            
            PRINTER_INFO_2* printers = (PRINTER_INFO_2*)printer_info;
            
            // Alocar array para os nomes
            impressoras = (char**)malloc(returned * sizeof(char*));
            
            for (DWORD i = 0; i < returned; i++) {
                if (printers[i].pPrinterName) {
                    impressoras[i] = _strdup(printers[i].pPrinterName);
                    (*count)++;
                }
            }
        }
        free(printer_info);
    }
    
    return impressoras;
}

// -----------------------------------------------------------
// Mostrar menu de seleção de impressora
// -----------------------------------------------------------
char* selecionar_impressora_menu()
{
    int count = 0;
    char** impressoras = listar_impressoras(&count);
    
    if (count == 0) {
        printf("Nenhuma impressora encontrada!\n");
        log_write("Nenhuma impressora encontrada no sistema");
        return NULL;
    }
    
    printf("\n=== IMPRESSORAS DISPONIVEIS ===\n");
    for (int i = 0; i < count; i++) {
        printf("[%d] - %s\n", i + 1, impressoras[i]);
    }
    printf("[0] - Nenhuma (apenas salvar arquivo)\n");
    printf("\nSelecione a impressora: ");
    
    int escolha;
    if (scanf("%d", &escolha) != 1) {
        // Limpar buffer de entrada em caso de erro
        int c;
        while ((c = getchar()) != '\n' && c != EOF);
        printf("Entrada invalida!\n");
        
        // Liberar memória
        for (int i = 0; i < count; i++) free(impressoras[i]);
        free(impressoras);
        return NULL;
    }
    
    char* selecionada = NULL;
    
    if (escolha == 0) {
        printf("Modo apenas arquivo selecionado.\n");
        log_write("Usuario selecionou: Nenhuma impressora (apenas arquivo)");
    }
    else if (escolha >= 1 && escolha <= count) {
        selecionada = _strdup(impressoras[escolha - 1]);
        printf("Impressora selecionada: %s\n", selecionada);
        log_printf("Usuario selecionou impressora: %s", selecionada);
    } else {
        printf("Opcao invalida!\n");
        log_printf("Usuario selecionou opcao invalida: %d", escolha);
    }
    
    // Liberar memória
    for (int i = 0; i < count; i++) free(impressoras[i]);
    free(impressoras);
    
    return selecionada;
}

// -----------------------------------------------------------
// Menu principal de configuração
// -----------------------------------------------------------
void menu_configuracao()
{
    int opcao;
    
    do {
        printf("\n=== CONFIGURACAO DA IMPRESSORA ===\n");
        if (impressora_selecionada) {
            printf("Impressora atual: %s\n", impressora_selecionada);
        } else {
            printf("Impressora atual: NENHUMA (apenas salvar arquivo)\n");
        }
        printf("\n[1] - Selecionar/alterar impressora\n");
        printf("[2] - Remover impressora selecionada\n");
        printf("[3] - Iniciar monitoramento COM4\n");
        printf("[4] - Sair\n");
        printf("\nEscolha: ");
        
        if (scanf("%d", &opcao) != 1) {
            // Limpar buffer de entrada em caso de erro
            int c;
            while ((c = getchar()) != '\n' && c != EOF);
            printf("Entrada invalida!\n");
            continue;
        }
        
        switch (opcao) {
            case 1: {
                if (impressora_selecionada) {
                    free(impressora_selecionada);
                    impressora_selecionada = NULL;
                }
                char* nova = selecionar_impressora_menu();
                if (nova) {
                    impressora_selecionada = nova;
                }
                break;
            }
            case 2:
                if (impressora_selecionada) {
                    printf("Impressora '%s' removida.\n", impressora_selecionada);
                    log_printf("Usuario removeu impressora: %s", impressora_selecionada);
                    free(impressora_selecionada);
                    impressora_selecionada = NULL;
                } else {
                    printf("Nenhuma impressora selecionada.\n");
                }
                break;
            case 3:
                printf("Iniciando monitoramento...\n");
                return; // Sai do menu e inicia o monitoramento
            case 4:
                log_write("Programa finalizado pelo usuario");
                exit(0);
            default:
                printf("Opcao invalida!\n");
        }
    } while (1);
}

// -----------------------------------------------------------
// Enviar arquivo para impressora
// -----------------------------------------------------------
BOOL enviar_para_impressora(const char* nome_arquivo, const char* nome_impressora)
{
    log_printf("Tentando enviar %s para impressora %s", nome_arquivo, nome_impressora);
    
    // Ler conteúdo do arquivo
    FILE* arquivo = fopen(nome_arquivo, "rb");
    if (!arquivo) {
        log_printf("Erro ao abrir arquivo %s", nome_arquivo);
        return FALSE;
    }
    
    fseek(arquivo, 0, SEEK_END);
    long tamanho = ftell(arquivo);
    fseek(arquivo, 0, SEEK_SET);
    
    char* dados = (char*)malloc(tamanho + 1);
    if (!dados) {
        fclose(arquivo);
        return FALSE;
    }
    
    fread(dados, 1, tamanho, arquivo);
    dados[tamanho] = '\0';
    fclose(arquivo);
    
    // Enviar para impressora
    HANDLE hPrinter;
    DOC_INFO_1 docInfo;
    DWORD bytesWritten;
    
    // Abrir impressora
    if (!OpenPrinterA((LPSTR)nome_impressora, &hPrinter, NULL)) {
        log_printf("Erro ao abrir impressora. Código: %lu", GetLastError());
        free(dados);
        return FALSE;
    }
    
    // Configurar documento
    docInfo.pDocName = "Etiqueta Traduzida";
    docInfo.pOutputFile = NULL;
    docInfo.pDatatype = "RAW";
    
    // Iniciar documento
    if (!StartDocPrinter(hPrinter, 1, (LPBYTE)&docInfo)) {
        log_printf("Erro ao iniciar documento na impressora");
        ClosePrinter(hPrinter);
        free(dados);
        return FALSE;
    }
    
    // Iniciar página
    if (!StartPagePrinter(hPrinter)) {
        log_printf("Erro ao iniciar página na impressora");
        EndDocPrinter(hPrinter);
        ClosePrinter(hPrinter);
        free(dados);
        return FALSE;
    }
    
    // Escrever dados
    if (!WritePrinter(hPrinter, dados, tamanho, &bytesWritten)) {
        log_printf("Erro ao escrever na impressora");
        EndPagePrinter(hPrinter);
        EndDocPrinter(hPrinter);
        ClosePrinter(hPrinter);
        free(dados);
        return FALSE;
    }
    
    // Finalizar
    EndPagePrinter(hPrinter);
    EndDocPrinter(hPrinter);
    ClosePrinter(hPrinter);
    free(dados);
    
    log_printf("Arquivo enviado com sucesso para a impressora. Bytes escritos: %lu", bytesWritten);
    return TRUE;
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
    
    // Iniciar contagem de tempo
    iniciar_tempo();
    
    printf("=== TRADUTOR ZPL para BPLB ===\n");
    
    // Mostrar menu de configuração inicial
    menu_configuracao();

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

    printf("\nAguardando ZPL na COM4...\n");
    printf("Impressora selecionada: %s\n", impressora_selecionada ? impressora_selecionada : "NENHUMA (apenas arquivo)");
    printf("Pressione Ctrl+C para parar\n\n");
    
    log_write("Iniciando monitoramento da COM4");
    if (impressora_selecionada) {
        log_printf("Impressora configurada: %s", impressora_selecionada);
    } else {
        log_write("Nenhuma impressora configurada - modo apenas arquivo");
    }

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
                // Marcar tempo de recebimento do ZPL completo
                log_tempo_recebimento_zpl();
                
                log_write("ZPL completo detectado (^XZ). Iniciando tradução...");

                FILE *out = fopen("saida_bplb.txt", "w");
                traduz_zpl_para_bplb(zpl, out);
                fclose(out);

                // Marcar tempo de conclusão da tradução
                log_tempo_traducao();

                printf("ZPL traduzido → saida_bplb.txt\n");
                log_write("Arquivo saida_bplb.txt gerado com sucesso.");

                // Enviar para impressora se configurada
                if (impressora_selecionada) {
                    if (enviar_para_impressora("saida_bplb.txt", impressora_selecionada)) {
                        log_write("Arquivo enviado com sucesso para a impressora");
                        printf("Arquivo enviado para: %s\n", impressora_selecionada);
                        
                        // Marcar tempo de conclusão da impressão
                        log_tempo_impressao();
                    } else {
                        log_write("Falha ao enviar para impressora");
                        printf("Erro ao enviar para impressora!\n");
                    }
                } else {
                    printf("Arquivo salvo (nenhuma impressora selecionada)\n");
                    log_tempo_apenas_arquivo();
                }

                memset(zpl, 0, sizeof(zpl));
                pos = 0;
            }
        }
        Sleep(10);
    }

    if (impressora_selecionada) free(impressora_selecionada);
    CloseHandle(h);
    return 0;
}