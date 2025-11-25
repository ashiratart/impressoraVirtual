#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <time.h>
#include <stdarg.h>
#include <stdlib.h>
#include "bplb.h"
#include "zpl.h"  // Inclui o novo header

#define BUF 4096

FILE *glog = NULL;
char* impressora_selecionada = NULL;

// Variáveis para medir tempo de execução
clock_t tempo_inicio;
clock_t tempo_recebimento_zpl;
clock_t tempo_traducao;
clock_t tempo_impressao;

// ------------------------------------------------------------------
// LOG (mantido igual)
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
// Funções para medir tempo (mantidas iguais)
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
// Função auxiliar (mantida para compatibilidade)
// -----------------------------------------------------------
int get_num(const char *src, const char *cmd)
{
    const char *p = strstr(src, cmd);
    if (!p) return -1;
    p += strlen(cmd);
    return atoi(p);
}

// -----------------------------------------------------------
// CONFIGURAÇÃO GLOBAL COM CABEÇALHO FIXO
// -----------------------------------------------------------
void process_global_settings_bplb(const char *zpl, FILE *out) {
    // CABEÇALHO FIXO - SEMPRE usar estas configurações
    bplb_clear_buffer(out);          // N
    bplb_set_density(out, 15);       // D12
    bplb_set_speed(out, 3);          // S3
    fprintf(out, "JF" BPLB_LF);      // JF
    bplb_set_label_width(out, 600);  // q600
    bplb_set_label_size(out, 260, 24); // Q260,24
    
    log_write("Configuracao fixa aplicada: N, D12, S3, JF, q600, Q260,24");
}

// -----------------------------------------------------------
// TRADUTOR ZPL2 → BPLB USANDO O NOVO HEADER (CORRIGIDO)
// -----------------------------------------------------------
void traduz_zpl_para_bplb_avancado(char *zpl, FILE *out) {
    log_write("Iniciando tradução ZPL2 -> BPLB com biblioteca avançada");
    
    // Processa configurações globais
    process_global_settings_bplb(zpl, out);
    
    // Processa ^XA (início do formato)
    if (strstr(zpl, "^XA")) {
        fprintf(out, "// Início do formato ZPL\n");
    }
    
    // Divide o ZPL em campos (separados por ^FS)
    char *field_start = zpl;
    char *field_end;
    int field_count = 0;
    
    while ((field_end = strstr(field_start, "^FS")) != NULL) {
        // Calcula tamanho do campo atual
        int field_len = field_end - field_start;
        if (field_len <= 0) {
            field_start = field_end + 3;
            continue;
        }
        
        // Copia campo para buffer temporário
        char field_buffer[1024] = {0};
        strncpy(field_buffer, field_start, field_len);
        field_buffer[field_len] = '\0';
        
        log_printf("Processando campo ZPL: %s", field_buffer);
        
        // Extrai coordenadas FO (se existir)
        int x = 0, y = 0;
        char *fo_pos = strstr(field_buffer, "^FO");
        if (fo_pos) {
            sscanf(fo_pos, "^FO%d,%d", &x, &y);
        }
        
        // ==================================================
        // CLASSIFICA E PROCESSA O CAMPO (VERSÃO CORRIGIDA)
        // ==================================================
        
        // 1. PRIMEIRO VERIFICA SE É CÓDIGO DE BARRAS (mesmo sem FD)
        if (fo_pos && (strstr(field_buffer, "^BC") || strstr(field_buffer, "^BEN") || 
                      strstr(field_buffer, "^BCN") || strstr(field_buffer, "^BE"))) {
                    
            char texto[256] = {0};
            int altura = 60; // padrão
                    
            // Extrai altura do código de barras
            if (strstr(field_buffer, "^BEN")) {
                char *ben_pos = strstr(field_buffer, "^BEN");
                char orientation;
                if (sscanf(ben_pos, "^BEN%c,%d", &orientation, &altura) == 2) {
                    // altura extraída
                }
            } else if (strstr(field_buffer, "^BC")) {
                char *bc_pos = strstr(field_buffer, "^BC");
                char orientation;
                if (sscanf(bc_pos, "^BC%c,%d", &orientation, &altura) == 2) {
                    // altura extraída
                }
            }

            // CORREÇÃO: Extração mais robusta para EAN
            char *fd_pos = strstr(field_buffer, "^FD");
            if (fd_pos) {
                fd_pos += 3; // Pula "^FD"

                // PARA EAN: corta em 14 caracteres (EAN-13 + dígito)
                if (strstr(field_buffer, "^BEN") || strstr(field_buffer, "^BE")) {
                    // É EAN - pega apenas 14 caracteres
                    int i = 0;
                    while (i < 14 && fd_pos[i] && fd_pos[i] != '^' && fd_pos[i] != '\'') {
                        texto[i] = fd_pos[i];
                        i++;
                    }
                    texto[i] = '\0';
                    log_printf("EAN extraído (14 chars): '%s'", texto);
                } else {
                    // É outro tipo de código - pega até o próximo comando
                    int i = 0;
                    while (fd_pos[i] && fd_pos[i] != '^' && i < 255) {
                        texto[i] = fd_pos[i];
                        i++;
                    }
                    texto[i] = '\0';
                    log_printf("Código extraído: '%s'", texto);
                }
            }

            // Se não encontrou no FD, tenta SN
            if (strlen(texto) == 0) {
                char *sn_pos = strstr(field_buffer, "^SN");
                if (sn_pos) {
                    sn_pos += 3;
                    int i = 0;
                    while (sn_pos[i] && sn_pos[i] != ',' && sn_pos[i] != '^' && i < 255) {
                        texto[i] = sn_pos[i];
                        i++;
                    }
                    texto[i] = '\0';
                    log_printf("Dados do SN: '%s'", texto);
                }
            }

            if (strlen(texto) > 0) {
                // AJUSTE DE POSIÇÃO
                int adjusted_x = x;
                if (x > 300) {
                    adjusted_x = 200;
                    log_printf("Ajustando posicao: %d -> %d", x, adjusted_x);
                }

                fprintf(out, "B%d,%d,0,1,2,4,%d,B,\"%s\"\n", adjusted_x, y, altura, texto);
                log_printf("Codigo barras: [%d,%d] h=%d '%s'", adjusted_x, y, altura, texto);
            }
        }
        // 2. DEPOIS VERIFICA SE É TEXTO
        else if (fo_pos && strstr(field_buffer, "^FD")) {
            
            char texto[256] = {0};
            char *fd_pos = strstr(field_buffer, "^FD");
            if (fd_pos) {
                fd_pos += 3;
                int i = 0;
                while (fd_pos[i] && fd_pos[i] != '^' && i < 255) {
                    texto[i] = fd_pos[i];
                    i++;
                }
                texto[i] = '\0';
            }
            
            if (strlen(texto) > 0 && strlen(texto) < 200) { // texto válido
                char font = '0';
                char *a_pos = strstr(field_buffer, "^A");
                if (a_pos) {
                    char orientation;
                    int height, width;
                    if (sscanf(a_pos, "^A%c%c,%d,%d", &font, &orientation, &height, &width) < 2) {
                        sscanf(a_pos, "^A%c", &font);
                    }
                }
                
                // Mapeamento de fonte ZPL para BPLB
                char bplb_font = '3';
                switch(font) {
                    case '0': bplb_font = '1'; break;
                    case '1': bplb_font = '2'; break;  
                    case '2': bplb_font = '3'; break;
                    default: bplb_font = '3'; break;
                }
                
                fprintf(out, "A%d,%d,0,%c,1,1,N,\"%s\"\n", x, y, bplb_font, texto);
                log_printf("Texto: [%d,%d] fonte=%c '%s'", x, y, bplb_font, texto);
            }
        }
        else if (fo_pos && strstr(field_buffer, "^GB")) {
            // CAIXA GRÁFICA
            int width = 100, height = 50, thickness = 2;
            char *gb_pos = strstr(field_buffer, "^GB");
            if (gb_pos) {
                sscanf(gb_pos, "^GB%d,%d,%d", &width, &height, &thickness);
            }
            convert_GB_to_BPLB(x, y, width, height, thickness, out);
            log_printf("Caixa: [%d,%d] %dx%d esp=%d", x, y, width, height, thickness);
        }
        else if (fo_pos && strstr(field_buffer, "^XG")) {
        // IMAGEM .GRF - Comando ZPL: ^XGR:NAME.GRF
            char image_name[128] = {0};
            char *xg_pos = strstr(field_buffer, "^XG");
            if (xg_pos) {
                xg_pos += 3; // Pula "^XG"
                int i = 0;
                
                // Pula o "R:" se existir (é parte do comando ZPL, não do nome)
                if (xg_pos[0] == 'R' && xg_pos[1] == ':') {
                    xg_pos += 2; // Pula "R:"
                }
                
                // Extrai o nome do arquivo (até .GRF ou próximo comando)
                while (xg_pos[i] && xg_pos[i] != ',' && xg_pos[i] != '^' && 
                       xg_pos[i] != ' ' && i < 127) {
                    image_name[i] = xg_pos[i];
                    i++;
                }
                image_name[i] = '\0';
                
                // Remove a extensão .GRF se existir (o comando GG não precisa)
                char *dot = strrchr(image_name, '.');
                if (dot) *dot = '\0';
                
                if (strlen(image_name) > 0) {
                    // Comando BPLB para imagem: GGx,y,"nome"
                    fprintf(out, "GG%d,%d,\"%s\"\n", x, y, image_name);
                    log_printf("Imagem GRF: [%d,%d] '%s'", x, y, image_name);
                }
            }
        }
        else if (strstr(field_buffer, "^FX")) {
            // COMENTÁRIO
            fprintf(out, "// %s\n", field_buffer + 3);
        }
        
        field_count++;
        field_start = field_end + 3; // Pula ^FS
    }
    
    // Processa quantidade de cópias (PQ)
    int pq = get_num(zpl, "^PQ");
    if (pq < 1) pq = 1;
    
    // Comando de impressão
    fprintf(out, "P%d,1\n", pq);
    
    // Processa ^XZ (fim do formato)
    if (strstr(zpl, "^XZ")) {
        fprintf(out, "// Fim do formato ZPL\n");
    }
    
    log_printf("Tradução BPLB concluída. %d campos processados.", field_count);
}

// -----------------------------------------------------------
// Tradutor ZPL → BPLB Básico (mantido para compatibilidade)
// -----------------------------------------------------------
void traduz_zpl_para_bplb(char *zpl, FILE *out)
{
    // Usa a versão avançada por padrão
    traduz_zpl_para_bplb_avancado(zpl, out);
}

// -----------------------------------------------------------
// Verificar permissões do sistema
// -----------------------------------------------------------
void verificar_permissoes()
{
    log_write("=== VERIFICACAO DE PERMISSOES ===");
    
    HANDLE hToken;
    if (OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken)) {
        log_write("Processo tem acesso ao token de segurança");
        CloseHandle(hToken);
    } else {
        log_printf("Erro ao abrir token: %lu", GetLastError());
    }
    
    // Tentar acessar o registro de impressoras
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\\CurrentControlSet\\Control\\Print\\Printers", 
                     0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        log_write("Acesso ao registro de impressoras: OK");
        RegCloseKey(hKey);
    } else {
        log_write("Acesso ao registro de impressoras: FALHA");
    }
}

// -----------------------------------------------------------
// Função alternativa para detectar impressoras
// -----------------------------------------------------------
void detectar_impressoras_alternativo()
{
    log_write("=== DETECÇÃO ALTERNATIVA DE IMPRESSORAS ===");
    
    // Método 1: Usando EnumPrinters com diferentes flags
    DWORD flags[] = {
        PRINTER_ENUM_LOCAL,
        PRINTER_ENUM_CONNECTIONS,
        PRINTER_ENUM_NETWORK,
        PRINTER_ENUM_REMOTE,
        PRINTER_ENUM_SHARED
    };
    
    const char* flag_names[] = {
        "LOCAL",
        "CONNECTIONS", 
        "NETWORK",
        "REMOTE",
        "SHARED"
    };
    
    for (int i = 0; i < 5; i++) {
        DWORD needed = 0, returned = 0;
        
        if (EnumPrinters(flags[i], NULL, 2, NULL, 0, &needed, &returned)) {
            log_printf("Flag %s: %lu impressoras", flag_names[i], returned);
        } else {
            DWORD error = GetLastError();
            if (error == ERROR_INSUFFICIENT_BUFFER) {
                log_printf("Flag %s: %lu impressoras (buffer insuficiente)", flag_names[i], returned);
            } else {
                log_printf("Flag %s: erro %lu", flag_names[i], error);
            }
        }
    }
}

// -----------------------------------------------------------
// Listar todas as impressoras disponíveis (VERSÃO MELHORADA)
// -----------------------------------------------------------
char** listar_impressoras(int* count)
{
    char** impressoras = NULL;
    *count = 0;
    
    DWORD needed = 0, returned = 0;
    
    log_write("Iniciando enumeração de impressoras...");
    
    // Primeiro tentamos com PRINTER_ENUM_NAME
    if (!EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, 
                     NULL, 0, &needed, &returned)) {
        DWORD error = GetLastError();
        if (error != ERROR_INSUFFICIENT_BUFFER) {
            log_printf("Erro no EnumPrinters (1): %lu", error);
            return NULL;
        }
    }
    
    if (needed > 0) {
        BYTE* printer_info = (BYTE*)malloc(needed);
        
        if (EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, 
                        printer_info, needed, &needed, &returned)) {
            
            log_printf("Encontradas %lu impressoras no sistema", returned);
            
            if (returned > 0) {
                PRINTER_INFO_2* printers = (PRINTER_INFO_2*)printer_info;
                
                // Alocar array para os nomes
                impressoras = (char**)malloc(returned * sizeof(char*));
                
                for (DWORD i = 0; i < returned; i++) {
                    if (printers[i].pPrinterName) {
                        impressoras[i] = _strdup(printers[i].pPrinterName);
                        (*count)++;
                        log_printf("Impressora %d: %s", i + 1, printers[i].pPrinterName);
                        
                        // Log adicional para debug
                        if (printers[i].pDriverName) {
                            log_printf("  Driver: %s", printers[i].pDriverName);
                        }
                        if (printers[i].pPortName) {
                            log_printf("  Porta: %s", printers[i].pPortName);
                        }
                    }
                }
            } else {
                log_write("Nenhuma impressora encontrada no sistema");
            }
        } else {
            DWORD error = GetLastError();
            log_printf("Erro no EnumPrinters (2): %lu", error);
        }
        
        free(printer_info);
    } else {
        log_write("Nenhuma impressora encontrada (buffer vazio)");
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
// Menu principal de configuração (VERSÃO MELHORADA)
// -----------------------------------------------------------
void menu_configuracao()
{
    int opcao;
    
    // Executar detecção alternativa no início
    detectar_impressoras_alternativo();
    
    do {
        printf("\n=== CONFIGURACAO DA IMPRESSORA ===\n");
        if (impressora_selecionada) {
            printf("Impressora atual: %s\n", impressora_selecionada);
        } else {
            printf("Impressora atual: NENHUMA (apenas salvar arquivo)\n");
        }
        printf("\n[1] - Selecionar/alterar impressora\n");
        printf("[2] - Remover impressora selecionada\n");
        printf("[3] - Listar impressoras (debug)\n");
        printf("[4] - Iniciar monitoramento COM4\n");
        printf("[5] - Sair\n");
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
                printf("\n=== LISTAGEM DE IMPRESSORAS (DEBUG) ===\n");
                detectar_impressoras_alternativo();
                printf("Verifique o arquivo tradutor.log para detalhes.\n");
                break;
            case 4:
                printf("Iniciando monitoramento...\n");
                return; // Sai do menu e inicia o monitoramento
            case 5:
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
// MAIN — Lê COM4 (lado livre do com0com)
// -----------------------------------------------------------
int main()
{
    log_open();
    log_write("Tradutor ZPL2->BPLB Avançado iniciado.");
    
    // Verificar permissões do sistema
    verificar_permissoes();
    
    // Iniciar contagem de tempo
    iniciar_tempo();
    
    printf("=== TRADUTOR ZPL para BPLB (AVANÇADO) ===\n");
    printf("Suporte a: Textos, Code128, Interleaved 2of5, EAN, Caixas, Imagens\n\n");
    
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
                traduz_zpl_para_bplb_avancado(zpl, out);
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