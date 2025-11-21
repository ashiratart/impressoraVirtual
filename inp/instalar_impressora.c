// install_printer.c
// Compilar: cl /EHsc install_printer.c
// Funcionalidade: instala impressora "Genérica RAW" com:
// - seleção automática de porta do par com0com (prefere COM3)
// - instalação do driver "Generic / Text Only" com vários fallbacks
// - tentativa de ativar recursos de impressão (PowerShell/DISM fallback)
// - validação da porta COM com CreateFile
// - log detalhado em install_printer.log na mesma pasta do exe
// - aceita parâmetro /port=COMx para forçar porta
//
// Observações: algumas operações (instalar drivers, ativar features) dependem da versão do Windows
// e permissões/Políticas. O programa tenta várias estratégias e grava logs para diagnóstico.

#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <winspool.h>
#include <stdio.h>
#include <time.h>
#include <string.h>

#pragma comment(lib, "Winspool.lib")
#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "Shell32.lib")

// Configuráveis
const char *PRINTER_NAME = "Genérica RAW";
const char *DRIVER_NAME = "Generic / Text Only";
const char *INF_RELATIVE = "\\inf\\ntprint.inf"; // expand with %WINDIR%

// ---------- Logging ----------
static char g_logPath[MAX_PATH] = {0};
static FILE *g_log = NULL;

void log_open(const char *exePath) {
    char path[MAX_PATH];
    snprintf(path, sizeof(path), "%sinstall_printer.log", exePath);
    strncpy_s(g_logPath, sizeof(g_logPath), path, _TRUNCATE);
    g_log = fopen(g_logPath, "a");
    if (!g_log) {
        MessageBoxA(NULL, "Falha ao abrir arquivo de log.", "Erro", MB_ICONERROR);
    }
}

void log_close() {
    if (g_log) fclose(g_log);
    g_log = NULL;
}

void log_printf(const char *fmt, ...) {
    if (!g_log) return;
    time_t t = time(NULL);
    struct tm lt;
    localtime_s(&lt, &t);
    char timestr[64];
    snprintf(timestr, sizeof(timestr), "%04d-%02d-%02d %02d:%02d:%02d",
             lt.tm_year+1900, lt.tm_mon+1, lt.tm_mday, lt.tm_hour, lt.tm_min, lt.tm_sec);

    fprintf(g_log, "[%s] ", timestr);
    va_list ap;
    va_start(ap, fmt);
    vfprintf(g_log, fmt, ap);
    va_end(ap);
    fprintf(g_log, "\n");
    fflush(g_log);
}

// ---------- Utilitários ----------
void mostrarErro(const char* contexto) {
    DWORD err = GetLastError();
    LPSTR msg = NULL;
    FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                   NULL, err, 0, (LPSTR)&msg, 0, NULL);
    char buf[1024];
    snprintf(buf, sizeof(buf), "%s\n\nErro: %lu\n%s", contexto, (unsigned long)err, msg ? msg : "(sem descrição)");
    MessageBoxA(NULL, buf, "ERRO", MB_ICONERROR);
    if (msg) LocalFree(msg);

    log_printf("ERRO: %s | GetLastError=%lu | msg=%s", contexto, (unsigned long)err, msg?msg:"(sem descricao)");
}

void getExecutablePath(char* out, size_t outSize) {
    GetModuleFileNameA(NULL, out, (DWORD)outSize);
    char* last = strrchr(out, '\\');
    if (last) *(last + 1) = '\0';
}

BOOL arquivoExiste(const char* caminho) {
    DWORD a = GetFileAttributesA(caminho);
    return (a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY));
}

// ---------- UAC / Relaunch as Admin ----------
void relaunchAsAdminIfNeeded(int argc, char **argv) {
    BOOL isAdmin = FALSE;
    PSID adminGroup = NULL;
    SID_IDENTIFIER_AUTHORITY ntAuthority = SECURITY_NT_AUTHORITY;
    if (AllocateAndInitializeSid(&ntAuthority, 2,
                                 SECURITY_BUILTIN_DOMAIN_RID,
                                 DOMAIN_ALIAS_RID_ADMINS,
                                 0,0,0,0,0,0,
                                 &adminGroup)) {
        CheckTokenMembership(NULL, adminGroup, &isAdmin);
        FreeSid(adminGroup);
    }

    if (!isAdmin) {
        char path[MAX_PATH];
        GetModuleFileNameA(NULL, path, MAX_PATH);

        // Rebuild command line
        char params[2048] = {0};
        for (int i = 1; i < argc; ++i) {
            strcat_s(params, sizeof(params), argv[i]);
            if (i+1 < argc) strcat_s(params, sizeof(params), " ");
        }

        SHELLEXECUTEINFOA sei = {0};
        sei.cbSize = sizeof(sei);
        sei.lpVerb = "runas";
        sei.lpFile = path;
        sei.lpParameters = params[0] ? params : NULL;
        sei.nShow = SW_NORMAL;

        if (!ShellExecuteExA(&sei)) {
            MessageBoxA(NULL, "É necessário executar como administrador.", "Permissão", MB_ICONERROR);
            exit(1);
        }
        exit(0);
    }
}

// ---------- Registry: detect com0com / obter par ----------
BOOL com0comInstalado() {
    HKEY hKey;
    if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, "HARDWARE\\DEVICEMAP\\SERIALCOMM", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return FALSE;

    char name[256], data[256];
    DWORD nameLen, dataLen, type;
    BOOL found = FALSE;

    for (DWORD i = 0;; i++) {
        nameLen = sizeof(name);
        dataLen = sizeof(data);
        LONG res = RegEnumValueA(hKey, i, name, &nameLen, NULL, &type, (LPBYTE)data, &dataLen);
        if (res != ERROR_SUCCESS) break;
        if (strstr(name, "CNCA") || strstr(name, "CNCB")) {
            found = TRUE;
            break;
        }
    }
    RegCloseKey(hKey);
    return found;
}

BOOL obterParExistente(char* portaA, size_t tamA, char* portaB, size_t tamB) {
    if (portaA && tamA>0) portaA[0]='\0';
    if (portaB && tamB>0) portaB[0]='\0';

    HKEY hKey;
    if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, "HARDWARE\\DEVICEMAP\\SERIALCOMM", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return FALSE;

    char name[256], data[256];
    DWORD nameLen, dataLen, type;
    BOOL achouA = FALSE, achouB = FALSE;

    for (DWORD i = 0;; i++) {
        nameLen = sizeof(name);
        dataLen = sizeof(data);
        LONG res = RegEnumValueA(hKey, i, name, &nameLen, NULL, &type, (LPBYTE)data, &dataLen);
        if (res != ERROR_SUCCESS) break;
        if (!achouA && strstr(name, "CNCA")) { strncpy_s(portaA, tamA, data, _TRUNCATE); achouA = TRUE; }
        if (!achouB && strstr(name, "CNCB")) { strncpy_s(portaB, tamB, data, _TRUNCATE); achouB = TRUE; }
        if (achouA && achouB) break;
    }
    RegCloseKey(hKey);
    return (achouA && achouB);
}

BOOL portaExisteSerialcomm(const char* comName) {
    HKEY hKey;
    if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, "HARDWARE\\DEVICEMAP\\SERIALCOMM", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return FALSE;

    char name[256], data[256];
    DWORD nameLen, dataLen, type;
    BOOL found = FALSE;

    for (DWORD i = 0;; i++) {
        nameLen = sizeof(name);
        dataLen = sizeof(data);
        LONG res = RegEnumValueA(hKey, i, name, &nameLen, NULL, &type, (LPBYTE)data, &dataLen);
        if (res != ERROR_SUCCESS) break;
        if (_stricmp(data, comName) == 0) { found = TRUE; break; }
    }
    RegCloseKey(hKey);
    return found;
}

// ---------- Testa porta COM (abre com CreateFile) ----------
BOOL testarPortaCom(const char* comName) {
    char alvo[64];
    snprintf(alvo, sizeof(alvo), "\\\\.\\%s", comName);
    HANDLE h = CreateFileA(alvo, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) {
        DWORD err = GetLastError();
        log_printf("Teste CreateFile(%s) falhou. GetLastError=%lu", alvo, (unsigned long)err);
        return FALSE;
    } else {
        log_printf("Teste CreateFile(%s) OK (porta aberta com sucesso).", alvo);
        CloseHandle(h);
        return TRUE;
    }
}

// ---------- Instalação driver: PrintUIEntry, AddPrinterDriverA, pnputil, PowerShell/DISM ----------
int run_system_and_log(const char *cmd) {
    log_printf("Executando comando: %s", cmd);
    int rc = system(cmd);
    log_printf("Comando retornou: %d", rc);
    return rc;
}

// tentativa 1: PrintUIEntry sem /h /v (mais compatível)
BOOL tentarPrintUIEntry() {
    char infPath[MAX_PATH];
    if (ExpandEnvironmentStringsA("%WINDIR%\\inf\\ntprint.inf", infPath, (DWORD)sizeof(infPath)) == 0) {
        log_printf("ExpandEnvironmentStrings falhou para %%WINDIR%%\\inf\\ntprint.inf");
        return FALSE;
    }
    if (!arquivoExiste(infPath)) {
        log_printf("INF ntprint.inf nao encontrado em: %s", infPath);
        return FALSE;
    }

    char cmd[1024];
    snprintf(cmd, sizeof(cmd), "rundll32 printui.dll,PrintUIEntry /ia /m \"%s\" /f \"%s\"", DRIVER_NAME, infPath);
    int rc = run_system_and_log(cmd);
    return (rc == 0);
}

// tentativa 2: AddPrinterDriverA via API
BOOL tentarAddPrinterDriverAPI() {
    DRIVER_INFO_3A di;
    memset(&di, 0, sizeof(di));
    di.cVersion = 3;
    di.pName = (LPSTR)DRIVER_NAME;
    di.pEnvironment = "Windows x64";
    di.pDriverPath = "C:\\Windows\\System32\\ntprint.dll";
    di.pConfigFile = "UNIDRV.DLL";
    di.pDataFile = "TTFSUB.GPD";
    BOOL ok = AddPrinterDriverA(NULL, 3, (BYTE*)&di);
    if (!ok) {
        DWORD err = GetLastError();
        log_printf("AddPrinterDriverA falhou. GetLastError=%lu", (unsigned long)err);
        if (err == ERROR_PRINTER_DRIVER_ALREADY_INSTALLED) {
            log_printf("Driver já instalado (ERROR_PRINTER_DRIVER_ALREADY_INSTALLED).");
            return TRUE;
        }
        return FALSE;
    }
    log_printf("AddPrinterDriverA retornou TRUE.");
    return TRUE;
}

// tentativa 3: pnputil para adicionar INF / instalar driver via driver store (fallback)
BOOL tentarPnPUtilAdd() {
    char infPath[MAX_PATH];
    if (ExpandEnvironmentStringsA("%WINDIR%\\inf\\ntprint.inf", infPath, (DWORD)sizeof(infPath)) == 0) {
        log_printf("ExpandEnvironmentStrings falhou para %%WINDIR%%\\inf\\ntprint.inf");
        return FALSE;
    }
    if (!arquivoExiste(infPath)) {
        log_printf("INF ntprint.inf nao encontrado para pnputil.");
        return FALSE;
    }

    char cmd[1024];
    // pnputil /add-driver <inf path> /install
    snprintf(cmd, sizeof(cmd), "pnputil /add-driver \"%s\" /install", infPath);
    int rc = run_system_and_log(cmd);
    // pnputil retorna 0 em sucesso
    return (rc == 0);
}

// tentativa 4: habilitar recursos relacionados à impressão via PowerShell (Enable-WindowsOptionalFeature)
// Observação: nomes de features podem variar entre versões; fazemos tentativa genérica segura
BOOL tentarAtivarFeaturesImpressao() {
    // comandos PowerShell - tenta várias features conhecidas; executa em modo noninteractive
    const char *comandos[] = {
        // Tentativa genérica: instalar feature de impressão (pode falhar sem efeito)
        "powershell -NoProfile -Command \"Enable-WindowsOptionalFeature -Online -FeatureName Printing-Foundation-Features -All -NoRestart -ErrorAction SilentlyContinue; exit $LASTEXITCODE\"",
        // Fallback para DISM:
        "cmd.exe /c dism /online /Enable-Feature /FeatureName:Printing-Foundation-Features /All /NoRestart",
        NULL
    };

    for (int i = 0; comandos[i]; ++i) {
        int rc = run_system_and_log(comandos[i]);
        if (rc == 0) {
            log_printf("Ativacao de feature via comando[%d] retornou 0 -> OK", i);
            return TRUE;
        } else {
            log_printf("Comando de feature[%d] retornou %d", i, rc);
        }
    }
    return FALSE;
}

// ---------- criar porta no spooler (tenta AddPortA dinamicamente) ----------
BOOL criarPortaNoSpooler(const char* porta) {
    HMODULE h = LoadLibraryA("winspool.drv");
    if (!h) {
        log_printf("LoadLibrary(winspool.drv) falhou");
        return FALSE;
    }
    typedef BOOL (WINAPI *PFN_AddPortA)(LPSTR, HWND, LPSTR);
    PFN_AddPortA pAddPort = (PFN_AddPortA)GetProcAddress(h, "AddPortA");
    if (!pAddPort) { FreeLibrary(h); log_printf("GetProcAddress(AddPortA) falhou"); return FALSE; }

    BOOL res = pAddPort(NULL, NULL, (LPSTR)porta);
    FreeLibrary(h);
    log_printf("AddPortA resultado: %d", res);
    return res;
}

// ---------- criar impressora RAW ----------
BOOL criarImpressoraRAW(const char* nome, const char* porta, const char* driver) {
    PRINTER_INFO_2A pi;
    memset(&pi, 0, sizeof(pi));
    pi.pPrinterName = (LPSTR)nome;
    pi.pPortName = (LPSTR)porta;
    pi.pDriverName = (LPSTR)driver;
    pi.pPrintProcessor = "WinPrint";
    pi.pDatatype = "RAW";

    HANDLE h = AddPrinterA(NULL, 2, (BYTE*)&pi);
    if (!h) {
        DWORD err = GetLastError();
        log_printf("AddPrinterA falhou. GetLastError=%lu", (unsigned long)err);
        return FALSE;
    }
    ClosePrinter(h);
    log_printf("AddPrinterA sucesso. Impressora criada: %s", nome);
    return TRUE;
}

// ---------- Parser de argumento /port=COMx ----------
int parse_forced_port(int argc, char **argv, char *outPort, size_t outSize) {
    for (int i=1;i<argc;i++) {
        if (_strnicmp(argv[i], "/port=", 6) == 0 || _strnicmp(argv[i], "--port=", 7) == 0) {
            const char *p = strchr(argv[i], '=');
            if (p && *(p+1)) {
                strncpy_s(outPort, outSize, p+1, _TRUNCATE);
                return 1;
            }
        }
    }
    return 0;
}

// ---------- MAIN ----------
int main(int argc, char **argv) {
    relaunchAsAdminIfNeeded(argc, argv);

    char exePath[MAX_PATH];
    getExecutablePath(exePath, sizeof(exePath));
    log_open(exePath);
    log_printf("=== Início do instalador de impressora ===");

    // Argumento opcional /port=COMx
    char portaForcada[64] = {0};
    if (parse_forced_port(argc, argv, portaForcada, sizeof(portaForcada))) {
        log_printf("Porta forcada via argumento: %s", portaForcada);
    }

    // Escolha de porta: prefer COM3 do par, senão a outra do par, senão COM3/COM4 detectados, senão fallback COM3
    char portaA[64] = {0}, portaB[64] = {0}, portaEscolhida[64] = {0};

    if (portaForcada[0]) {
        strncpy_s(portaEscolhida, sizeof(portaEscolhida), portaForcada, _TRUNCATE);
        log_printf("Usando porta forcada: %s", portaEscolhida);
    } else {
        if (com0comInstalado() && obterParExistente(portaA, sizeof(portaA), portaB, sizeof(portaB))) {
            log_printf("Par com0com detectado: %s <-> %s", portaA, portaB);
            if (_stricmp(portaA, "COM3") == 0 || _stricmp(portaB, "COM3") == 0) {
                strncpy_s(portaEscolhida, sizeof(portaEscolhida), "COM3", _TRUNCATE);
            } else {
                if (portaA[0]) strncpy_s(portaEscolhida, sizeof(portaEscolhida), portaA, _TRUNCATE);
                else if (portaB[0]) strncpy_s(portaEscolhida, sizeof(portaEscolhida), portaB, _TRUNCATE);
            }
        } else {
            log_printf("com0com não detectado ou par não encontrado.");
            if (portaExisteSerialcomm("COM3")) {
                strncpy_s(portaEscolhida, sizeof(portaEscolhida), "COM3", _TRUNCATE);
                log_printf("COM3 detectado em SERIALCOMM -> usando COM3");
            } else if (portaExisteSerialcomm("COM4")) {
                strncpy_s(portaEscolhida, sizeof(portaEscolhida), "COM4", _TRUNCATE);
                log_printf("COM4 detectado em SERIALCOMM -> usando COM4");
            } else {
                strncpy_s(portaEscolhida, sizeof(portaEscolhida), "COM3", _TRUNCATE);
                log_printf("Nenhuma porta detectada; fallback para COM3");
            }
        }
    }

    char msg[512];
    snprintf(msg, sizeof(msg), "Usando porta: %s\nNome da impressora: %s", portaEscolhida, PRINTER_NAME);
    MessageBoxA(NULL, msg, "Instalar Impressora", MB_OK);
    log_printf("Porta escolhida: %s", portaEscolhida);

    // Testar porta com CreateFile (opcional; falha aqui não impede continuar)
    log_printf("Testando porta %s com CreateFile...", portaEscolhida);
    BOOL portaAberta = testarPortaCom(portaEscolhida);
    log_printf("Resultado teste porta: %d", portaAberta);

    // Tentar instalar driver em múltiplas estratégias
    log_printf("Tentando instalar driver '%s' (PrintUIEntry)...", DRIVER_NAME);
    BOOL driverOk = FALSE;
    if (tentarPrintUIEntry()) {
        log_printf("PrintUIEntry instalou driver (exit 0).");
        driverOk = TRUE;
    } else {
        log_printf("PrintUIEntry falhou; tentando AddPrinterDriverA...");
        if (tentarAddPrinterDriverAPI()) {
            log_printf("AddPrinterDriverA instalou driver.");
            driverOk = TRUE;
        } else {
            log_printf("AddPrinterDriverA falhou; tentando pnputil...");
            if (tentarPnPUtilAdd()) {
                log_printf("pnputil adicionou o driver (INF).");
                driverOk = TRUE;
            } else {
                log_printf("pnputil falhou; tentando ativar features de impressao (PowerShell/DISM)...");
                if (tentarAtivarFeaturesImpressao()) {
                    log_printf("Ativacao de features retornou sucesso; tentando PrintUI novamente...");
                    if (tentarPrintUIEntry()) {
                        log_printf("PrintUIEntry após ativar features retornou 0");
                        driverOk = TRUE;
                    }
                }
            }
        }
    }

    if (!driverOk) {
        MessageBoxA(NULL,
                   "Falha ao instalar driver 'Generic / Text Only'. O instalador irá tentar criar a impressora, mas pode falhar com erro 1797.\n"
                   "Verifique o log (install_printer.log) na pasta do executável para detalhes.",
                   "Aviso", MB_ICONWARNING);
        log_printf("Driver nao instalado apos todas as tentativas.");
    } else {
        MessageBoxA(NULL, "Driver 'Generic / Text Only' instalado (ou já presente). Prosseguindo...", "Info", MB_OK);
    }

    // Tentar criar porta no spooler (não-fatal)
    log_printf("Tentando AddPortA para %s (não-fatal)", portaEscolhida);
    if (!criarPortaNoSpooler(portaEscolhida)) {
        log_printf("AddPortA retornou falha (esperado em muitos casos).");
    } else {
        log_printf("AddPortA sucesso.");
    }

    // Criar impressora RAW
    log_printf("Tentando criar impressora RAW: %s on %s", PRINTER_NAME, portaEscolhida);
    if (!criarImpressoraRAW(PRINTER_NAME, portaEscolhida, DRIVER_NAME)) {
        // Falha - registrar e mostrar erro
        DWORD err = GetLastError();
        log_printf("AddPrinterA falhou. GetLastError=%lu", (unsigned long)err);
        mostrarErro("Falha ao criar impressora (AddPrinterA). Possível causa: driver ausente (erro 1797) ou parâmetros inválidos.");
        log_printf("Instalação abortada devido a falha em AddPrinterA.");
        log_close();
        return 1;
    }

    log_printf("Impressora criada com sucesso: %s", PRINTER_NAME);
    MessageBoxA(NULL, "Impressora criada com sucesso!", "Sucesso", MB_OK);

    log_printf("=== Fim do instalador de impressora ===");
    log_close();
    return 0;
}
