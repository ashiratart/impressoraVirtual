// install_com0com.c
// Compilar: cl /EHsc install_com0com.c
// Propósito: instalar apenas o com0com (procura com0com.zip na mesma pasta do exe e chama setupc.exe).
// Não cria impressoras nem altera spooler.

#include <windows.h>
#include <stdio.h>
#include <string.h>

#pragma comment(lib, "Advapi32.lib")   // para verificação de token/Grupo Administradores
#pragma comment(lib, "Shell32.lib")

// ---------- utilitários ----------
void mostrarErro(const char* contexto) {
    DWORD err = GetLastError();
    LPSTR msg = NULL;
    FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                   NULL, err, 0, (LPSTR)&msg, 0, NULL);
    char buffer[1024];
    snprintf(buffer, sizeof(buffer), "%s\n\nErro: %lu\n%s", contexto, err, msg ? msg : "(sem descrição)");
    MessageBoxA(NULL, buffer, "ERRO", MB_ICONERROR);
    if (msg) LocalFree(msg);
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

// ---------- elevação / UAC ----------
void relaunchAsAdmin() {
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

        SHELLEXECUTEINFOA sei = {0};
        sei.cbSize = sizeof(sei);
        sei.lpVerb = "runas";
        sei.lpFile = path;
        sei.nShow = SW_NORMAL;

        if (!ShellExecuteExA(&sei)) {
            MessageBoxA(NULL, "É necessário executar como administrador.", "Permissão", MB_ICONERROR);
            exit(1);
        }
        exit(0);
    }
}

// ---------- detecção com0com (verifica entradas CNCA/CNCB em SERIALCOMM) ----------
BOOL com0comInstalado() {
    HKEY hKey;
    LONG r = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "HARDWARE\\DEVICEMAP\\SERIALCOMM", 0, KEY_READ, &hKey);
    if (r != ERROR_SUCCESS) return FALSE;

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

// ---------- localizar setupc.exe na pasta extraída ----------
BOOL executarSetupcComPortas(const char* caminhoSetupc, const char* porta1, const char* porta2) {
    if (!arquivoExiste(caminhoSetupc)) return FALSE;

    // tenta execução simples
    char comando[1024];
    snprintf(comando, sizeof(comando), "cmd.exe /c \"%s\" install PortName=%s PortName=%s", caminhoSetupc, porta1, porta2);
    int ret = system(comando);
    if (ret == 0) return TRUE;

    // tenta via ShellExecuteEx com runas (elevação)
    char params[256];
    snprintf(params, sizeof(params), "install PortName=%s PortName=%s", porta1, porta2);

    SHELLEXECUTEINFOA sei = {0};
    sei.cbSize = sizeof(sei);
    sei.lpVerb = "runas";
    sei.lpFile = caminhoSetupc;
    sei.lpParameters = params;
    sei.nShow = SW_NORMAL;
    if (ShellExecuteExA(&sei)) return TRUE;

    return FALSE;
}

// ---------- extrair ZIP (tenta PowerShell Expand-Archive ou tar) ----------
BOOL extrairZip(const char* arquivoZip, const char* pastaDestino) {
    CreateDirectoryA(pastaDestino, NULL);

    char cmd[1024];
    char msg[512];

    // PowerShell Expand-Archive
    snprintf(cmd, sizeof(cmd),
             "powershell -NoProfile -Command \"if (Test-Path '%s') { Expand-Archive -LiteralPath '%s' -DestinationPath '%s' -Force } else { exit 1 }\"",
             arquivoZip, arquivoZip, pastaDestino);
    snprintf(msg, sizeof(msg), "Executando: %s", cmd);
    MessageBoxA(NULL, msg, "Extração", MB_OK);
    if (system(cmd) == 0) return TRUE;

    // tar fallback
    snprintf(cmd, sizeof(cmd), "tar -xf \"%s\" -C \"%s\"", arquivoZip, pastaDestino);
    if (system(cmd) == 0) return TRUE;

    return FALSE;
}

// ---------- main: só instala com0com ----------
int main(void) {
    relaunchAsAdmin();

    char exePath[MAX_PATH];
    getExecutablePath(exePath, sizeof(exePath));

    char intro[1024];
    snprintf(intro, sizeof(intro),
             "Instalador com0com (somente com0com)\n\nPasta do executável:\n%s\n\nEste programa:\n- Verifica se com0com já está instalado\n- Se não estiver, procura com0com.zip na pasta do exe\n- Extrai e tenta executar setupc.exe para instalar um par (COM5<->COM6)\n\nClique OK para continuar.", exePath);
    MessageBoxA(NULL, intro, "Instalar com0com", MB_OK);

    if (com0comInstalado()) {
        MessageBoxA(NULL, "com0com já detectado no sistema. Nenhuma ação necessária.", "Info", MB_ICONINFORMATION);
        return 0;
    }

    // procura com0com.zip na mesma pasta do exe
    char caminhoZip[MAX_PATH];
    snprintf(caminhoZip, sizeof(caminhoZip), "%scom0com.zip", exePath);
    if (!arquivoExiste(caminhoZip)) {
        char m[512];
        snprintf(m, sizeof(m),
                 "Arquivo com0com.zip não encontrado em:\n%s\n\nColoque com0com.zip na mesma pasta do executável e execute novamente.", caminhoZip);
        MessageBoxA(NULL, m, "Arquivo ausente", MB_ICONERROR);
        return 1;
    }

    // extrai
    char pastaDestino[MAX_PATH];
    snprintf(pastaDestino, sizeof(pastaDestino), "%scom0com_extracted\\", exePath);
    if (!extrairZip(caminhoZip, pastaDestino)) {
        MessageBoxA(NULL, "Falha ao extrair com0com.zip automaticamente. Extraia manualmente e execute setupc.exe como administrador.", "Erro", MB_ICONERROR);
        return 1;
    }

    // tenta localizar setupc.exe em caminhos comuns dentro de pastaDestino
    const char* possiveis[] = {
        "com0com-3.0.0.0-i386-and-x64-signed\\setupc.exe",
        "com0com\\setupc.exe",
        "setupc.exe",
        NULL
    };

    BOOL instalado = FALSE;
    for (int i = 0; possiveis[i]; ++i) {
        char caminhoFull[MAX_PATH];
        snprintf(caminhoFull, sizeof(caminhoFull), "%s%s", pastaDestino, possiveis[i]);
        if (arquivoExiste(caminhoFull)) {
            if (executarSetupcComPortas(caminhoFull, "COM5", "COM6")) {
                MessageBoxA(NULL, "com0com instalado com sucesso (tentativa: COM5 <-> COM6).", "Sucesso", MB_OK);
                instalado = TRUE;
                break;
            } else {
                // tentar próxima alternativa
            }
        }
    }

    if (!instalado) {
        MessageBoxA(NULL,
                   "Não foi possível instalar automaticamente.\nPor favor execute manualmente (como administrador):\nsetupc.exe install PortName=COM5 PortName=COM6",
                   "Ação manual necessária", MB_ICONERROR);
        return 1;
    }

    return 0;
}
