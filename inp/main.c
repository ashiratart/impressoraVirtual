// installer_com0com_zebra.c
// Compilar com Visual Studio (ex: cl /EHsc installer_com0com_zebra.c)
// ===================================================================================
// Installer completo com criação de impressora virtual Zebra (com detecção de com0com)
// ===================================================================================

#include <windows.h>
#include <wininet.h>
#include <winspool.h>
#include <stdio.h>
#include <string.h>

#pragma comment(lib, "Wininet.lib")
#pragma comment(lib, "winspool.lib")
#pragma comment(lib, "winspool.drv")

// -------------------------------------------------
// Variável global: porta selecionada para a impressora
// -------------------------------------------------
char g_portaSelecionada[64] = "COM5"; // fallback padrão

// =============================================
// 1. Solicitar modo administrador (UAC)
// =============================================
void relaunchAsAdmin() {
    BOOL isAdmin = FALSE;
    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
    PSID AdministratorsGroup = NULL;

    if (AllocateAndInitializeSid(
        &NtAuthority, 2,
        SECURITY_BUILTIN_DOMAIN_RID,
        DOMAIN_ALIAS_RID_ADMINS,
        0, 0, 0, 0, 0, 0,
        &AdministratorsGroup)) {
        
        CheckTokenMembership(NULL, AdministratorsGroup, &isAdmin);
        FreeSid(AdministratorsGroup);
    }

    if (!isAdmin) {
        char path[MAX_PATH];
        GetModuleFileNameA(NULL, path, MAX_PATH);

        SHELLEXECUTEINFOA sei = { 0 };
        sei.cbSize = sizeof(sei);
        sei.lpVerb = "runas";
        sei.lpFile = path;
        sei.hwnd = NULL;
        sei.nShow = SW_NORMAL;

        if (!ShellExecuteExA(&sei)) {
            MessageBoxA(NULL, "Permissão de administrador é necessária.", "Erro", MB_ICONERROR);
            exit(1);
        }
        exit(0);
    }
}

// =============================================
// 2. Obter caminho da pasta do executável
// =============================================
void getExecutablePath(char* buffer, size_t size) {
    GetModuleFileNameA(NULL, buffer, (DWORD)size);
    
    // Remove o nome do executável, mantendo apenas o caminho
    char* lastBackslash = strrchr(buffer, '\\');
    if (lastBackslash) {
        *(lastBackslash + 1) = '\0';  // Mantém a barra no final
    }
}

// =============================================
// 3. Verificar se arquivo existe
// =============================================
BOOL arquivoExiste(const char *caminho) {
    DWORD attrs = GetFileAttributesA(caminho);
    return (attrs != INVALID_FILE_ATTRIBUTES && !(attrs & FILE_ATTRIBUTE_DIRECTORY));
}

// =============================================
// 4. Extrair arquivo ZIP
// =============================================
BOOL extrairZip(const char *arquivoZip, const char *pastaDestino) {
    char comando[1024];
    char mensagem[512];
    
    // Cria pasta de destino se não existir
    CreateDirectoryA(pastaDestino, NULL);
    
    // Tenta com PowerShell primeiro (mais confiável)
    sprintf_s(comando, sizeof(comando),
        "powershell -Command \"Expand-Archive -Path '%s' -DestinationPath '%s' -Force\"",
        arquivoZip, pastaDestino);
    
    sprintf_s(mensagem, sizeof(mensagem), "Executando: %s", comando);
    MessageBoxA(NULL, mensagem, "DEBUG Extração", MB_OK);
    
    int resultado = system(comando);
    
    if (resultado == 0) {
        sprintf_s(mensagem, sizeof(mensagem), "Arquivo extraído com sucesso!\nOrigem: %s\nDestino: %s", arquivoZip, pastaDestino);
        MessageBoxA(NULL, mensagem, "Sucesso", MB_OK);
        return TRUE;
    } else {
        // Fallback: tenta com tar (Windows 10+)
        sprintf_s(comando, sizeof(comando), "tar -xf \"%s\" -C \"%s\"", arquivoZip, pastaDestino);
        resultado = system(comando);
        
        if (resultado == 0) {
            sprintf_s(mensagem, sizeof(mensagem), "Arquivo extraído com tar!\nOrigem: %s\nDestino: %s", arquivoZip, pastaDestino);
            MessageBoxA(NULL, mensagem, "Sucesso", MB_OK);
            return TRUE;
        }
    }
    
    sprintf_s(mensagem, sizeof(mensagem),
        "Falha na extração automática.\n\n"
        "Por favor extraia manualmente:\n"
        "1. Clique direito em '%s'\n"
        "2. 'Extrair Tudo'\n"
        "3. Pasta destino: '%s'",
        arquivoZip, pastaDestino);
    MessageBoxA(NULL, mensagem, "Extração Manual", MB_ICONINFORMATION);
    
    return FALSE;
}

// =============================================
// mostrarErro - utilitário com GetLastError -> descrição
// =============================================
void mostrarErro(const char* contexto)
{
    DWORD err = GetLastError();
    char* msg = NULL;

    FormatMessageA(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL,
        err,
        0,
        (LPSTR)&msg,
        0,
        NULL
    );

    char buffer[1024];
    sprintf_s(buffer, sizeof(buffer),
        "%s\n\nErro: %lu\nDescrição: %s",
        contexto, err, msg ? msg : "(sem descrição)"
    );

    MessageBoxA(NULL, buffer, "LOG DE ERRO", MB_ICONERROR);

    if (msg)
        LocalFree(msg);
}

// =============================================
// VERIFICAR SE com0com JÁ ESTÁ INSTALADO
// (procura por entradas CNCAx/CNCBx em SERIALCOMM)
// =============================================
BOOL com0comInstalado() {
    HKEY hKey;
    if (RegOpenKeyExA(
            HKEY_LOCAL_MACHINE,
            "HARDWARE\\DEVICEMAP\\SERIALCOMM",
            0,
            KEY_READ,
            &hKey
        ) != ERROR_SUCCESS) 
    {
        return FALSE;
    }

    char nomeValor[256];
    char dado[256];
    DWORD tamNome = sizeof(nomeValor);
    DWORD tamDado = sizeof(dado);
    DWORD tipo;

    BOOL encontrado = FALSE;

    for (DWORD i = 0;; i++) {
        tamNome = sizeof(nomeValor);
        tamDado = sizeof(dado);

        LONG res = RegEnumValueA(
            hKey,
            i,
            nomeValor,
            &tamNome,
            NULL,
            &tipo,
            (LPBYTE)dado,
            &tamDado
        );

        if (res != ERROR_SUCCESS)
            break;

        // >>> com0com costuma expor CNCAx / CNCBx no nome do valor <<<
        if (strstr(nomeValor, "CNCA") || strstr(nomeValor, "CNCB")) {
            encontrado = TRUE;
            break;
        }
    }

    RegCloseKey(hKey);
    return encontrado;
}

// =============================================
// OBTER PAR DE PORTAS JÁ INSTALADAS PELO com0com
// (preenche portaA e portaB com strings como "COM3", "COM4")
// retorna TRUE se encontrou ambos
// =============================================
BOOL obterParExistente(char* portaA, size_t tamA, char* portaB, size_t tamB) {
    // Inicializa como vazias
    if (portaA && tamA > 0) portaA[0] = '\0';
    if (portaB && tamB > 0) portaB[0] = '\0';

    HKEY hKey;
    if (RegOpenKeyExA(
            HKEY_LOCAL_MACHINE,
            "HARDWARE\\DEVICEMAP\\SERIALCOMM",
            0,
            KEY_READ,
            &hKey
        ) != ERROR_SUCCESS) 
    {
        return FALSE;
    }

    char nomeValor[256];
    char dado[256];
    DWORD tamNome, tamDado, tipo;
    BOOL achouA = FALSE, achouB = FALSE;

    for (DWORD i = 0;; i++) {
        tamNome = sizeof(nomeValor);
        tamDado = sizeof(dado);

        LONG res = RegEnumValueA(
            hKey,
            i,
            nomeValor,
            &tamNome,
            NULL,
            &tipo,
            (LPBYTE)dado,
            &tamDado
        );

        if (res != ERROR_SUCCESS)
            break;

        if (!achouA && strstr(nomeValor, "CNCA")) {
            // dado contém algo como "COM3"
            strncpy_s(portaA, tamA, dado, _TRUNCATE);
            achouA = TRUE;
        }
        if (!achouB && strstr(nomeValor, "CNCB")) {
            strncpy_s(portaB, tamB, dado, _TRUNCATE);
            achouB = TRUE;
        }

        if (achouA && achouB)
            break;
    }

    RegCloseKey(hKey);
    return (achouA && achouB);
}

// =============================================
// Pergunta ao usuário qual porta do par deseja usar
// Preenche 'saida' com a porta selecionada (ex: "COM3")
// Retorna TRUE se escolha foi feita
// =============================================
BOOL selecionarPortaUsuario(const char* portaA, const char* portaB, char* saida, size_t tamSaida) {
    if (!portaA || !portaA[0]) {
        // se só portaB existe, usa portaB
        if (portaB && portaB[0]) {
            strncpy_s(saida, tamSaida, portaB, _TRUNCATE);
            return TRUE;
        }
        return FALSE;
    }

    if (!portaB || !portaB[0]) {
        // se só portaA existe, usa portaA
        strncpy_s(saida, tamSaida, portaA, _TRUNCATE);
        return TRUE;
    }

    // Monta mensagem; usamos MB_YESNO:
    // Yes -> usar a primeira porta (portaA)
    // No  -> usar a segunda porta (portaB)
    char mensagem[512];
    sprintf_s(mensagem, sizeof(mensagem),
        "Par detectado: %s  <->  %s\n\n"
        "Qual porta deseja usar para a impressora?\n\n"
        "Clique Sim para usar %s\n"
        "Clique Não para usar %s",
        portaA, portaB, portaA, portaB);

    int resp = MessageBoxA(NULL, mensagem, "Escolher Porta", MB_ICONQUESTION | MB_YESNO);

    if (resp == IDYES) {
        strncpy_s(saida, tamSaida, portaA, _TRUNCATE);
        return TRUE;
    } else {
        strncpy_s(saida, tamSaida, portaB, _TRUNCATE);
        return TRUE;
    }
}

// =============================================
// Executar setupc.exe com nomes de portas passadas
// (caminho completo para setupc.exe e PortName args)
// =============================================
BOOL executarSetupcComPortas(const char* caminhoCompletoSetupc, const char* porta1, const char* porta2) {
    if (!arquivoExiste(caminhoCompletoSetupc)) return FALSE;

    char comando[1024];
    // command example: cmd.exe /c "C:\path\setupc.exe" install PortName=COM5 PortName=COM6
    sprintf_s(comando, sizeof(comando), "cmd.exe /c \"%s\" install PortName=%s PortName=%s", caminhoCompletoSetupc, porta1, porta2);

    int resultado = system(comando);
    if (resultado == 0) {
        return TRUE;
    } else {
        // tenta com elevação via ShellExecuteEx
        SHELLEXECUTEINFOA sei = { 0 };
        sei.cbSize = sizeof(sei);
        sei.lpVerb = "runas";
        sei.lpFile = caminhoCompletoSetupc;
        // monta parâmetros:
        char params[256];
        sprintf_s(params, sizeof(params), "install PortName=%s PortName=%s", porta1, porta2);
        sei.lpParameters = params;
        sei.nShow = SW_NORMAL;

        if (ShellExecuteExA(&sei)) {
            return TRUE;
        }
    }

    return FALSE;
}

// =============================================
// Função principal para instalar com0com (integra lógica de detecção)
// =============================================
void instalarCom0Com() {
    char pastaExe[MAX_PATH];
    getExecutablePath(pastaExe, MAX_PATH);

    // Se já está instalado -> detectar par e perguntar ao usuário qual porta usar
    if (com0comInstalado()) {
        char a[64] = {0}, b[64] = {0};
        if (obterParExistente(a, sizeof(a), b, sizeof(b))) {
            // pergunta ao usuário qual porta usar
            char escolha[64] = {0};
            if (selecionarPortaUsuario(a, b, escolha, sizeof(escolha))) {
                strncpy_s(g_portaSelecionada, sizeof(g_portaSelecionada), escolha, _TRUNCATE);
                char msg[256];
                sprintf_s(msg, sizeof(msg),
                    "com0com detectado.\nPortas: %s <-> %s\nUsando: %s",
                    a, b, g_portaSelecionada);
                MessageBoxA(NULL, msg, "com0com detectado", MB_OK);
                return; // não reinstala
            } else {
                // se não escolheu, usa a primeira como fallback
                strncpy_s(g_portaSelecionada, sizeof(g_portaSelecionada), a, _TRUNCATE);
                return;
            }
        } else {
            MessageBoxA(NULL,
                "com0com detectado, mas não foi possível localizar o par CNCA/CNCB.\nSerá necessária intervenção manual.",
                "Aviso", MB_ICONWARNING);
            return;
        }
    }

    // Se não estiver instalado -> procurar com0com.zip e instalar COM5<->COM6
    char caminhoZip[MAX_PATH];
    sprintf_s(caminhoZip, sizeof(caminhoZip), "%scom0com.zip", pastaExe);

    if (!arquivoExiste(caminhoZip)) {
        char msg[512];
        sprintf_s(msg, sizeof(msg),
            "com0com.zip não encontrado em:\n%s\n\n"
            "Por favor coloque o arquivo com0com.zip\n"
            "na mesma pasta do executável e execute novamente.",
            caminhoZip);
        MessageBoxA(NULL, msg, "Arquivo Não Encontrado", MB_ICONERROR);
        return;
    }

    // extrair
    char pastaDestino[MAX_PATH];
    sprintf_s(pastaDestino, sizeof(pastaDestino), "%scom0com", pastaExe);
    MessageBoxA(NULL, "Extraindo com0com...", "Aguarde", MB_OK);
    if (!extrairZip(caminhoZip, pastaDestino)) {
        MessageBoxA(NULL, "Falha na extração do com0com.zip", "Erro", MB_ICONERROR);
        return;
    }

    // localizar setupc.exe em alguns caminhos
    const char *caminhosPossiveis[] = {
        "com0com-3.0.0.0-i386-and-x64-signed\\setupc.exe",
        "com0com\\setupc.exe",
        "setupc.exe",
        NULL
    };

    BOOL instalado = FALSE;
    for (int i = 0; caminhosPossiveis[i] != NULL; i++) {
        char caminhoCompleto[MAX_PATH];
        sprintf_s(caminhoCompleto, sizeof(caminhoCompleto), "%s%s", pastaDestino, caminhosPossiveis[i]);

        if (arquivoExiste(caminhoCompleto)) {
            // tenta instalar COM5 <-> COM6
            if (executarSetupcComPortas(caminhoCompleto, "COM5", "COM6")) {
                installed:
                strncpy_s(g_portaSelecionada, sizeof(g_portaSelecionada), "COM5", _TRUNCATE);
                MessageBoxA(NULL, "com0com instalado com sucesso (COM5 <-> COM6).", "Sucesso", MB_OK);
                instalado = TRUE;
                break;
            } else {
                // tentar executar com elevação foi feito no executarSetupcComPortas
                // ainda assim, se falhar, informar
                MessageBoxA(NULL, "Falha ao instalar com0com com setupc.exe", "Erro", MB_ICONERROR);
            }
        } else {
            // também tentar caso setupc.exe esteja diretamente na pasta do executável
            char caminhoAlternativo[MAX_PATH];
            sprintf_s(caminhoAlternativo, sizeof(caminhoAlternativo), "%s%s", pastaExe, caminhosPossiveis[i]);
            if (arquivoExiste(caminhoAlternativo)) {
                if (executarSetupcComPortas(caminhoAlternativo, "COM5", "COM6")) {
                    goto installed;
                }
            }
        }
    }

    if (!instalado) {
        MessageBoxA(NULL,
            "Não foi possível instalar o com0com automaticamente.\nExecute setupc.exe manualmente como administrador:\nsetupc.exe install PortName=COM5 PortName=COM6",
            "Instalação Manual Necessária", MB_ICONERROR);
    }
}

// =============================================
// 7. Criar impressora virtual (usando g_portaSelecionada)
// =============================================
void instalarImpressora()
{
    MessageBoxA(NULL, 
        "Iniciando instalação da impressora virtual...\n"
        "Logs detalhados serão exibidos.",
        "LOG", MB_OK);

    char msg[256];
    sprintf_s(msg, sizeof(msg), "Usando porta: %s", g_portaSelecionada);
    MessageBoxA(NULL, msg, "LOG", MB_OK);

    //----------------------------------------------
    // 1. Criar porta (se necessário)
    //----------------------------------------------
    char aviso[256];
    sprintf_s(aviso, sizeof(aviso), "Criando porta %s usando AddPortA...", g_portaSelecionada);
    MessageBoxA(NULL, aviso, "LOG", MB_OK);

    BOOL portResult = AddPortA(NULL, NULL, g_portaSelecionada);

    if (!portResult)
    {
        mostrarErro("Falha em AddPortA(NULL, NULL, porta)");
    }
    else
    {
        MessageBoxA(NULL, "Porta criada (ou já existia) com sucesso!", "LOG", MB_OK);
    }


    //----------------------------------------------
    // 2. Instalar driver 'Generic / Text Only'
    //----------------------------------------------
    MessageBoxA(NULL, 
        "Instalando driver 'Generic / Text Only'...\n"
        "Chamando AddPrinterDriverA...",
        "LOG", MB_OK);

    DRIVER_INFO_3A di;
    memset(&di, 0, sizeof(di));

    di.cVersion = 3;
    di.pName = "Generic / Text Only";
    di.pEnvironment = "Windows x64";
    di.pDriverPath = "C:\\Windows\\System32\\ntprint.dll";
    di.pConfigFile = "UNIDRV.DLL";
    di.pDataFile = "TTFSUB.GPD";

    if (!AddPrinterDriverA(NULL, 3, (BYTE*)&di))
    {
        // não interrompe totalmente — exibe erro
        mostrarErro("Falha em AddPrinterDriverA");
    }
    else
    {
        MessageBoxA(NULL, 
            "Driver instalado (ou já estava instalado).", 
            "LOG", MB_OK);
    }


    //----------------------------------------------
    // 3. Criar impressora 'Zebra Virtual RAW'
    //----------------------------------------------
    MessageBoxA(NULL, 
        "Criando impressora 'Zebra Virtual RAW'...\n"
        "Chamando AddPrinterA...",
        "LOG", MB_OK);

    PRINTER_INFO_2A pi;
    memset(&pi, 0, sizeof(pi));

    pi.pPrinterName = "Zebra Virtual RAW";
    pi.pPortName = g_portaSelecionada;
    pi.pDriverName = "Generic / Text Only";
    pi.pPrintProcessor = "WinPrint";
    pi.pDatatype = "RAW";

    HANDLE h = AddPrinterA(NULL, 2, (BYTE *)&pi);

    if (!h)
    {
        mostrarErro("Falha em AddPrinterA");
        return;
    }

    char sucessoMsg[512];
    sprintf_s(sucessoMsg, sizeof(sucessoMsg),
        "Impressora instalada com sucesso!\n\n"
        "Nome: Zebra Virtual RAW\n"
        "Porta: %s\n"
        "Driver: Generic / Text Only\n\n"
        "Pronta para receber ZPL2.",
        g_portaSelecionada);

    MessageBoxA(NULL, sucessoMsg, "Sucesso", MB_OK);
}



// =============================================
// MAIN
// =============================================
int main() {
    relaunchAsAdmin();
    
    char pastaExe[MAX_PATH];
    getExecutablePath(pastaExe, MAX_PATH);
    
    char mensagem[1024];
    sprintf_s(mensagem, sizeof(mensagem),
        "Instalador de Impressora Virtual Zebra\n\n"
        "Pasta do executável:\n%s\n\n"
        "Este programa irá:\n"
        "1. Verificar se com0com já está instalado\n"
        "2. Se instalado: detectar par de portas e perguntar qual usar\n"
        "3. Se não instalado: buscar com0com.zip, extrair e instalar COM5 <-> COM6\n"
        "4. Criar porta selecionada no spooler\n"
        "5. Instalar impressora virtual Zebra (Generic / Text Only -> porta selecionada)\n\n"
        "Clique OK para continuar...",
        pastaExe);
    
    MessageBoxA(NULL, mensagem, "Instalador", MB_OK);
    
    instalarCom0Com();      // atualiza g_portaSelecionada conforme lógica
    instalarImpressora();   // usa g_portaSelecionada
    
    MessageBoxA(NULL, "Processo finalizado!", "Sucesso", MB_OK);
    return 0;
}
