// ===================================================================================
// Installer completo com criação de impressora virtual Zebra (versão simples / RAW)
// ===================================================================================
#include <windows.h>
#include <wininet.h>
#include <winspool.h>
#include <stdio.h>

#pragma comment(lib, "Wininet.lib")
#pragma comment(lib, "winspool.lib")
#pragma comment(lib, "winspool.drv")

// =============================================
// 1. Solicitar modo administrador (UAC)
// =============================================
void relaunchAsAdmin() {
    BOOL isAdmin = FALSE;
    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
    PSID AdministratorsGroup;

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

        SHELLEXECUTEINFOA sei = { sizeof(sei) };
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
// 5. Executar setup do com0com
// (mantive teu código original com pequenas correções)
// =============================================
void executarSetupCom0com() {
    char comando[1024];
    char mensagem[1024];
    char pastaExe[MAX_PATH];
    
    getExecutablePath(pastaExe, MAX_PATH);
    
    MessageBoxA(NULL, "Instalando com0com...", "Aguarde", MB_OK);
    
    // Tenta encontrar e executar o setupc.exe em várias localizações possíveis
    const char *caminhos[] = {
        "com0com-3.0.0.0-i386-and-x64-signed\\setupc.exe",
        "com0com\\setupc.exe", 
        "setupc.exe",
        NULL
    };
    
    BOOL encontrado = FALSE;
    BOOL instalado = FALSE;
    
    for (int i = 0; caminhos[i] != NULL; i++) {
        char caminhoCompleto[MAX_PATH];
        sprintf_s(caminhoCompleto, sizeof(caminhoCompleto), "%s%s", pastaExe, caminhos[i]);
        
        if (arquivoExiste(caminhoCompleto)) {
            // tenta instalar pares COM (COM5 <-> COM6)
            sprintf_s(comando, sizeof(comando), "cmd.exe /c \"%s\" install PortName=COM5 PortName=COM6", caminhoCompleto);
            sprintf_s(mensagem, sizeof(mensagem), "Executando comando:\n%s", comando);
            MessageBoxA(NULL, mensagem, "Instalação com0com", MB_OK);
            
            int resultado = system(comando);
            
            if (resultado == 0) {
                sprintf_s(mensagem, sizeof(mensagem),
                    "com0com instalado com sucesso!\n\n"
                    "Portas virtuais criadas:\n"
                    "COM5 <-> COM6");
                MessageBoxA(NULL, mensagem, "Sucesso", MB_OK);
                instalado = TRUE;
            } else {
                sprintf_s(mensagem, sizeof(mensagem),
                    "Erro na instalação do com0com.\n"
                    "Código: %d\n\n"
                    "Tentando com permissões elevadas...", resultado);
                MessageBoxA(NULL, mensagem, "Aviso", MB_ICONWARNING);
                
                // Tenta com ShellExecute para elevação (parâmetros para criar COM5<->COM6)
                SHELLEXECUTEINFOA sei = { sizeof(sei) };
                sei.lpVerb = "runas";
                sei.lpFile = caminhoCompleto;
                sei.lpParameters = "install PortName=COM5 PortName=COM6";
                sei.nShow = SW_NORMAL;
                
                if (ShellExecuteExA(&sei)) {
                    MessageBoxA(NULL, "Instalação iniciada com permissões elevadas!", "Sucesso", MB_OK);
                    instalado = TRUE;
                }
            }
            
            encontrado = TRUE;
            break;
        }
    }
    
    if (!encontrado) {
        sprintf_s(mensagem, sizeof(mensagem),
            "setupc.exe não encontrado!\n\n"
            "Por favor verifique:\n"
            "1. O arquivo com0com.zip foi extraído?\n"
            "2. A pasta 'com0com' foi criada?\n"
            "3. Execute setupc.exe manualmente como administrador:\n"
            "   setupc.exe install PortName=COM5 PortName=COM6");
        MessageBoxA(NULL, mensagem, "Arquivo Não Encontrado", MB_ICONERROR);
    } else if (!instalado) {
        MessageBoxA(NULL, 
            "Instalação não concluída.\n"
            "Execute setupc.exe manualmente como administrador.", 
            "Instalação Manual Necessária", MB_ICONINFORMATION);
    }
}

// =============================================
// 6. Instalar com0com (versão final)
// =============================================
void instalarCom0Com() {
    char mensagem[512];
    char pastaExe[MAX_PATH];
    char caminhoZip[MAX_PATH];
    
    getExecutablePath(pastaExe, MAX_PATH);
    sprintf_s(caminhoZip, sizeof(caminhoZip), "%scom0com.zip", pastaExe);
    
    // Verifica se o ZIP já existe na pasta
    if (arquivoExiste(caminhoZip)) {
        DWORD tamanho = 0;
        HANDLE hArq = CreateFileA(caminhoZip, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hArq != INVALID_HANDLE_VALUE) {
            tamanho = GetFileSize(hArq, NULL);
            CloseHandle(hArq);
        }
        
        sprintf_s(mensagem, sizeof(mensagem),
            "com0com.zip encontrado!\n"
            "Local: %s\n"
            "Tamanho: %lu bytes\n\n"
            "Clique OK para usar este arquivo.",
            caminhoZip, (unsigned long)tamanho);
            
        MessageBoxA(NULL, mensagem, "Arquivo Local Encontrado", MB_OK);
        
        // Extrai o arquivo ZIP
        char pastaDestino[MAX_PATH];
        sprintf_s(pastaDestino, sizeof(pastaDestino), "%scom0com", pastaExe);
        
        MessageBoxA(NULL, "Extraindo com0com...", "Aguarde", MB_OK);
        if (extrairZip(caminhoZip, pastaDestino)) {
            executarSetupCom0com();
        }
        
    } else {
        sprintf_s(mensagem, sizeof(mensagem),
            "com0com.zip não encontrado em:\n%s\n\n"
            "Por favor coloque o arquivo com0com.zip\n"
            "na mesma pasta do executável e execute novamente.",
            caminhoZip);
        MessageBoxA(NULL, mensagem, "Arquivo Não Encontrado", MB_ICONERROR);
    }
}

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
    sprintf(buffer,
        "%s\n\nErro: %lu\nDescrição: %s",
        contexto, err, msg ? msg : "(sem descrição)"
    );

    MessageBoxA(NULL, buffer, "LOG DE ERRO", MB_ICONERROR);

    if (msg)
        LocalFree(msg);
}



// =============================================
// 7. Criar impressora virtual (versão simples / RAW)
// =============================================
void instalarImpressora()
{
    MessageBoxA(NULL, 
        "Iniciando instalação da impressora virtual...\n"
        "Logs detalhados serão exibidos.",
        "LOG", MB_OK);

    //----------------------------------------------
    // 1. Criar porta COM5
    //----------------------------------------------
    MessageBoxA(NULL, "Criando porta COM5 usando AddPortA...", "LOG", MB_OK);

    BOOL portResult = AddPortA(NULL, NULL, "COM5");

    if (!portResult)
    {
        mostrarErro("Falha em AddPortA(NULL, NULL, \"COM5\")");
    }
    else
    {
        MessageBoxA(NULL, "Porta COM5 criada com sucesso!", "LOG", MB_OK);
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
    pi.pPortName = "COM5";
    pi.pDriverName = "Generic / Text Only";
    pi.pPrintProcessor = "WinPrint";
    pi.pDatatype = "RAW";

    HANDLE h = AddPrinterA(NULL, 2, (BYTE *)&pi);

    if (!h)
    {
        mostrarErro("Falha em AddPrinterA");
        return;
    }

    MessageBoxA(NULL,
        "Impressora instalada com sucesso!\n\n"
        "Nome: Zebra Virtual RAW\n"
        "Porta: COM5\n"
        "Driver: Generic / Text Only\n\n"
        "Pronta para receber ZPL2.",
        "Sucesso", MB_OK);
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
        "1. Buscar com0com.zip na pasta atual\n"
        "2. Extrair e instalar com0com (COM5 <-> COM6)\n"
        "3. Criar porta COM5 no spooler\n"
        "4. Instalar impressora virtual Zebra (Generic / Text Only -> COM5)\n\n"
        "Clique OK para continuar...",
        pastaExe);
    
    MessageBoxA(NULL, mensagem, "Instalador", MB_OK);
    
    instalarCom0Com();
    instalarImpressora();
    
    MessageBoxA(NULL, "Processo finalizado!", "Sucesso", MB_OK);
    return 0;
}
