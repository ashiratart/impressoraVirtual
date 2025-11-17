// Cria um monitor de porta virtual
BOOL create_virtual_port_monitor() {
    // Usa um named pipe para simular a porta USB virtual
    HANDLE hPipe = CreateNamedPipe(
        "\\\\.\\pipe\\VirtualUSBPrinter",
        PIPE_ACCESS_DUPLEX,
        PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT,
        1,           // Máximo de instâncias
        8192,        // Buffer de saída
        8192,        // Buffer de entrada
        0,           // Timeout padrão
        NULL         // Segurança
    );
    
    if (hPipe == INVALID_HANDLE_VALUE) {
        printf("Erro criando pipe virtual: %d\n", GetLastError());
        return FALSE;
    }
    
    printf("Porta USB virtual criada: \\\\.\\pipe\\VirtualUSBPrinter\n");
    return TRUE;
}

// Função principal de conversão ZPL2 para Elgin
char* convert_zpl_to_elgin(const char* zpl_data) {
    char* result = malloc(strlen(zpl_data) * 2); // Buffer generoso
    strcpy(result, "");
    
    // Comandos básicos de mapeamento
    if (strstr(zpl_data, "^XA")) {
        strcat(result, "\x1B\x40"); // Inicializar impressora Elgin
    }
    
    // Conversão de texto básico
    // Você precisará expandir esta lógica conforme seus templates
    char* pos = (char*)zpl_data;
    while (*pos) {
        if (strncmp(pos, "^FD", 3) == 0) {
            // Field Data - extrai o texto
            pos += 3;
            char* end = strstr(pos, "^FS");
            if (end) {
                int len = end - pos;
                char text[256];
                strncpy(text, pos, len);
                text[len] = '\0';
                
                // Adiciona comando de texto Elgin
                strcat(result, "\x1B\x61\x00"); // Alinhamento esquerdo
                strcat(result, text);
                pos = end + 3;
                continue;
            }
        }
        pos++;
    }
    
    // Comando de corte (se existir ^XZ)
    if (strstr(zpl_data, "^XZ")) {
        strcat(result, "\n\n\n\n\n\x1B\x64\x02"); // Avança e corta
    }
    
    return result;
}