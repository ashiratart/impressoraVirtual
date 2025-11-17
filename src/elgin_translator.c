// Encontra e conecta na impressora Elgin real
HANDLE find_and_connect_elgin() {
    HANDLE hPrinter;
    char printerName[256];
    
    // Tenta encontrar impressora Elgin
    if (find_elgin_printer(printerName, sizeof(printerName))) {
        if (OpenPrinter(printerName, &hPrinter, NULL)) {
            printf("Conectado à impressora Elgin: %s\n", printerName);
            return hPrinter;
        }
    }
    
    // Fallback: tenta via porta USB direta
    hPrinter = open_usb_direct();
    if (hPrinter != INVALID_HANDLE_VALUE) {
        return hPrinter;
    }
    
    return INVALID_HANDLE_VALUE;
}

// Abre via USB direto (quando a impressora é reconhecida como dispositivo USB)
HANDLE open_usb_direct() {
    // Tenta portas USB comuns
    const char* usb_ports[] = {
        "USB001", "USB002", "USB003", "USB004"
    };
    
    for (int i = 0; i < 4; i++) {
        char port_name[32];
        sprintf(port_name, "%s:", usb_ports[i]);
        
        HANDLE hPort = CreateFile(port_name, 
            GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
        
        if (hPort != INVALID_HANDLE_VALUE) {
            printf("Conectado via USB: %s\n", port_name);
            return hPort;
        }
    }
    
    return INVALID_HANDLE_VALUE;
}

// Envia dados para impressora Elgin
BOOL send_to_elgin_printer(const char* data, size_t length) {
    HANDLE hElgin = find_and_connect_elgin();
    
    if (hElgin == INVALID_HANDLE_VALUE) {
        printf("Erro: Não foi possível conectar à impressora Elgin\n");
        return FALSE;
    }
    
    DOC_INFO_1 docInfo;
    docInfo.pDocName = "Etiqueta Convertida";
    docInfo.pOutputFile = NULL;
    docInfo.pDatatype = "RAW";
    
    DWORD jobId = StartDocPrinter(hElgin, 1, (LPBYTE)&docInfo);
    if (jobId > 0) {
        StartPagePrinter(hElgin);
        
        DWORD bytesWritten;
        BOOL success = WritePrinter(hElgin, (LPVOID)data, length, &bytesWritten);
        
        EndPagePrinter(hElgin);
        EndDocPrinter(hElgin);
        
        ClosePrinter(hElgin);
        
        if (success) {
            printf("Dados enviados para Elgin: %d bytes\n", bytesWritten);
            return TRUE;
        }
    }
    
    ClosePrinter(hElgin);
    return FALSE;
}