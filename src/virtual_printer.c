#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h>
#include <winspool.h>

// Estrutura para gerenciar a impressora virtual
typedef struct {
    char virtual_printer_name[256];
    char real_printer_name[256];
    HANDLE hRealPrinter;
} PrinterRedirect;

// Função para instalar a impressora virtual
BOOL install_virtual_printer() {
    PRINTER_INFO_2 printerInfo;
    memset(&printerInfo, 0, sizeof(printerInfo));
    
    printerInfo.pPrinterName = "Zebra Virtual (Redirect to Elgin)";
    printerInfo.pPortName = "VIRTUAL_USB:";
    printerInfo.pDriverName = "ZDesigner GK420t";
    printerInfo.pPrintProcessor = "WinPrint";
    printerInfo.pDatatype = "RAW";
    printerInfo.Attributes = PRINTER_ATTRIBUTE_DIRECT;
    
    HANDLE hPrinter;
    if (!AddPrinter(NULL, 2, (LPBYTE)&printerInfo)) {
        printf("Erro instalando impressora virtual: %d\n", GetLastError());
        return FALSE;
    }
    
    printf("Impressora virtual instalada: %s\n", printerInfo.pPrinterName);
    ClosePrinter(hPrinter);
    return TRUE;
}

// Função para monitorar jobs de impressão
DWORD WINAPI monitor_print_jobs(LPVOID lpParam) {
    PrinterRedirect* redirect = (PrinterRedirect*)lpParam;
    
    while (1) {
        Sleep(1000); // Verifica a cada segundo
        
        // Enumera printers para encontrar nossa virtual
        DWORD needed, returned;
        EnumPrinters(PRINTER_ENUM_LOCAL, NULL, 2, NULL, 0, &needed, &returned);
        
        if (needed > 0) {
            BYTE* buffer = malloc(needed);
            if (EnumPrinters(PRINTER_ENUM_LOCAL, NULL, 2, buffer, needed, &needed, &returned)) {
                PRINTER_INFO_2* printers = (PRINTER_INFO_2*)buffer;
                
                for (DWORD i = 0; i < returned; i++) {
                    if (strstr(printers[i].pPrinterName, "Zebra Virtual")) {
                        // Monitora jobs desta impressora
                        monitor_printer_jobs(printers[i].pPrinterName, redirect);
                    }
                }
            }
            free(buffer);
        }
    }
    return 0;
}