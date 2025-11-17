int main() {
    printf("=== Driver Virtual ZPL2 -> Elgin USB ===\n");
    
    // Instala impressora virtual
    if (!install_virtual_printer()) {
        printf("Falha na instalação da impressora virtual\n");
        return 1;
    }
    
    // Cria porta virtual
    if (!create_virtual_port_monitor()) {
        printf("Falha na criação da porta virtual\n");
        return 1;
    }
    
    printf("\nDriver instalado com sucesso!\n");
    printf("Configure seus programas para usar: 'Zebra Virtual (Redirect to Elgin)'\n");
    printf("Pressione Enter para sair...\n");
    getchar();
    
    return 0;
}