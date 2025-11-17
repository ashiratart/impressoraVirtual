# impressoraVirtual
interprete zpl2 para elgin pplb

Estrutura

drivers_impressora_virtual/
├── src/
│   ├── printer_controller.c     // Ponto 4 - Principal
│   ├── virtual_printer.c        // Ponto 1 - Máquina Virtual
│   ├── job_monitor.c           // Ponto 2 - Gerenciamento
│   ├── elgin_translator.c      // Ponto 3 - Tradutor/Comunicação
│   └── config_manager.c        // Gerenciador de Configurações
├── include/
├── build/
└── README.md