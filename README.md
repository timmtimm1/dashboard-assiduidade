flowchart TD
  A([Início]) --> B{Pasta "screenshots_ponto" existe e tem imagens?}
  B -- Não --> Z1[Log: Erro "pasta vazia"] --> Z2([Fim])
  B -- Sim --> C[Listar *.png/*.jpg]
  C --> D[Para cada imagem]
  D --> E[Extrair nome do colaborador do nome do arquivo]
  E --> F[OCR: partition_image(..., ocr_only)]
  F --> G{Regex HH:MM(:SS)? encontrou hora?}
  G -- Sim --> H[Registro: Bateu Ponto=Sim; Hora=hh:mm:ss]
  G -- Não --> I[Registro: Bateu Ponto=Não; Hora=N/A]
  F -- Exceção --> J[Registro: Bateu Ponto=Erro; Hora=N/A]
  H & I & J --> K[Construir df_hoje]
  K --> L[Normalizar nomes; Data = hoje (dd/mm/aaaa)]
  L --> M[Merge com df_func para trazer Turno]
  M --> N[Selecionar colunas: Data, Colaborador, Turno, Bateu Ponto, Hora]
  N --> O[Carregar df_existente do Excel (_ler_existente_excel)]

  subgraph Leitura_e_Normalização_Excel
    O --> O1{Arquivo existe?}
    O1 -- Não --> O2[Cria DF vazio com COLS_PADRAO]
    O1 -- Sim --> O3[Ler aba "Registros" (dtype=str)]
    O3 --> O4[Garantir colunas padrão]
    O4 --> O5[Normalizar: Colaborador(title), Data(date), Hora/Bateu Ponto(texto)]
    O5 --> O6[Se Data==hoje e Hora vazia/N/A → Hora="00:00:00"]
    O6 --> O7[Data_txt = dd/mm/aaaa]
  end

  O7 --> P[_consolidar_excel(df_exist, df_hoje)]
  subgraph Consolidação
    P --> P1[Concatenar df_exist + df_hoje]
    P1 --> P2[Prioridade = (Bateu==SIM)*10 + (temHora)*1]
    P2 --> P3[Ordenar: Data↑, Colaborador↑, Prioridade↓, Hora↓]
    P3 --> P4[Herdar _PowerAppsId_ por (Data, Colaborador)]
    P4 --> P5[Remover duplicados por (Data, Colaborador) (keep first)]
    P5 --> P6[Ordenar final: Data↓, Colaborador↑]
  end

  P6 --> Q[Formatar: Hora→"hh:mm:ss"; Data_txt; ordem COLS_PADRAO]
  Q --> R[_excel_writer_retry → gravar Excel]

  subgraph Escrita_com_Styling
    R --> R1[Salvar DF na aba "Registros"]
    R1 --> R2[Formatar: Data=dd/mm/yyyy; Data_txt & Hora como Texto (@)]
    R2 --> R3[Ajustar larguras de coluna]
    R3 --> R4[Criar/atualizar Tabela "Registros" cobrindo todo o range]
  end

  R4 --> S[Log: "✅ Excel atualizado: <caminho>"]
  S --> T([Fim])
