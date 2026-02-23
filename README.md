# 📊 MME – Projetos VBA

Este repositório reúne automações desenvolvidas em **VBA (Visual Basic for Applications)** para apoiar a análise e geração de relatórios relacionados ao setor elétrico, especialmente voltados ao **Informe Semanal**, **Despacho Térmico**, **ENAs**, **EARs**, e demais dados do SIN.

Os códigos deste repositório foram extraídos de planilhas .xlsm utilizadas para automatizar operações internas e acelerar etapas repetitivas de análise.

---

## 📁 Estrutura do Repositório
MME-projetos/
├── 01-InformeSemanal/
│   ├── src/              → módulos .bas/.cls/.frm exportados do Excel
│   ├── exemplos/         → (opcional) planilhas ou prints explicativos
│   └── README.md         → documentação específica do projeto
└── README.md             → este arquivo
Cada pasta representa um projeto independente escrito em VBA.

---

# 📌 Projeto 01 – Informe Semanal (VBA + Excel)

O objetivo deste projeto é **automatizar a geração do Informe Semanal**, consolidando:

- Despacho térmico (**OME, REL e OME+REL**)  
- Indicadores de **ENA** por subsistema  
- Indicadores de **EAR** por subsistema  
- Exportações/importações por dia  
- Resumo semanal automático

Este projeto opera com as seguintes abas:

- **Informe** → resumo geral  
- **Inserir** → entrada dos dados semanais  
- **BDD** → base de dados de usinas  
- **Despacho** → cálculos de ordem de mérito e restrições  

---

## ⚙️ Como Executar

1. Baixe o arquivo original `.xlsm` (sem dados sensíveis).  
2. Habilite macros no Excel.  
3. Abra `Alt + F11`.  
4. Importe os módulos da pasta `src/`.  
5. Execute a macro principal (exemplo):
Sub AtualizarInformeSemanal()
*(Se o nome correto da macro principal for outro, substitua acima.)*

---

## 🧩 Principais Módulos (com base nos arquivos enviados)

| Módulo | Função esperada |
|--------|------------------|
| `ArmazenamentoDados.bas` | cálculos de despacho térmico e consolidação por subsistema |
| `ArmazenamentoTabela` | atualização dos valores de ENA semanal |
| `DespachoTabela` | cálculo e comparação dos níveis de EAR |
| `DespachoDados` | funções auxiliares usadas em vários módulos |
| `TextoExportação`| exportação do informe em PDF |
| `TextoImportação`

> Ajuste a tabela acima conforme os nomes reais dos seus módulos.

---

## 🧪 Dependências

- Excel com macros habilitadas  
- Permissão ativada:  
**Arquivo → Opções → Central de Confiabilidade → Configurações → “Confiar no acesso ao modelo de objeto do projeto VBA”**

---

## 📬 Contato

João Guilherme  
*(adicione seu LinkedIn, email, portfólio etc.)*

---

## 📄 Licença

Este repositório utiliza a licença **MIT** por padrão.  
Pode alterar se preferir.
