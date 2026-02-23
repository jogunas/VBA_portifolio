# 📊 MME – Projetos VBA

Este repositório reúne automações desenvolvidas em **VBA (Visual Basic for Applications)** para apoiar a análise e geração de relatórios relacionados ao setor elétrico, especialmente voltados ao **Informe Semanal**, **Despacho Térmico**, **ENAs**, **EARs**, e demais dados do SIN.

Os códigos deste repositório foram extraídos de planilhas .xlsm utilizadas para automatizar operações internas e acelerar etapas repetitivas de análise.

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
4. Execute a macro principal (exemplo):
---

## 🧩 Principais Módulos (com base nos arquivos enviados)

| Módulo | Função esperada |
|--------|------------------|
| [ArmazenamentoDados](https://github.com/jogunas/VBA_portifolio/blob/main/ArmazenamentoDados.bas) |  |
| `ArmazenamentoTabela` |  |
| `DespachoTabela` |  |
| `DespachoDados` |  |
| `TextoExportação`|  |
| `TextoImportação` |  |

> Ajuste a tabela acima conforme os nomes reais dos seus módulos.

---


## 📬 Contato

João Guilherme  
*(adicione seu LinkedIn, email, portfólio etc.)*
