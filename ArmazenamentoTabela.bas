Attribute VB_Name = "ArmazenamentoTabela"
Sub Atualiza_comparacao_reservatorios()

Workbooks(1).Worksheets("Histórico de dados").Activate
Workbooks(1).Worksheets("Tabela de armazenamentos").Activate


'dia que começa o informe
data_segunda = Format(Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(1, "K").Value, "dd/mm/yyyy")

'dia que acaba o informe
data_domingo = Format(Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(2, "K").Value, "dd/mm/yyyy")

' ultimo dia do mes anterior
data_m_1 = Format(DateSerial(Year(data_domingo), Month(data_domingo), 0), "dd/mm/yyyy")

' ultimo dia de 2 meses atrás
data_m_2 = Format(DateSerial(Year(data_domingo), Month(data_domingo) - 1, 0), "dd/mm/yyyy")

data_ano_passado = Format(DateSerial(Year(data_domingo) - 1, Month(data_domingo), Day(data_domingo)), "dd/mm/yyyy")

data_2021 = Format(DateSerial(2021, Month(data_domingo), Day(data_domingo)), "dd/mm/yyyy")

'procura colunas

coluna_segunda = Workbooks(1).Worksheets("Histórico de dados").Rows(1).Find(What:=data_segunda).Column
coluna_domingo = Workbooks(1).Worksheets("Histórico de dados").Rows(1).Find(What:=data_domingo).Column
coluna_passado = Workbooks(1).Worksheets("Histórico de dados").Rows(1).Find(What:=data_ano_passado).Column
coluna_2021 = Workbooks(1).Worksheets("Histórico de dados").Rows(1).Find(What:=data_2021).Column
coluna_m_1 = Workbooks(1).Worksheets("Histórico de dados").Rows(1).Find(What:=data_m_1).Column
coluna_m_2 = Workbooks(1).Worksheets("Histórico de dados").Rows(1).Find(What:=data_m_2).Column


'Cabeçalho
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(3, "D").Value = "Armazenamento em " & data_segunda
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(3, "E").Value = "Armazenamento em " & data_domingo

'Água Vermelha
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(4, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(955, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(4, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(955, coluna_domingo).Value / 100

'Barra Grande
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(5, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1036, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(5, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1036, coluna_domingo).Value / 100

'Emborcação
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(6, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(937, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(6, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(937, coluna_domingo).Value / 100

'Foz do Areia
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(7, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1022, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(7, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1022, coluna_domingo).Value / 100

'Furnas
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(8, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(947, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(8, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(947, coluna_domingo).Value / 100

'I.Solteira
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(9, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(991, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(9, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(991, coluna_domingo).Value / 100

'Itumbiara
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(10, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(938, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(10, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(938, coluna_domingo).Value / 100

'Marimbondo
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(11, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(950, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(11, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(950, coluna_domingo).Value / 100

'Nova Ponte
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(12, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(931, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(12, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(931, coluna_domingo).Value / 100

'Santo Santiago
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(13, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1024, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(13, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1024, coluna_domingo).Value / 100

'São Simão
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(14, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(940, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(14, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(940, coluna_domingo).Value / 100

'Serra da Mesa'
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(15, "D").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1048, coluna_segunda).Value / 100
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(15, "E").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(1048, coluna_domingo).Value / 100

'Itaipu
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(16, "D").Value = Format(Workbooks(1).Worksheets("Histórico de dados").Cells(668, coluna_segunda).Value, "0.0") & " m"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(16, "E").Value = Format(Workbooks(1).Worksheets("Histórico de dados").Cells(668, coluna_domingo).Value, "0.0") & " m"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(16, "F").Value = Format(Workbooks(1).Worksheets("Histórico de dados").Cells(668, coluna_domingo).Value - Workbooks(1).Worksheets("Histórico de dados").Cells(668, coluna_segunda).Value, "0.0") & " m"


'Cabeçalho
' ENA
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "S").Value = "ENA em " & StrConv(Format(data_m_2, "mmm/yyyy"), vbProperCase)
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "T").Value = "ENA em " & StrConv(Format(data_m_1, "mmm/yyyy"), vbProperCase)
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "U").Value = "ENA em " & data_domingo
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "V").Value = "ENA em " & data_ano_passado
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "W").Value = "ENA em " & data_2021

Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "R").Value = "Subsistemas"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "R").Value = "SE-CO"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "R").Value = "Sul"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "R").Value = "Nordeste"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "R").Value = "Norte"

' EAR
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "L").Value = "EAR em " & data_segunda
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "M").Value = "EAR em " & data_domingo
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "N").Value = "Evolução"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "O").Value = "EAR em " & data_ano_passado
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "P").Value = "EAR em " & data_2021

Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(18, "K").Value = "Subsistemas"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "K").Value = "SE-CO"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "K").Value = "Sul"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "K").Value = "Nordeste"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "K").Value = "Norte"
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(23, "K").Value = "Total"





'Set IPDO_antigo = Workbooks("IPDO-" & Format(data_antiga, "dd-mm-yyyy") & ".xlsm").Worksheets("IPDO")
'Set IPDO_recente = Workbooks("IPDO-" & Format(data_recente, "dd-mm-yyyy") & ".xlsm").Worksheets("IPDO")

'IPDO_antigo.Activate
'IPDO_recente.Activate


'EAR


'SE/CO

Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "L").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(178, coluna_segunda).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "M").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(178, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "O").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(178, coluna_passado).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "P").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(178, coluna_2021).Value


'S
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "L").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(177, coluna_segunda).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "M").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(177, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "O").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(177, coluna_passado).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "P").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(177, coluna_2021).Value


'NE
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "L").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(176, coluna_segunda).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "M").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(176, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "O").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(176, coluna_passado).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "P").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(176, coluna_2021).Value


'N
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "L").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(175, coluna_segunda).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "M").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(175, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "O").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(175, coluna_passado).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "P").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(175, coluna_2021).Value


'SIN
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(23, "L").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(179, coluna_segunda).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(23, "M").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(179, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(23, "O").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(179, coluna_passado).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(23, "P").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(179, coluna_2021).Value


'ENA

'SE/CO
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "S").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_2).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "T").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_1).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "U").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "V").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_2021).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(19, "W").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_passado).Value

'S
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "S").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_2).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "T").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_1).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "U").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(167, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "V").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(167, coluna_2021).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(20, "W").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(167, coluna_passado).Value

'NE
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "S").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_2).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "T").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_1).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "U").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(166, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "V").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(166, coluna_2021).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(21, "W").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(166, coluna_passado).Value

'N
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "S").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_2).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "T").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(168, coluna_m_1).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "U").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(165, coluna_domingo).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "V").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(165, coluna_2021).Value
Workbooks(1).Worksheets("Tabela de armazenamentos").Cells(22, "W").Value = Workbooks(1).Worksheets("Histórico de dados").Cells(165, coluna_passado).Value


'Workbooks("IPDO-" & Format(data_antiga, "dd-mm-yyyy") & ".xlsm").Close SaveChanges:=False
'Workbooks("IPDO-" & Format(data_recente, "dd-mm-yyyy") & ".xlsm").Close SaveChanges:=False


End Sub


