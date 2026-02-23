Attribute VB_Name = "ArmazenamentoDados"
Private Sub CommandButton1_Click()
Dim i, j, UltimaColuna, LinhaDemandaMaxima1, LinhaDemandaMaxima2, LinhaAfluencia1, LinhaAfluencia2, LinhaBusca, Linha1, Linha2, Linha3, Linha4, Linha5, Linha6, cont As Integer
Dim CelulaDemandaMaxima1, CelulaDemandaMaxima2, CelulaAfluencia1, CelulaAfluencia2, teste As Range
Dim usinas(1 To 500) As String
'Dim celula As Range

' ===== Declarações =====
Dim wsOrigem As Worksheet
Dim wsDestino As Worksheet
Dim linhaDestino As Long


Application.ScreenUpdating = False


'Copiar IPDO
    'Limpar planilha Importação
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
        
    Workbooks(2).Worksheets("IPDO").Activate
    Worksheets("IPDO").Columns("K:X").Copy
    Workbooks(1).Worksheets("Histórico de dados").Activate
    Worksheets("Importação").Range("A1:A1").PasteSpecial xlPasteValues
    Workbooks(2).Activate
    ActiveWorkbook.Close SaveChanges:=False

'Inserção de dados com referência fixa
UltimaColuna = Worksheets("Histórico de dados").Cells(1, Columns.Count).End(xlToLeft).Column + 1

Worksheets("Histórico de dados").Cells(1, UltimaColuna).Select
ActiveCell.FormulaR1C1 = Worksheets("Importação").Cells(6, 10)
i = 5
While i <= 1500

    If Not (IsEmpty(Worksheets("Histórico de dados").Cells(i, 1))) Then
    
        Worksheets("Histórico de dados").Cells(i, UltimaColuna) = Worksheets("Importação").Range(Worksheets("Histórico de dados").Cells(i, 1))
        Worksheets("Histórico de dados").Cells(i, UltimaColuna) = CDbl(Worksheets("Histórico de dados").Cells(i, UltimaColuna))
        
    End If

    i = i + 1

Wend

'Busca dos dados de demanda máxima
With Worksheets("Importação").Range("A1:L2000")
Set CelulaDemandaMaxima1 = .Find(What:="Dados de Dem. Máx.", LookIn:=xlValues)
LinhaDemandaMaxima1 = CelulaDemandaMaxima1.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaDemandaMaxima2 = .Find(What:="BALANÇO DE ENERGIA NA DEMANDA MÁXIMA DO SIN", LookIn:=xlValues)
LinhaDemandaMaxima2 = CelulaDemandaMaxima2.Row
End With

'Carregamento dos dados de demanda máxima
'SIN
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 3, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 1, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 4, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 2, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 7, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 3, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 8, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 4, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 9, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 5, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 10, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 6, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 2, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 7, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 11, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 8, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 12, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 9, 3)
'Itaipu
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 5, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 19, 7)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 6, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 20, 7)
'Norte

Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 15, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 15, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 16, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 12, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 17, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 13, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 18, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 14, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 19, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 16, 3)

'Ajustes devido a mudança no IPDO (inicio)
LinhaDemandaMaxima2 = LinhaDemandaMaxima2 + 1
'LinhaDemandaMaxima1 = LinhaDemandaMaxima1 + 1
'Ajustes devido a mudança no IPDO (termino)
'Nordeste
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 21, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 24, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 22, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 19, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 23, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 20, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 24, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 21, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 25, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 22, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 26, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 23, 3)
'Sudeste/Centro-Oeste
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 29, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 30, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 30, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 27, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 31, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 28, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 32, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 29, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 33, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 31, 3)
'Sul
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 36, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 37, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 37, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 34, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 38, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 35, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 39, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 36, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 40, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 38, 3)
'Intercâmbios líquidos
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 43, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 10, 7)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 44, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 11, 7)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 45, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 12, 7)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 46, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 13, 7)
'Intercâmbio internacional
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 49, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 28, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 50, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 22, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 51, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 23, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 52, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 24, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 53, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 25, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 54, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 26, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 48, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 27, 12)
'Demandas máximas atuais (MW)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 58, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 41, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 59, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 42, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 60, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 43, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 61, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 44, 3)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 62, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 46, 3)
'Demandas máximas atuais (hh:mm)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 64, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 41, 5)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 65, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 42, 5)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 66, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 43, 5)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 67, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 44, 5)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 68, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 46, 5)
'Demandas máximas históricas (MW)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 72, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 41, 10)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 73, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 42, 10)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 74, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 43, 10)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 75, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 44, 10)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 76, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 46, 10)
'Demandas máximas atuais (data)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 78, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 41, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 79, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 42, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 80, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 43, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 81, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 44, 12)
Worksheets("Histórico de dados").Cells(LinhaDemandaMaxima2 + 82, UltimaColuna) = Worksheets("Importação").Cells(LinhaDemandaMaxima1 + 46, 12)

'Busca dos dados hidráulicos
With Worksheets("Importação").Range("A1:L2000")
Set CelulaAfluencia1 = .Find(What:="Afluência", LookIn:=xlValues)
LinhaAfluencia1 = CelulaAfluencia1.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="DADOS HIDRÁULICOS - AFLUÊNCIAS", LookIn:=xlValues)
LinhaAfluencia2 = CelulaAfluencia2.Row
End With

'Consistência dos dados hidráulicos
i = LinhaAfluencia1 + 1
j = 1
While Worksheets("Importação").Cells(i, 3) <> "SE"
    If Not (IsNumeric(Worksheets("Importação").Cells(i, 3))) And Worksheets("Importação").Cells(i, 3) <> "Armaz" And Not (IsEmpty(Worksheets("Importação").Cells(i, 3))) Then
        
        With Worksheets("Histórico de dados").Range("C1:C2000")
            Set CelulaAfluencia2 = .Find(What:=Worksheets("Importação").Cells(i, 3), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If Not (TypeName(CelulaAfluencia2) = "Range") Then
                'ALERTAR ERRO
                usinas(j) = Worksheets("Importação").Cells(i, 3)
                j = j + 1
            End If
        End With
        
    End If
    i = i + 1
Wend
'ALERTA DE USINAS NÃO IDENFICADAS NA PLANILHA
Dim msg As String
Dim title As String
i = 1
msg = "Foram localizadas as seguintes UHEs no IPDO e não na planilha: " & vbCrLf
While usinas(i) <> ""
    msg = msg & usinas(i) & vbCrLf
    i = i + 1
Wend
title = "Alerta!"
If msg <> "Foram localizadas as seguintes UHEs no IPDO e não na planilha: " & vbCrLf Then
    response = MsgBox(msg, 16, title)
End If

'Busca dos dados hidráulicos - AFLUÊNCIAS
With Worksheets("Importação").Range("A1:L2000")
Set CelulaAfluencia1 = .Find(What:="Afluência", LookIn:=xlValues)
LinhaAfluencia1 = CelulaAfluencia1.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="DADOS HIDRÁULICOS - AFLUÊNCIAS", LookIn:=xlValues)
LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
j = 0

While Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 1) <> "BACIA"

    If Not (IsNumeric(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Not (IsEmpty(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3) = Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, 3) Then
        
        Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 6)
        
        j = j + 1
        
    End If

i = i + 1

Wend

'Busca dos dados hidráulicos - DEFLUÊNCIAS

With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="DADOS HIDRÁULICOS - DEFLUÊNCIAS", LookIn:=xlValues)
LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
j = 0

While Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 1) <> "BACIA"

    If Not (IsNumeric(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Not (IsEmpty(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3) = Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, 3) Then
        
        Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 7)
        
        j = j + 1
        
    End If

i = i + 1

Wend

'Busca dos dados hidráulicos - NÍVEL

With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="DADOS HIDRÁULICOS - NÍVEL", LookIn:=xlValues)
LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
j = 0

While Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 1) <> "BACIA"

    If Not (IsNumeric(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Not (IsEmpty(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3) = Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, 3) Then
        
        Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 8)
        
        j = j + 1
        
    End If

i = i + 1

Wend

'Busca dos dados hidráulicos - VERTIMENTO

With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="DADOS HIDRÁULICOS - VOLUME", LookIn:=xlValues)
LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
j = 0

While Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 1) <> "BACIA"

    If Not (IsNumeric(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Not (IsEmpty(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3) = Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, 3) Then
        
        Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 10)
        
        j = j + 1
        
    End If

i = i + 1

Wend

'Busca dos dados hidráulicos - VOLUME

With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="DADOS HIDRÁULICOS - VERTIMENTO", LookIn:=xlValues)
LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
j = 0

While Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 1) <> "BACIA"

    If Not (IsNumeric(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Not (IsEmpty(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3))) And Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 3) = Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, 3) Then
        
        If Len(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 12)) >= 4 Then
        
            Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + j, UltimaColuna) = Right(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 12), Len(Worksheets("Importação").Cells(LinhaAfluencia1 + 1 + i, 12)) - 4)
        
        End If
        
        j = j + 1
        
    End If

i = i + 1

Wend

'Busca dos dados de bacias
'Armazenamento
With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="DADOS POR BACIA HIDROGRÁFICA", LookIn:=xlValues)
    LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
With Worksheets("Importação").Range("A1:A2000")
    Set CelulaAfluencia1 = .Find(What:="Bacias", LookIn:=xlValues)
    LinhaBusca = CelulaAfluencia1.Row
End With


While Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + i, 2) <> "ENA dia bacia %"
    With Worksheets("Importação").Range("A" & LinhaBusca & ":A2000")
    Set CelulaAfluencia1 = .Find(What:=Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + i, 2), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If TypeName(CelulaAfluencia1) = "Range" Then
            LinhaAfluencia1 = CelulaAfluencia1.Row
            Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + i, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1, 3)
        Else
                'LISTA DE ALERTAS
                msg = "Não foi localizada a bacia hidrográfica " & Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + i, 2) & " no IPDO lido."  ' Define messagem.
                title = "Alerta!"
                response = MsgBox(msg, 16, title)
        End If
    End With
    
    
    i = i + 1
Wend

'ENA dia
With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="ENA dia bacia %", LookIn:=xlValues)
    LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
While Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + i, 2) <> "ENA armazenável bacia %"
    With Worksheets("Importação").Range("A" & LinhaBusca & ":A2000")
    Set CelulaAfluencia1 = .Find(What:=Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, 2), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If TypeName(CelulaAfluencia1) = "Range" Then
            LinhaAfluencia1 = CelulaAfluencia1.Row
            Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1, 5)
        End If
    End With
    
    
    i = i + 1
Wend

'ENA armazenável
With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="ENA armazenável bacia %", LookIn:=xlValues)
    LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
While Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + i, 2) <> "ENA bruta bacia %"
    With Worksheets("Importação").Range("A" & LinhaBusca & ":A2000")
    Set CelulaAfluencia1 = .Find(What:=Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, 2), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If TypeName(CelulaAfluencia1) = "Range" Then
            LinhaAfluencia1 = CelulaAfluencia1.Row
            Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1, 6)
        End If
    End With
    
    
    i = i + 1
Wend

'ENA bruta
With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="ENA bruta bacia %", LookIn:=xlValues)
    LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
While Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 2 + i, 2) <> "Geração verificada bacia MWmédios"
    With Worksheets("Importação").Range("A" & LinhaBusca & ":A2000")
    Set CelulaAfluencia1 = .Find(What:=Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, 2), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If TypeName(CelulaAfluencia1) = "Range" Then
            LinhaAfluencia1 = CelulaAfluencia1.Row
            Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1, 7)
        End If
    End With
    
    
    i = i + 1
Wend

'Geração verificada MWmédios
With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="Geração verificada bacia MWmédios", LookIn:=xlValues)
    LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
While Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, 2) <> "Geração programada bacia MWmédios"
    With Worksheets("Importação").Range("A" & LinhaBusca & ":A2000")
    Set CelulaAfluencia1 = .Find(What:=Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, 2), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If TypeName(CelulaAfluencia1) = "Range" Then
            LinhaAfluencia1 = CelulaAfluencia1.Row
            Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1, 8)
        End If
    End With
    
    
    i = i + 1
Wend

'Geração programada MWmédios
With Worksheets("Histórico de dados").Range("B1:B2000")
Set CelulaAfluencia2 = .Find(What:="Geração programada bacia MWmédios", LookIn:=xlValues)
    LinhaAfluencia2 = CelulaAfluencia2.Row
End With

i = 0
While Not (IsEmpty(Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, 2)))
    With Worksheets("Importação").Range("A" & LinhaBusca & ":A2000")
    Set CelulaAfluencia1 = .Find(What:=Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, 2), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If TypeName(CelulaAfluencia1) = "Range" Then
            LinhaAfluencia1 = CelulaAfluencia1.Row
            Worksheets("Histórico de dados").Cells(LinhaAfluencia2 + 1 + i, UltimaColuna) = Worksheets("Importação").Cells(LinhaAfluencia1, 10)
        End If
    End With
    
    
    i = i + 1
Wend

'Consistência dos dados de UHEs
With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="DADOS HIDRÁULICOS - AFLUÊNCIAS", LookIn:=xlValues)
Linha1 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="DADOS HIDRÁULICOS - DEFLUÊNCIAS", LookIn:=xlValues)
Linha2 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="DADOS HIDRÁULICOS - NÍVEL", LookIn:=xlValues)
Linha3 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="DADOS HIDRÁULICOS - VERTIMENTO", LookIn:=xlValues)
Linha4 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="DADOS HIDRÁULICOS - VOLUME", LookIn:=xlValues)
Linha5 = celula.Row
End With

i = 2
cont = 0
While Worksheets("Histórico de dados").Cells(Linha1 + i, 3) <> ""
If Not (StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 3), Worksheets("Histórico de dados").Cells(Linha2 + i, 3)) = 0 And StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 3), Worksheets("Histórico de dados").Cells(Linha3 + i, 3)) = 0 And StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 3), Worksheets("Histórico de dados").Cells(Linha4 + i, 3)) = 0 And StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 3), Worksheets("Histórico de dados").Cells(Linha5 + i, 3)) = 0) Then
    cont = cont + 1
End If
i = i + 1
Wend

If cont >= 1 Then
    msg = "A lista de UHEs não está consistente (Afluências, Defluências, Nível, Vertimento e Volume)!"  ' Define messagem.
    title = "Alerta!"
    response = MsgBox(msg, 16, title)
End If

'Consistência dos dados de Bacias
With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="Armazenamento bacia %", LookIn:=xlValues)
Linha1 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="ENA dia bacia %", LookIn:=xlValues)
Linha2 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="ENA armazenável bacia %", LookIn:=xlValues)
Linha3 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="ENA bruta bacia %", LookIn:=xlValues)
Linha4 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="Geração verificada bacia MWmédios", LookIn:=xlValues)
Linha5 = celula.Row
End With

With Worksheets("Histórico de dados").Range("B1:B2000")
Set celula = .Find(What:="Geração programada bacia MWmédios", LookIn:=xlValues)
Linha6 = celula.Row
End With


i = 1
cont = 0
While Worksheets("Histórico de dados").Cells(Linha1 + i, 2) <> ""
If Not (StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 2), Worksheets("Histórico de dados").Cells(Linha2 + i, 2)) = 0 And StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 2), Worksheets("Histórico de dados").Cells(Linha3 + i, 2)) = 0 And StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 2), Worksheets("Histórico de dados").Cells(Linha4 + i, 2)) = 0 And StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 2), Worksheets("Histórico de dados").Cells(Linha5 + i, 2)) = 0 And StrComp(Worksheets("Histórico de dados").Cells(Linha1 + i, 2), Worksheets("Histórico de dados").Cells(Linha6 + i, 2)) = 0) Then
    cont = cont + 1
End If
i = i + 1
Wend

If cont >= 1 Then
    msg = "A lista de Bacias não está consistente (Armazenamento bacia %, ENA dia bacia %, ENA armazenável bacia %, ENA bruta bacia %, Geração verificada bacia MWmédios e Geração programada bacia MWmédios)!"  ' Define messagem.
    title = "Alerta!"
    response = MsgBox(msg, 16, title)
End If

'inserir em Histórico de dados
    '    Set wsOrigem = Sheets("importação")
    '    Set wsDestino = Sheets("Dados bi")
    '
    '
    'linhaDestino = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row + 1
    '
    '    wsDestino.Range("A" & linhaDestino & ":A" & linhaDestino + 3).Value = wsOrigem.Range("A62:A65").Value
    '
    '    wsDestino.Range("B" & linhaDestino & ":B" & linhaDestino + 3).Value = wsOrigem.Range("N6").Value
    '
    '    wsDestino.Range("C" & linhaDestino & ":C" & linhaDestino + 3).Value = wsOrigem.Range("C62:C65").Value
    '
    '    wsDestino.Range("D" & linhaDestino & ":D" & linhaDestino + 3).Value = wsOrigem.Range("E62:E65").Value
    '
    '    wsDestino.Range("E" & linhaDestino & ":E" & linhaDestino + 3).Value = wsOrigem.Range("F62:F65").Value
    '
    '    wsDestino.Range("F" & linhaDestino & ":F" & linhaDestino + 3).Value = wsOrigem.Range("G62:G65").Value
    '
    '    wsDestino.Range("G" & linhaDestino & ":G" & linhaDestino + 3).Value = wsOrigem.Range("H62:H65").Value
    

End Sub

