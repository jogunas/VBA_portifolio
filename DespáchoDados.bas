Attribute VB_Name = "DespáchoDados"
Sub EnviarDadosParaDespacho()

    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsInserir As Worksheet
    Dim linhaDestino As Long
    Dim ultimaLinha As Long
    Dim ultimaData As Date
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoPicture Then
            shp.Delete
        End If
    Next shp

    Set wsOrigem = Sheets("BDD")
    Set wsDestino = Sheets("Despacho")
    Set wsInserir = Sheets("Inserir")
    
    
    ' Descobre a última linha preenchida na coluna B
    ultimaLinha = wsDestino.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' Pega o valor da última data
    ultimaData = wsDestino.Cells(ultimaLinha, "B").Value + 1
    wsInserir.Cells(5, "N").Value = " Última data adicinada: " & ultimaData
    
    ' Escreve o próximo dia na linha seguinte
    wsDestino.Cells(ultimaLinha + 1, "B").Value = ultimaData
    wsDestino.Cells(ultimaLinha + 1, "I").Value = ultimaData
    wsDestino.Cells(ultimaLinha + 1, "P").Value = ultimaData
    
    soma_ome = "=SOMA(OME[@[SE-CO]:[Norte]])"
    wsDestino.Cells(ultimaLinha + 1, "G").Formula2Local = soma_ome
    wsDestino.Cells(ultimaLinha + 1, "G").NumberFormat = "0"

    soma_rel = "=SOMA(REL[@[SE-CO]:[Norte]])"
    wsDestino.Cells(ultimaLinha + 1, "N").Formula2Local = soma_rel
    wsDestino.Cells(ultimaLinha + 1, "N").NumberFormat = "0"
    
    wsDestino.Cells(ultimaLinha + 1, "Q").Formula2Local = "=SOMA(OME[@[SE-CO]];REL[@[SE-CO]])"
    wsDestino.Cells(ultimaLinha + 1, "Q").NumberFormat = "0"
    
    wsDestino.Cells(ultimaLinha + 1, "R").Formula2Local = "=SOMA(OME[@Sul];REL[@Sul])"
    wsDestino.Cells(ultimaLinha + 1, "R").NumberFormat = "0"
    
    wsDestino.Cells(ultimaLinha + 1, "S").Formula2Local = "=SOMA(OME[@Nordeste];REL[@Nordeste])"
    wsDestino.Cells(ultimaLinha + 1, "S").NumberFormat = "0"
    
    wsDestino.Cells(ultimaLinha + 1, "T").Formula2Local = "=SOMA(OME[@Norte];REL[@Norte])"
    wsDestino.Cells(ultimaLinha + 1, "T").NumberFormat = "0"
    
    wsDestino.Cells(ultimaLinha + 1, "U").Formula2Local = "=SOMA(OME_REL[@[SE-CO]:[Norte]])"
    wsDestino.Cells(ultimaLinha + 1, "U").NumberFormat = "0"
    wsDestino.Cells(ultimaLinha + 1, "U").Font.Bold = False

    ' Encontra a primeira linha vazia na coluna C da aba Despacho
    linhaDestino = wsDestino.Cells(2, 3).End(xlDown).Row + 1

    ' Copia os valores para as colunas
    'OME
    'SU_CO
    With wsDestino.Cells(linhaDestino, 3)
        .Value = wsOrigem.Range("B58").Value
        .NumberFormat = wsOrigem.Range("B58").NumberFormat
    End With
    'S
    With wsDestino.Cells(linhaDestino, 4)
        .Value = wsOrigem.Range("G19").Value
        .NumberFormat = wsOrigem.Range("G19").NumberFormat
    End With
    'NE
    With wsDestino.Cells(linhaDestino, 5)
        .Value = wsOrigem.Range("L37").Value
        .NumberFormat = wsOrigem.Range("L37").NumberFormat
    End With
    'NO
    With wsDestino.Cells(linhaDestino, 6)
        .Value = wsOrigem.Range("Q40").Value
        .NumberFormat = wsOrigem.Range("Q40").NumberFormat
    End With
    
    ' REL
    'SU_CO
    
    With wsDestino.Cells(linhaDestino, 10)
        .Value = wsOrigem.Range("C58").Value
        .NumberFormat = wsOrigem.Range("C58").NumberFormat
    End With
    'S
    With wsDestino.Cells(linhaDestino, 11)
        .Value = wsOrigem.Range("H19").Value
        .NumberFormat = wsOrigem.Range("H19").NumberFormat
    End With
    'NE
    With wsDestino.Cells(linhaDestino, 12)
        .Value = wsOrigem.Range("M37").Value
        .NumberFormat = wsOrigem.Range("M37").NumberFormat
    End With
    'NO
    With wsDestino.Cells(linhaDestino, 13)
        .Value = wsOrigem.Range("R40").Value
        .NumberFormat = wsOrigem.Range("R40").NumberFormat
    End With
    
    
Dim alvo As Range
Dim linhaAlvo As Long

linhaAlvo = ultimaLinha + 1
Set alvo = wsDestino.Range("B" & linhaAlvo & ":U" & linhaAlvo)
    
    If Weekday(ultimaData, vbSunday) = 1 Then

With alvo.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
wsDestino.Cells(linhaAlvo, "H").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
wsDestino.Cells(linhaAlvo, "O").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
End If
    'criação da tabela
    
    

   ' MsgBox "Valores foram enviados para a aba Despacho com sucesso!"

End Sub

