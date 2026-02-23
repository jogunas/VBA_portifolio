Attribute VB_Name = "DespachoTabela"
Option Explicit

Sub CriarTabela()
Dim wsOrigem As Worksheet
Dim wsDestino As Worksheet
Dim Seg_tab As Date
Dim Seg_lin As Long
Dim dias As Long

Set wsOrigem = ThisWorkbook.Worksheets("Despacho")
Set wsDestino = ThisWorkbook.Worksheets("inserir")

Seg_tab = wsDestino.Cells(1, "Q").Value
Seg_lin = wsOrigem.Columns(2).Find(What:=Seg_tab).Row
dias = wsDestino.Cells(2, "Q").Value


Dim cel As Range, rngBloco As Range
Set cel = wsOrigem.Cells(Seg_lin, "B")

' copia de Despacho

Set rngBloco = wsOrigem.Range(cel, cel.Offset(dias - 1, 19))

' Apaga de X a AQ

wsDestino.Range("X:AQ").Clear
' cola a tabela
wsDestino.Range("X1").Resize(rngBloco.Rows.Count, rngBloco.Columns.Count).Value = rngBloco.Value


' cabeçalho


With wsDestino.Cells(1, "R")
    .Value = "De " & Format(Seg_tab, "dd/mm") & " a " & Format(Seg_tab + dias - 1, "dd/mm") & " (MWmed)"
End With


End Sub
