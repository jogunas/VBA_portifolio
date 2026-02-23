Attribute VB_Name = "Linhadespacho"
Sub linhasDomingo()
Dim wslinha As Worksheet
Dim linhaatual As Long, primeiraLinha As Long, ultimaLinha As Long
Dim colData As Long, ultimaColLinha As Long
Dim dt As Date
Dim rngLinha As Range
Dim dia As String

Set wslinha = Sheets("Despacho")
'Parametros
primeiraLinha = 3
ultimaLinha = 884
dia = vbSunday   ' último dia da semana do despacho
colData = wslinha.Columns("B").Column

Application.ScreenUpdating = False
Application.EnableEvents = False

    On Error GoTo TrataErro


For linhaatual = primeiraLinha To ultimaLinha
Set rngLinha = wslinha.Range(wslinha.Cells(linhaatual, "B"), wslinha.Cells(linhaatual, "U"))
                With rngLinha.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(200, 200, 200)
                End With
                With rngLinha.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(200, 200, 200)
                End With
                With rngLinha.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(200, 200, 200)
                End With
' wslinha.Range(wslinha.Cells(linhaatual, "B"), wslinha.Cells(linhaatual, "U")).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone

 ' Verifica se há uma data válida na coluna B dessa linha

        If IsDate(wslinha.Cells(linhaatual, colData).Value) Then
            dt = CDate(wslinha.Cells(linhaatual, colData).Value)

            ' É domingo?
            If Weekday(dt, dia) = 1 Then
                ' A linha pode ter conteúdo até qualquer coluna; achamos a última usada
                ultimaColLinha = wslinha.Cells(linhaatual, wslinha.Columns.Count).End(xlToLeft).Column
                If ultimaColLinha < 1 Then ultimaColLinha = 1

                ' Aplica borda inferior MÉDIA (xlMedium)
                With rngLinha.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With
                wslinha.Cells(linhaatual, "H").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
                wslinha.Cells(linhaatual, "O").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
            Else
                ' (Opcional) remover borda nas linhas que não são domingo
                'wslinha.Range(wslinha.Cells(linhaatual, 1), wslinha.Cells(linhaatual, wslinha.Columns.Count).End(xlToLeft)).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
            End If
        End If
    Next linhaatual

Saida:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

TrataErro:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Erro em linhasDomingo: " & Err.Description, vbExclamation
End Sub

