Attribute VB_Name = "TextoImportação"
Sub GeraTextoImportacoes()
    Dim ws As Worksheet: Set ws = ActiveSheet

    Const HEADER_ROW As Long = 3
    Const COL_DATA As Long = 1           ' Coluna A
    Const HDR_ARG As String = "argentina"
    Const HDR_URU As String = "uruguai"

    Dim lastRow As Long, lastCol As Long
    Dim colArg As Long, colUru As Long
    Dim c As Long, r As Long
    Dim h As String
    Dim d As Variant
    Dim v As Double
    Dim txtArg As String, txtUru As String
    Dim finalText As String

    lastRow = 10
    lastCol = 3
    
    colArg = 0
    colUru = 0

    ' Descobre colunas pelos cabeçalhos na linha 3
    For c = 1 To lastCol
        h = LCase$(Trim$(Replace(Replace(CStr(ws.Cells(HEADER_ROW, c).Value), "–", "-"), "  ", " ")))
        If h = LCase$(HDR_ARG) Then colArg = c
        If h = LCase$(HDR_URU) Then colUru = c
    Next c

    If colArg = 0 Or colUru = 0 Then
        MsgBox "Cabeçalhos não encontrados na linha " & HEADER_ROW, vbExclamation
        Exit Sub
    End If

    ' Varre as linhas de dados
    For r = HEADER_ROW + 1 To lastRow
        d = ws.Cells(r, COL_DATA).Value
        If IsDate(d) Then
            ' Argentina
            v = GetNumeric(ws.Cells(r, colArg).Value)
            If v > 0 Then
                If Len(txtArg) > 0 Then txtArg = txtArg & ", "
                txtArg = txtArg & Format$(d, "d/mm") & " (" & FormatNumberClean(v) & " MWmed)"
            End If
            ' Uruguai
            v = GetNumeric(ws.Cells(r, colUru).Value)
            If v > 0 Then
                If Len(txtUru) > 0 Then txtUru = txtUru & ", "
                txtUru = txtUru & Format$(d, "d/mm") & " (" & FormatNumberClean(v) & " MWmed)"
            End If
        End If
    Next r

    ' Monta a frase final
Dim argCount As Long, uruCount As Long
    argCount = CountItems(txtArg)
    uruCount = CountItems(txtUru)

If argCount = 0 And uruCount = 0 Then
    finalText = "Sem importações na semana"
Else
    finalText = "Houve importação "
    
    ' --- Argentina: no/nos + dia/dias ---
    If argCount > 0 Then
        finalText = finalText & "da Argentina " & IIf(argCount = 1, "no dia ", "nos dias ") & txtArg
    End If
    
    ' Conector " e " se tiver os dois países
    If argCount > 0 And uruCount > 0 Then
        finalText = finalText & " e "
    End If
    
    ' --- Uruguai: (mantive sem no/nos; se quiser igual à Argentina, aviso abaixo) ---
    If uruCount > 0 Then
        finalText = finalText & "do Uruguai " & IIf(uruCount = 1, "no dia ", "nos dias ") & txtUru
    End If
End If

' --- Garantir ponto final e evitar vírgula no fim ---
finalText = Trim$(finalText)
If Right$(finalText, 1) = "," Then finalText = Left$(finalText, Len(finalText) - 1)
If Right$(finalText, 1) <> "." Then finalText = finalText & "."
    
    ' --- garantir ponto final, sem vírgula no fim e sem espaços extras ---
finalText = Trim$(finalText)
If Right$(finalText, 1) = "," Then finalText = Left$(finalText, Len(finalText) - 1)
If Right$(finalText, 1) <> "." Then finalText = finalText & "."

    ' Saída
    ws.Range("I3").Value = finalText

End Sub

' ---- Utilitários ----

Private Function GetNumeric(v As Variant) As Double
    Dim s As String: s = Trim$(CStr(v))
    If s = "" Then
        GetNumeric = 0
        Exit Function
    End If
    If Not IsNumeric(s) Then
        s = Replace(s, " ", "")
        s = Replace(s, ",", Application.DecimalSeparator)
        s = Replace(s, ".", Application.DecimalSeparator)
    End If
    If IsNumeric(s) Then
        GetNumeric = CDbl(s)
    Else
        GetNumeric = 0
    End If
End Function

Private Function FormatNumberClean(x As Double) As String
    If x = Fix(x) Then
        FormatNumberClean = CStr(Fix(x))
    Else
        FormatNumberClean = Format$(x, "0.############")
    End If
End Function

Private Function CountItems(ByVal s As String) As Long
    s = Trim$(s)
    If Len(s) = 0 Then
        CountItems = 0
    Else
        CountItems = UBound(Split(s, ",")) + 1
    End If
End Function


