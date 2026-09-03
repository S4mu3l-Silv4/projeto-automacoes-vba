Sub EnviarEmailCobrancaVOVivo()
    Dim ws As Worksheet
    Dim colunas As Variant
    Dim ultimaLinha As Long
    Dim horaAtual As Integer
    Dim saudacao As String
    Dim tabelaHTML As String
    Dim i As Long
    Dim linha As Long
    Dim valor As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim assinatura As String
    Dim corpoEmail As String
    Set ws = Sheets("aba-base-dados-vos")
    ws.range("G:G").NumberFormat = "0"
    colunas = Array("A", "E", "G", "I", "V", "AM", "AW")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    horaAtual = Hour(Now)
    If horaAtual >= 18 Then
        saudacao = "Boa noite, "
    ElseIf horaAtual >= 12 Then
        saudacao = "Boa tarde, "
    ElseIf horaAtual >= 6 Then
        saudacao = "Bom dia, "
    Else
        saudacao = "Boa noite, "
    End If
    tabelaHTML = "<table border='1' cellpadding='5' cellspacing='0' " & _
        "style='border-collapse:collapse;font-family:Calibri;font-size:11pt;'>"
    tabelaHTML = tabelaHTML & "<tr style='font-weight:bold;'>"
    For i = LBound(colunas) To UBound(colunas)
        tabelaHTML = tabelaHTML & _
        "<td>" & HtmlEncode(CStr(ws.range(colunas(i) & "1").Value)) & "</td>"
    Next i
    tabelaHTML = tabelaHTML & "</tr>"
    For linha = 2 To ultimaLinha
        If ws.Rows(linha).Hidden = False Then
            tabelaHTML = tabelaHTML & "<tr>"
            For i = LBound(colunas) To UBound(colunas)
                valor = ws.range(colunas(i) & linha).Text
                tabelaHTML = tabelaHTML & _
                "<td>" & HtmlEncode(valor) & "</td>"
            Next i
            tabelaHTML = tabelaHTML & "</tr>"
        End If
    Next linha
    tabelaHTML = tabelaHTML & "</table>"
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    Set outlookMail = outlookApp.CreateItem(0)
    With outlookMail
        .To = "exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com"
        .CC = "exemplo@xxx.com; exemplo@xxx.com"
        .Subject = "Cobrança de VO - Vivo"
        .Display
        assinatura = .HTMLBody
        corpoEmail = _
            "<div style='font-family:Calibri;font-size:11pt;'>" & _
                saudacao & "<br><br>" & _
                "Poderiam verificar as VOs abaixo por gentileza? " & _
                "Esses casos estão com pendência de aprovação: <br><br>" & _
                tabelaHTML & _
            "</div>"
        .HTMLBody = corpoEmail & assinatura
    End With
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    Set ws = Nothing
End Sub
Function HtmlEncode(ByVal texto As String) As String
    texto = Replace(texto, "&", "&amp;")
    texto = Replace(texto, "<", "&lt;")
    texto = Replace(texto, ">", "&gt;")
    texto = Replace(texto, """", "&quot;")
    HtmlEncode = texto
End Function