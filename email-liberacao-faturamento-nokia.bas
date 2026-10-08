Sub EnviarEmailLiberacaoFaturamentoNokia()
    Dim range As range
    Dim horaAtual As Integer
    Dim saudacao As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim assinatura As String
    Dim corpoEmail As String
    Dim wordEditor As Object
    Dim tabela As Object
    Dim rangeTabela As Object
    Dim caminhoPasta As String
    Dim arquivoPDF As String
    On Error Resume Next
    Set range = Sheets("faturamento-nokia").range("A1").CurrentRegion.Resize(, 7).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
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
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    Set outlookMail = outlookApp.CreateItem(0)
    With outlookMail
        .To = "exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com"
        .CC = "exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com"
        .Subject = "Liberação para faturamento Nokia TCO " & _
        Sheets("faturamento-nokia").range("A2").Value
        .Display
        assinatura = .HTMLBody
        corpoEmail = "<div style='font-family:Calibri;font-size:11pt;'>" & _
            saudacao & "<br><br>" & _
            "Segue liberação para faturamento liberada pela Nokia: <br>" & _
            "</div>"
        .HTMLBody = corpoEmail & assinatura
        Set wordEditor = .GetInspector.wordEditor
        range.Copy
        With wordEditor.Application.Selection
            .HomeKey 6
            .MoveDown Unit:=5, Count:=4
            .PasteExcelTable False, False, True
        End With
        If wordEditor.Tables.Count > 0 Then
            Set tabela = wordEditor.Tables(wordEditor.Tables.Count)
            With tabela
                .AutoFitBehavior 1
            End With
            Set rangeTabela = tabela.range
            rangeTabela.Collapse 0
            rangeTabela.InsertParagraphAfter
        End If
        caminhoPasta = ThisWorkbook.Path
        If Right(caminhoPasta, 1) <> "\" Then
            caminhoPasta = caminhoPasta & "\"
        End If
        arquivoPDF = Dir(caminhoPasta & "*.pdf")
        Do While arquivoPDF <> ""
            .Attachments.Add caminhoPasta & arquivoPDF
            arquivoPDF = Dir
        Loop
    End With
    Set rangeTabela = Nothing
    Set tabela = Nothing
    Set range = Nothing
    Set wordEditor = Nothing
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub