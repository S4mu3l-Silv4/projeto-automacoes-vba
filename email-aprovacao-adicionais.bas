Sub EnviarEmailSolicitacaoAdicional()
    Dim range As Range
    Dim horaAtual As Integer
    Dim saudacao As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim assinatura As String
    Dim corpoEmail As String
    Dim wordEditor As Object
    Dim tabela As Object
    Dim rangeTabela As Object
    On Error Resume Next
    Set range = Sheets("Planilha1").Range("A1").CurrentRegion.Resize(, 11).SpecialCells(xlCellTypeVisible)
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
        .To = "exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com"
        .CC = ""
        .Subject = "Site(s) - Adicional por negociação"
        .Display
        assinatura = .HTMLBody
        corpoEmail = "<div style='font-family:Calibri;font-size:11pt;'>" & _
            saudacao & "<br><br>" & _
            "Xxx, solicito a sua autorização para podermos seguir com os valores abaixo: <br>" & _
        "</div><br>"
        .HTMLBody = corpoEmail & assinatura
        Set wordEditor = .GetInspector.WordEditor
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
            Set rangeTabela = tabela.Range
            rangeTabela.Collapse 0
            rangeTabela.InsertParagraphAfter
        End If
    End With
    Set rangeTabela = Nothing
    Set tabela = Nothing
    Set range = Nothing
    Set wordEditor = Nothing
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub