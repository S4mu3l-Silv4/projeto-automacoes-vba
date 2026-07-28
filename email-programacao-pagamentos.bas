Sub EnviarEmailProgramacaoPagamentos()
    Dim caminhoArquivo As String
    Dim range As Range
    Dim horaAtual As Integer
    Dim saudacao As String
    Dim numSemana As Integer
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim assinatura As String
    Dim corpoEmail As String
    Dim wordEditor As Object
    caminhoArquivo = "C:\Users\PREMCELL\Downloads\controle-CO\FSP - Form. Solicitação de Pagamentos - v2026_1_CO.xlsm"
    On Error Resume Next
    Set range = Sheets("FSP - Pagamentos Premcell").Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
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
    numSemana = DatePart("ww", Date, vbMonday, vbFirstFourDays)
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    Set outlookMail = outlookApp.CreateItem(0)
    With outlookMail
        .To = "exemplo@xxx.com; exemplo@xxx.com"
        .CC = "exemplo@xxx.com; exemplo@xxx.com"
        .Subject = "Programação de Pagamentos - CO - W" & numSemana
        .Attachments.Add caminhoArquivo
        .Display
        assinatura = .HTMLBody
        corpoEmail = "<div style='font-family:Calibri;font-size:11pt;'>" & _
            saudacao & "<br><br>" & _
            "Segue a Programação de Pagamentos: <br>" & _
        "</div>"
        .HTMLBody = corpoEmail & assinatura
        Set wordEditor = .GetInspector.WordEditor
        range.Copy
        With wordEditor.Application.Selection
            .HomeKey 6
            .MoveDown Unit:=5, Count:=4
            .PasteExcelTable False, False, True
        End With
        If wordEditor.Tables.Count > 0 Then
            With wordEditor.Tables(wordEditor.Tables.Count)
                .AutoFitBehavior 1
            End With
        End If
    End With
    Set range = Nothing
    Set wordEditor = Nothing
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub