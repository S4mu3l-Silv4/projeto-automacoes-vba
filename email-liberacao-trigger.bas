Sub EnviarEmailLiberacaoTrigger()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim WordEditor As Object
    Dim rng As Range
    Dim saudacao As String
    Dim horaAtual As Integer
    On Error Resume Next
    Set rng = Sheets("VOs_CO").Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
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
    Set OutlookApp = GetObject(Class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    Set OutlookMail = OutlookApp.CreateItem(0)
    With OutlookMail
        .To = "exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com"
        .CC = "exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com"
        .Subject = "Liberação de trigger - CO"
        .Display
        Set WordEditor = .GetInspector.WordEditor
        With WordEditor.Application.Selection
            .TypeText saudacao
            .TypeParagraph
            .TypeParagraph
            .TypeText "Poderiam verificar as VOs aprovadas abaixo por gentileza? Esses casos estão com pendência de aprovação: "
            .TypeParagraph
            .TypeParagraph
            .TypeText "Informo que todos os requisitos de aplicação estão atendidos: "
            .TypeParagraph
            .TypeParagraph
        End With
        rng.Copy
        WordEditor.Application.Selection.PasteExcelTable False, False, True
    End With
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub