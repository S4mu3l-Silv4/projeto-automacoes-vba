Sub EnviarEmailSolicitacaoAdicional()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim WordEditor As Object
    Dim rng As Range
    Dim saudacao As String
    Dim horaAtual As Integer
    Dim CorpoEmail As String
    Dim Assinatura As String
    On Error Resume Next
    Set rng = Sheets("Planilha1").Range("A1").CurrentRegion.Resize(, 11).SpecialCells(xlCellTypeVisible)
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
        .CC = ""
        .Subject = "Site(s) - Adicional por negociação"
        .Display
        Assinatura = .HTMLBody
        CorpoEmail = "<div style='font-family:Calibri;font-size:11pt;'>" & _
            saudacao & "<br><br>" & _
            "Xxx, solicito a sua autorização para podemos seguir com os valores abaixo: <br>" & _
        "</div><br>"
        .HTMLBody = CorpoEmail & Assinatura
        Set WordEditor = .GetInspector.WordEditor
        rng.Copy
        With WordEditor.Application.Selection
            .HomeKey 6
            .MoveDown Unit:=5, Count:=4
            .PasteExcelTable False, False, True
        End With
        If WordEditor.Tables.Count > 0 Then
            With WordEditor.Tables(WordEditor.Tables.Count)
                .AutoFitBehavior 1
            End With
        End If
    End With
    Set rng = Nothing
    Set WordEditor = Nothing
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub