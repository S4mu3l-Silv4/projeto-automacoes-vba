Sub EnviarEmailProgramacaoPagamentos()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim numSemana As Integer
    Dim Assinatura As String
    Dim CorpoEmail As String
    Dim CaminhoArquivo As String
    Dim saudacao As String
    Dim horaAtual As Integer
    CaminhoArquivo = "C:\Users\OPERAÇÃO 03\Dropbox\Subcon_CO\FSP - Form. Solicitação de Pagamentos - v2026_1_CO.xlsm"
    numSemana = DatePart("ww", Date, vbMonday, vbFirstFourDays)
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
        .To = "financeiro@premcell.com.br; kelly.martins@premcell.com.br"
        .Cc = "joao.moreira@premcell.com.br; luan.pereira@premcell.com.br"
        .Subject = "Programação de Pagamentos - CO - W" & numSemana
        .Attachments.Add CaminhoArquivo
        .Display
        Assinatura = .HTMLBody
        CorpoEmail = "<div style='font-family: Calibri; font-size: 11pt;'>" & _
            saudacao & "<br><br>" & _
            "Segue a Programação de Pagamentos: <br>" & _
        "</div>"
        .HTMLBody = CorpoEmail & Assinatura
    End With
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub
