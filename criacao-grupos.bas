Sub CriarGrupos()
    Dim comando As String
    Dim nomeGrupo As String
    Dim contatos(1 To 3) As String
    Dim i As Integer, g As Integer, c As Integer
    Dim celulas As Variant
    celulas = Array("B1")
    contatos(1) = "exemplo - Premcell"
    contatos(2) = "exemplo - Premcell"
    contatos(3) = "exemplo - Premcell"
    comando = "cmd.exe /c start whatsapp:"
    Shell comando, vbHide
    Application.Wait (Now + TimeValue("00:00:08"))
    For g = LBound(celulas) To UBound(celulas)
        nomeGrupo = Range(celulas(g)).Value
        If nomeGrupo <> "" Then
            SendKeys "^+N", True
            Application.Wait (Now + TimeValue("00:00:03"))
            For c = 1 To 3
                SendKeys contatos(c), True
                Application.Wait (Now + TimeValue("00:00:03"))
                SendKeys "{ENTER}", True
                Application.Wait (Now + TimeValue("00:00:01"))
            Next c
            SendKeys "{ENTER}", True
            Application.Wait (Now + TimeValue("00:00:03"))
            Application.Wait (Now + TimeValue("00:00:02"))
            SendKeys nomeGrupo, True
            Application.Wait (Now + TimeValue("00:00:03"))
            SendKeys "{ENTER}", True
            Application.Wait (Now + TimeValue("00:00:05"))
        End If
    Next g
End Sub