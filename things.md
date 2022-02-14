    Sub enviarEmail()

        '0 Selecionar folha de trabalho
        Planilha1.Select

        '1 Declarar variaveis
        Dim nome, email, assunto, corpo As String

        Set obj_outlook = CreateObject("Outlook.application")
        Set novoEmail = obj_outlook.createitem(0)

        '2 Atribuir valores as variaveis
        nome = Range("c3").Value
        email = Range("d3").Value
        assunto = Range("g2").Value
        corpo = Range("g3").Value

        '3 Enviar Email

        With novoEmail
            .to = email
            .cc = email
            .Subject = assunto
            .body = corpo
            .display
            .send

        End With

    End Sub
