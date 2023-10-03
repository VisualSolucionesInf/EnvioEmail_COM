Public Interface IAcciones
    Function Enviar(hostSMTP As String, remitente As String, destinatarios As String, asunto As String, cuerpo As String, adjunto As String, user As String, pass As String) As Integer
End Interface
