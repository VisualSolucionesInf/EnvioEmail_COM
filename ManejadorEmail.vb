Imports System.Net
Imports System.Net.Mail

Public Class ManejadorEmail
	Private ReadOnly mensaje As MailMessage
	Private ReadOnly clienteSmtp As SmtpClient
	Private ReadOnly puertoSmtp As Integer = 587

	Public Sub New()
		mensaje = New MailMessage()
		clienteSmtp = New SmtpClient With {
			.UseDefaultCredentials = True,
			.Port = puertoSmtp, '25, 465, 587
			.EnableSsl = True,
			.DeliveryMethod = SmtpDeliveryMethod.Network
		}
	End Sub

	Public Property Remitente() As String
		Get
			Return mensaje.From.Address
		End Get
		Set(value As String)
			mensaje.From = New MailAddress(value)
			If clienteSmtp.Host = "" Then
				Dim partHost As String = value.Substring(value.IndexOf("@"))
				HostSmtpAuto(partHost)
			End If
		End Set
	End Property

	Private Sub HostSmtpAuto(auto As String)
		clienteSmtp.Host = "smtp." & auto
	End Sub

	Public ReadOnly Property Destinatarios() As List(Of String)
		Get
			Dim lista As New List(Of String)
			For Each destinatario As MailAddress In mensaje.To
				lista.Add(destinatario.Address)
			Next
			For Each destinatarioCopia As MailAddress In mensaje.CC
				lista.Add(destinatarioCopia.Address)
			Next
			Return lista
		End Get
	End Property

	Public Sub AddDestinatario(email As String)
		mensaje.To.Add(email)
	End Sub

	Public Sub AddDestinatarioCopia(email As String)
		mensaje.CC.Add(email)
	End Sub

	Public Property Asunto() As String
		Get
			Return mensaje.Subject
		End Get
		Set(value As String)
			mensaje.Subject = value
		End Set
	End Property

	Public Property CuerpoMensaje() As String
		Get
			Return mensaje.Body
		End Get
		Set(value As String)
			mensaje.Body = value
		End Set
	End Property

	Public ReadOnly Property Adjuntos() As AttachmentCollection
		Get
			Return mensaje.Attachments
		End Get
	End Property

	Public Sub AddAdjunto(fichero As String)
		Dim adjunto As New Attachment(fichero)
		mensaje.Attachments.Add(adjunto)
	End Sub

	Public Property HostSmtp() As String
		Get
			Return clienteSmtp.Host
		End Get
		Set(value As String)
			clienteSmtp.Host = value
		End Set
	End Property

	Public Sub Credenciales(usuario As String, password As String, Optional dominio As String = "")
		clienteSmtp.UseDefaultCredentials = False
		clienteSmtp.Credentials = New NetworkCredential(usuario, password, dominio)
	End Sub

	Public Property Ssl() As Boolean
		Get
			Return clienteSmtp.EnableSsl
		End Get
		Set(value As Boolean)
			clienteSmtp.EnableSsl = value
		End Set
	End Property

	Public Property Puerto() As Integer
		Get
			Return clienteSmtp.Port
		End Get
		Set(value As Integer)
			clienteSmtp.Port = value
		End Set
	End Property

	Public Property MetodoEnvio() As SmtpDeliveryMethod
		Get
			Return clienteSmtp.DeliveryMethod
		End Get
		Set(value As SmtpDeliveryMethod)
			clienteSmtp.DeliveryMethod = value
		End Set
	End Property

	Public Function Enviar(Optional ByRef mensajeError = "") As Boolean
		Try
			clienteSmtp.Send(mensaje)
		Catch excepcion As Exception
			mensajeError = excepcion.Message
			If excepcion.InnerException IsNot Nothing Then
				mensajeError = mensajeError & " : " & excepcion.InnerException.Message
			End If
			Return False
		End Try
		Return True
	End Function
End Class
