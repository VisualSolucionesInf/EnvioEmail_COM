Imports System.IO
Imports System.Runtime.InteropServices

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.None)>
Public Class Acciones
    Implements IAcciones

    Private Const MSG_DEST_01 = "Atención: Uno de los destinatarios tiene un formato incorrecto de dirección email: {0}"
    Private Const MSG_DEST_02 = "Atención: Uno de los destinatarios está vacío o nulo."
    Private Const MSG_DEST_03 = "No hay destinatarios válidos para recibir el mensaje."

    Public Sub New()
        ' Necesario para COM.
    End Sub

    <ComVisible(True)>
    Public Function Enviar(hostSMTP As String, remitente As String, destinatarios As String, asunto As String, cuerpo As String, adjunto As String, user As String, pass As String) As Integer Implements IAcciones.Enviar
        Dim manejadorCorreo As New ManejadorEmail() With
        {
            .HostSmtp = hostSMTP,
            .Remitente = remitente
        }
        Dim enviadoConErrorDestinatario = False ' Para marcar que alguna dirección es errónea, pero no todas.
        Dim identificadorLOG As String

        If adjunto <> "" Then
            identificadorLOG = adjunto
            If identificadorLOG.LastIndexOf("\") >= 0 Then
                Try
                    identificadorLOG = identificadorLOG.Substring(identificadorLOG.LastIndexOf("\") + 1)
                Catch ex As Exception
                    identificadorLOG = asunto
                End Try
            End If
        Else
            identificadorLOG = asunto
        End If

        If identificadorLOG = "" Then identificadorLOG = Now.ToString("yy_MM_dd_HH_mm_ss")

        ' Destinatarios separados por ";".
        For Each destinatario As String In destinatarios.Split(";"c)
            Try
                manejadorCorreo.AddDestinatario(destinatario)
            Catch formatEx As FormatException
                Console.Error.WriteLine(String.Format(MSG_DEST_01, destinatario))
                EscribeLog(identificadorLOG, String.Format(MSG_DEST_01, destinatario))
                enviadoConErrorDestinatario = True
            Catch ex As Exception
                Console.Error.WriteLine(MSG_DEST_02)
                EscribeLog(identificadorLOG, MSG_DEST_02)
                enviadoConErrorDestinatario = True
            End Try
        Next

        If manejadorCorreo.Destinatarios.Count = 0 Then
            Console.Error.WriteLine(MSG_DEST_03)
            EscribeLog(identificadorLOG, MSG_DEST_03)
            Return 0
        End If

        manejadorCorreo.Asunto = asunto
        manejadorCorreo.CuerpoMensaje = IIf(cuerpo <> "", cuerpo, "Sin mensaje")
        If adjunto <> "" Then manejadorCorreo.AddAdjunto(adjunto)
        manejadorCorreo.Credenciales(user, pass)

        Dim mensajeError = ""
        Net.ServicePointManager.ServerCertificateValidationCallback = Function() True
        Net.ServicePointManager.SecurityProtocol = CType(3072, Net.SecurityProtocolType)
        Try
            manejadorCorreo.Enviar(mensajeError)
        Catch ex As Exception
            Console.Error.WriteLine(ex.Message)
            EscribeLog(identificadorLOG, "Error: " & ex.Message)
            Return 0
        End Try

        If mensajeError IsNot Nothing AndAlso mensajeError IsNot String.Empty Then
            Console.Error.WriteLine(mensajeError)
            EscribeLog(identificadorLOG, "Error: " & mensajeError)
            Return 0
        Else
            ' A pesar del envío correcto, si algún destinatario tenía un mal formato devolvemos un retorno "0".
            If enviadoConErrorDestinatario Then
                EscribeLog(identificadorLOG, "OK pero con error de formato en algún destinarario")
                Return 0
            Else
                ' Salida TODO OK "1".
                EscribeLog(identificadorLOG, "OK")
                Return 1
            End If
        End If
    End Function

    Private Shared Sub EscribeLog(identificador As String, mensaje As String)
        Dim path = My.Settings.directorio_logs

        If path(path.Length - 1) <> "\" Then path += "\"
        path += identificador & ".txt"

        Try
            Using fs = New FileStream(path, FileMode.Append)
                Using sw = New StreamWriter(fs)
                    sw.WriteLine($"{Now} - {mensaje}")
                End Using
            End Using
        Catch ex As Exception
            Console.Error.WriteLine("Error tratando el fichero de LOG: " & ex.Message)
        End Try
    End Sub
End Class
