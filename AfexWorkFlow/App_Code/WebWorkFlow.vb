Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports net.utilidades

<WebService(Namespace:="http://www.afex.cl/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class WebWorkFlow
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function enviarmail() As String

        Dim sendmailnet As New webcomponentes ' Clase Utilidades

        Dim swhtmlbody As System.IO.StringWriter
        Dim twTextWriter As HtmlTextWriter
        Dim recibir As String, aviso As String

        swhtmlbody = New System.IO.StringWriter
        twTextWriter = New HtmlTextWriter(swhtmlbody)

        aviso = "<br><br>Correo Automatico Generado desde Afex Portal<br>"
        aviso = aviso & "<br><br><br> Nombre: Prueba " '& nombres
        aviso = aviso & "Apellido : Prueba" '& apellidos

        twTextWriter.RenderBeginTag("html")
        twTextWriter.RenderBeginTag("head")
        twTextWriter.RenderBeginTag("title")
        twTextWriter.Write("Incidentes")
        twTextWriter.RenderEndTag()
        twTextWriter.RenderEndTag()
        twTextWriter.AddAttribute("bgcolor", "#ffffff")
        twTextWriter.RenderBeginTag("body")
        twTextWriter.AddAttribute("src", "http://www.afex.cl/Img/img_pagempresa/sucursal01_.jpg")
        twTextWriter.RenderBeginTag("img")
        twTextWriter.RenderEndTag()
        twTextWriter.AddAttribute("face", "Verdana, Arial, Helvetica, sans-serif")
        twTextWriter.AddAttribute("size", "2")
        twTextWriter.RenderBeginTag("font")
        twTextWriter.WriteLine(aviso)
        twTextWriter.RenderEndTag()
        twTextWriter.RenderEndTag()
        twTextWriter.RenderEndTag()

        recibir = sendmailnet.enviar_mail("oscar.pinto@afex.cl", "info@afex.cl",nothing, "Correo de Prueba", swhtmlbody.ToString)

        Return 1
    End Function

End Class
