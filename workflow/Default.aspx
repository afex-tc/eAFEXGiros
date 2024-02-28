<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>


<% 
    
    Dim clsweb As New TareaWorkFlow.WSIniciaProcedimientoService
    
    Dim valor()
    Dim param
    Dim hora, fecha
    Dim sAsunto, nombre, apellido, ciudad, pais, email, telefono, celular, fax
    Dim opinion, tipo, consulta, yo, ip, sdescripcion
  
    hora = "100000" 'Replace(time, ":", "")
    fecha = "29122006" 'Replace(date, "/", "")
    fecha = "29122006" 'Replace(fecha, "-", "")
  
    sAsunto = "AFEX Ltda.   (CN" & fecha & hora & ")"

    nombre = Request("nombre")
    apellido = Request("apellido")
    ciudad = Request("ciudad")
    pais = Request("pais")
    email = Request("email")
    telefono = Request("telefono")
    celular = Request("telefonocelular")
    fax = Request("fax")
    opinion = Request("opinion")
    tipo = Request("tipo")
    consulta = Request("consulta")
    yo = Request("yo")
    ip = Request.ServerVariables("REMOTE_ADDR")

    sdescripcion = "			  Formulario de Contactenos		" & vbCrLf
    sdescripcion = sdescripcion & "Nombre    : " & Trim(nombre) & vbCrLf
    sdescripcion = sdescripcion & "Apellido  : " & Trim(apellido) & vbCrLf
    sdescripcion = sdescripcion & "Ciudad    : " & Trim(ciudad) & vbCrLf
    sdescripcion = sdescripcion & "Pais      : " & Trim(pais) & vbCrLf
    sdescripcion = sdescripcion & "Email     : " & Trim(email) & vbCrLf
    sdescripcion = sdescripcion & "Telefono  : " & Trim(telefono) & vbCrLf
    sdescripcion = sdescripcion & "Celular   : " & Trim(celular) & vbCrLf
    sdescripcion = sdescripcion & "Fax       : " & Trim(fax) & vbCrLf
    sdescripcion = sdescripcion & "Opinión   : " & Trim(opinion) & vbCrLf
    sdescripcion = sdescripcion & "Tipo      : " & Trim(tipo) & vbCrLf
    sdescripcion = sdescripcion & "Consulta  : " & Trim(consulta) & vbCrLf
    sdescripcion = sdescripcion & "IP Remoto : " & ip & vbCrLf
    sdescripcion = sdescripcion & "Mi Correo Privado es : " & yo
    
    param = "servidor::192.168.111.140//de::" & email & "//para::domingo.avila@afex.cl" & "//referencia::" & sAsunto & "//mensaje::" & sdescripcion
    
    Try
        Dim datos() As String = {3, "192.168.111.141", "eworkflow", "eworkflow", "eworkflow", "MAIL", "1", "invitado", "1", "invitado", "1", "Tarea Web", "E:/AfexWorkflow/Apache Software Foundation/Tomcat 5.0/webapps/eworkflow", "http://roble:8080/eworkflow", param}
        
        valor = clsweb.iniciaProcedimiento(datos)
        'Response.Redirect("respuesta.asp") '("Tarea generada en forma correcta :" & valor(2))        
    Catch exp As System.Security.SecurityException
    End Try
        
    'args[0]   = "3";             // Tipo de base de datos
    'args[1]   = "localhost";     // Host de base de datos
    'args[2]   = "xnear10x";      // Nombre de base de datos
    'args[3]   = "sa";            // Usuario de base de datos
    'args[4]   = "sa";            // Password de base de datos
    'args[5]   = "Inquietudes";   // Nombre Procedimiento
    'args[6]   = "1";             // Empresa Procedimiento
    'args[7]   = "fgomez";        // Rol creador
    'args[8]   = "1";             // Empresa rol creador
    'args[9]   = "mmendez";       // Rol destino
    'args[10]  = "1";             // Empresa rol destino
    'args[11]  = "Creacion de tarea via WS";     // Asunto    
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Documento sin t&iacute;tulo</title>
<style type="text/css">
<!--
.Estilo3 {font-size: 14px}
-->
</style>
</head>

    <script language="vbscript">
    <!--
    
        window.dialogWidth = 26
	    window.dialogHeight = 16
	    window.dialogLeft = 240
	    window.dialogTop = 220
	    window.defaultstatus = ""
    
        sub window_onload()            
            window.returnvalue = <%=valor(2)%>
            window.close()           
        end sub
    
    -->
    </script>

<body>
</body>
</html>