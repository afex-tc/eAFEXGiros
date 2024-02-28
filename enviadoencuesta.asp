<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
  
  Dim objEMail
  Dim hora, fecha
  
  hora= replace(time,":","")
  fecha= replace(date,"/","")
  fecha= replace(fecha,"-","")
  
  sAsunto = "Encuesta AFEX Ltda.   (CN" & fecha & hora &")"


  nombre   = request("nombre")
  apellido = request("apellido")
  ciudad   = request("ciudad")
  pais     = request("pais")
  email    = request("email")
  telefono = request("telefono")  
  celular  = request("telefonocelular")
  empresa  = request("empresa")
  oficina  = request("oficina")
  pregunta2   = request("pregunta2")
  pregunta3   = request("pregunta3")
  pregunta4   = request("pregunta4")
  pregunta5   = request("pregunta5")
  pregunta6   = request("pregunta6")
  pregunta7   = request("pregunta7")
  rapidez         = request("rapidez")
  confiabilidad   = request("confiabilidad")
  preciocalidad   = request("preciocalidad")
  atencion        = request("atencion")
  pregunta9       = request("pregunta9")
  pregunta10      = request("pregunta10")
  comentario      = request("comentario")
  
  yo       = request("yo")
  ip       = request.ServerVariables("REMOTE_ADDR")
  
	' JFMG 05-01-2010 se comenta para asignar la lista desde la BD
	'if ucase(pais) = "CHILE" then
	'   contactos = "julio.greene@afex.cl,andres.aguilar@afex.cl,domingo.avila@afex.cl,hugo.sepulveda@afex.cl,arturo.munoz@afex.cl"
	'else
    '  contactos = "andres.aguilar@afex.cl," & _
    '              "thomas.greene@afex.cl,domingo.avila@afex.cl,hugo.sepulveda@afex.cl" 
	' 
 	'end if  
 	dim sPara, sCopia, sMensajeErrorCorreo, i
	sPara = ObtenerCorreoElectronicoContacto(2)
	if sPara = "" then
		sMensajeErrorCorreo = "No se encontró la lista de correos."
	else
		i = instr(sPara, "//")
		sCopia = mid(sPara, i + 2)
		sPara = left(sPara, i - 1)		
	end if
	' ********** FIN JFMG 05-01-2010 **********************************	
 	
		     
  sDescripcion = "	   Formulario de Encuesta AFEX		" & vbCrlf 
  sDescripcion = sDescripcion & vbCrlf & "Nombre    : " & trim(nombre)
  sDescripcion = sDescripcion & vbCrlf & "Apellido  : " & trim(apellido)
  sDescripcion = sDescripcion & vbCrlf & "Ciudad    : " & trim(ciudad)
  sDescripcion = sDescripcion & vbCrlf & "Pais      : " & trim(pais)
  sDescripcion = sDescripcion & vbCrlf & "Email     : " & trim(email)
  sDescripcion = sDescripcion & vbCrlf & "Telefono  : " & trim(telefono)
  sDescripcion = sDescripcion & vbCrlf & "Celular   : " & trim(celular)
  sDescripcion = sDescripcion & vbCrlf & "Oficina   : " & trim(oficina)
  sDescripcion = sDescripcion & vbCrlf & "¿Cuánto tiempo lleva utilizando los servicios Afex?  : " & trim(pregunta2)  
  sDescripcion = sDescripcion & vbCrlf & "¿Qué operaciones hace Ud. usualmente en Afex?        : " & trim(pregunta3)
  sDescripcion = sDescripcion & vbCrlf & "¿Con qué frecuencia usa Ud. los servicios Afex?      : " & trim(pregunta4)
  sDescripcion = sDescripcion & vbCrlf & "¿Cómo llegó a operar a través de Afex?               : " & trim(pregunta5)
  sDescripcion = sDescripcion & vbCrlf & "¿Cuan satisfecho está Ud. con los servicios otorgados por Afex?     : " & trim(pregunta6)
  sDescripcion = sDescripcion & vbCrlf & "¿Cómo evaluaría Ud. la atención por parte de los cajeros de Afex?   : " & trim(pregunta7)
  sDescripcion = sDescripcion & vbCrlf & "Nota - Rapidez   : " & trim(rapidez)
  sDescripcion = sDescripcion & vbCrlf & "Nota - Confiabilidad   : " & trim(Confiabilidad)
  sDescripcion = sDescripcion & vbCrlf & "Nota - Precio-Calidad  : " & trim(preciocalidad)
  sDescripcion = sDescripcion & vbCrlf & "Nota - Atención        : " & trim(atencion)
  sDescripcion = sDescripcion & vbCrlf & "¿Recomendaría Ud. los servicios de Afex a otra persona? : " & trim(pregunta9)  
  sDescripcion = sDescripcion & vbCrlf & "¿Buscaría Ud. empresas similares para realizar sus operaciones? : " & trim(pregunta10)
  sDescripcion = sDescripcion & vbCrlf & "Comentario  : " & trim(comentario)  

  sDescripcion = sDescripcion & vbCrlf & "IP Remoto : " & ip  
  sDescripcion = sDescripcion & vbCrlf & "Mi Correo Privado es : " & yo	

  sNombre = trim(nombre) & " " & trim(apellido) 

  ' JFMG 06-01-2010 se agrega lista desde la BD
  'EnviarEmail sNombre & " <" & trim(email) & ">", "laura.greene@afex.cl", contactos, sAsunto, sDescripcion, 0
  EnviarEmail sNombre & " <" & trim(email) & ">", sPara, sCopia, sAsunto, sDescripcion, 0
  ' ********** FIN JFMG 06-01-2010 ***********************
  response.Redirect("respuesta.asp")
  
%>