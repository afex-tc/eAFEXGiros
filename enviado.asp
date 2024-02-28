<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
  
  Dim objEMail
  Dim hora, fecha
  
  hora= replace(time,":","")
  fecha= replace(date,"/","")
  fecha= replace(fecha,"-","")
  
  sAsunto = "AFEX Ltda.   (CN" & fecha & hora &")"


  nombre   = request("nombre")
  apellido = request("apellido")
  ciudad   = request("ciudad")
  pais     = request("pais")
  email    = request("email")
  telefono = request("telefono")  
  celular  = request("telefonocelular")
  fax      = request("fax")
  opinion  = request("opinion")
  tipo     = request("tipo")
  consulta = request("consulta")
  yo       = request("yo")
  ip       = request.ServerVariables("REMOTE_ADDR")
  
   if ucase(pais) = "CHILE" then
	   contactos = "julio.greene@afex.cl,andres.aguilar@afex.cl,domingo.avila@afex.cl,hugo.sepulveda@afex.cl,arturo.munoz@afex.cl"
   else
	   contactos = "julio.greene@afex.cl,andres.aguilar@afex.cl," & _
	               "thomas.greene@afex.cl,domingo.avila@afex.cl,hugo.sepulveda@afex.cl,arturo.munoz@afex.cl"   
	end if  
	
	     
  sDescripcion = "			  Formulario de Contactenos		" & vbCrlf 
  sDescripcion = sDescripcion & vbCrlf & "Nombre    : " & trim(nombre)
  sDescripcion = sDescripcion & vbCrlf & "Apellido  : " & trim(apellido)
  sDescripcion = sDescripcion & vbCrlf & "Ciudad    : " & trim(ciudad)
  sDescripcion = sDescripcion & vbCrlf & "Pais      : " & trim(pais)
  sDescripcion = sDescripcion & vbCrlf & "Email     : " & trim(email)
  sDescripcion = sDescripcion & vbCrlf & "Telefono  : " & trim(telefono)
  sDescripcion = sDescripcion & vbCrlf & "Celular   : " & trim(celular)
  sDescripcion = sDescripcion & vbCrlf & "Fax       : " & trim(fax)
  sDescripcion = sDescripcion & vbCrlf & "Opinión   : " & trim(opinion)
  sDescripcion = sDescripcion & vbCrlf & "Tipo      : " & trim(tipo)
  sDescripcion = sDescripcion & vbCrlf & "Consulta  : " & trim(consulta)  
  sDescripcion = sDescripcion & vbCrlf & "IP Remoto : " & ip  
  sDescripcion = sDescripcion & vbCrlf & "Mi Correo Privado es : " & yo	

  sNombre = trim(nombre) & " " & trim(apellido) 

  EnviarEmail sNombre & " <" & trim(email) & ">", "laura.greene@afex.cl", contactos, sAsunto, sDescripcion, 0
  response.Redirect("respuesta.asp")
  

%>