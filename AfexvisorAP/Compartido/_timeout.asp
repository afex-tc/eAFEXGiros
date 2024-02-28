<!-- _timeout.asp -->
<%
	Response.Expires = 0
	Response.buffer = true	
	Response.expires = 0
	Response.expiresabsolute = Now() - 1
	Response.addHeader "pragma", "no-cache"
	Response.addHeader "cache-control", "private"
	Response.CacheControl = "no-cache"
	If Not Session("SesionActiva") Then
	%>
		<HTML>
		<head>
		<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<title>Error de Sesion</title>
		</head>
		<BODY style="margin: 0 0 0 10; font-family: Verdana; font-size: 14pt">
			No se registró ninguna sesión activa
		</BODY>
		</HTML>
	<%
		response.end
	End If
	
%>