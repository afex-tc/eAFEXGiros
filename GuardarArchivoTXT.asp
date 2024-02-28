<%@ Language=VBScript %>
<%
	Dim sInvoice
	Dim sNombreR
	Dim sApellidoR
	Dim sNombreB
	Dim sApellidoB
	Dim sDireccionB
	Dim sCiudadB
	Dim sPaisB
	Dim ddiPaisB
	Dim ddiAreaB
	Dim cFonoB
	Dim sMensaje
	Dim sNota
	Dim cMonto	
	Dim sArchivo, bOkCAT
	
	sArchivo = "c:\Sistemas\Desarrollo\AfexMoneyWeb\Archivos\prueba.txt"

	EscribirEnArchivo
	bOkCAT = CargarArchivoTXT(Session("afxCnxAFEXpress"), Session("afxCnxCorporativa"), sArchivo, "IB", "AFEXWEB")	
	'response.Redirect "http:compartido/informacion.asp?detalle=" & bOkCAT

	Function EscribirEnArchivo
	  Const ParaLeer = 1, ParaEscribir = 2
	  Dim fso, f
	  Set fso = CreateObject("Scripting.FileSystemObject")
	  Set f = fso.OpenTextFile(sArchivo, 2, True)
	  f.Write Request.Form("contenido")
	  f.Close
	End Function


	Public Function CargarArchivoTXT(ByVal Conexion, _
	                                 ByVal ConexionCP, _
	                                 ByVal Archivo, _
	                                 ByVal Agente, _
	                                 ByVal Usuario)
		Dim afxGiroCAT, fs, f, sLinea, i, aLinea1
		
	   CargarArchivoTXT = False
	   On Error Resume Next
		'Set afxGiroCAT = server.CreateObject("AFEXGiroXP.Giro")
		Set  fs = CreateObject("Scripting.FileSystemObject")
		Set f = fs.OpenTextFile(Archivo, 1, False)
		sLinea = f.readall				
		aLinea1 = Split(sLinea, vbCrLf)
	   For i = 0 To UBound(aLinea1)-1
	      sLinea = aLinea1(i)	      
	      SplitLinea ConexionCP, sLinea, Agente
	      response.Write i & ", " & sInvoice & ", " & cMonto & ", " & sNombreR & ", " & sApellidoR & ", " & sNombreB & ", " & sApellidoB & ", " & sDireccionB & ", " & sCiudadB & ", " & sPaisB & ", " & ddiPaisB & ", " & ddiAreaB & "<BR><BR>"
	      'afxGiroCAT.Enviar(Conexion, Agente, sAFEX, cMonto, 0, afxGiroNormal, afxPagoSucursal, afxPagoEfectivo, "USD", "USD", sMensaje, sNota, "", "", "", sNombreB & " " & sApellidoB, "", sDireccionB, sCiudadB, "", sPaisB, ddiPaisB, ddiAreaB, cFonoB, "", "", "", sNombreR, sApellidoR, "", "", "", "", 0, 0, 0, Usuario, , , sInvoice)
	   Next
	   CargarArchivoTXT = True
	   Set afxGiroCAT = Nothing
	   f.Close

	End Function

	Public Sub SplitLinea(ByVal Conexion, ByVal Linea, ByVal Agente)
	   Dim aLinea
	   Dim i
	   Dim sFono
	   Dim sNombrePais
	   Dim sNombreCiudad
		
		'response.Write Agente
	   Select Case Agente
	   Case "IB"         'Ibero
	      Linea = UCase(Linea)
	      Linea = LimpiarCampo(Linea)
	      Linea = Replace(Linea, Chr(34), "")
	      aLinea = Split(Linea, vbTab)
	      sInvoice = Replace(UCase(aLinea(3)), "AFEX", "")
	      cMonto = CCur(0 & Trim(aLinea(4)))
	      sNombreR = Trim(aLinea(7))
	      sApellidoR = Trim(aLinea(8))
	      sNombreB = Trim(aLinea(12))
	      sApellidoB = Trim(aLinea(13))
	      sDireccionB = Trim(aLinea(14))
	      sNombrePais = Trim(aLinea(17))
	      sFono = ObtenerFono(Trim(aLinea(18)))
	      sMensaje = Trim(aLinea(19))
	      sNota = Trim(aLinea(20))
	      sNombreCiudad = Trim(aLinea(23))
	      sCiudadB = ""
	      sPaisB = ""
	      ddiPaisB = 0
	      ddiAreaB = 0
	      ObtenerUbicacion Conexion, sCiudadB, sPaisB, sNombreCiudad, sNombrePais, ddiPaisB, ddiAreaB
	      If sPaisB = "CL" Then
	         If sCiudadB = "SCL" Then
	            cFonoB = CCur(0 & Mid(sFono, 4, 7))
	         Else
	            cFonoB = CCur(0 & Mid(sFono, 5, 6))
	         End If
	      Else
	         cFonoB = CCur(0 & sFono)
	      End If
	      'response.Write 1 & ", " & sNombreB
	   End Select
	End Sub

	Public Function ObtenerFono(ByVal Cadena)
	   Dim i
	   Dim sChar
	   Dim sFono

	   For i = 1 To Len(Cadena)
	      sChar = Mid(Cadena, i, 1)
	      If sChar >= "0" And sChar <= "9" Then sFono = sFono + sChar
	   Next
	   If Trim(sFono) = "" Then sFono = "0"
	   ObtenerFono = sFono
	End Function

	Public Function LimpiarCampo(ByVal Cadena)
	   Cadena = Replace(Cadena, Chr(10), "")
	   Cadena = Replace(Cadena, Chr(13), "")
	   Cadena = Replace(Cadena, Chr(126), "")
	   LimpiarCampo = Cadena
	End Function


	Public Function ObtenerNumero(ByVal Numero)
	'Objetivo   : Permite sacarle al número los caracteres extraños
	'Parametros : Numero  contiene el número a limpiar
	'Devuelve   : solo numeros
	Dim sNumero
	Dim sChar
	Dim i
	sNumero = ""
	For i = 1 To Len(Numero)
	   sChar = Mid(Numero, i, 1)
	   If (sChar >= "0" And sChar <= "9") Or sChar = "." Or sChar = "," Then
	      If sChar = "." Or sChar = "," Then sChar = "."
	      sNumero = sNumero + sChar
	   End If
	Next
	If sNumero = "" Then sNumero = "0"
	ObtenerNumero = sNumero
	End Function


	Public Function ObtenerUbicacion(ByVal Conexion, _
	                                 ByRef CodigoCiudad, _
	                                 ByRef CodigoPais, _
	                                 ByRef NombreCiudad, _
	                                 ByRef NombrePais, _
	                                 ByRef ddiPais, _
	                                 ByRef ddiArea)
	   Dim rsCiudad
	   Dim sSQL

	   'Screen.MousePointer = vbHourglass
	   ObtenerUbicacion = False
	   'Manejo de errores
	   On Error Resume Next

	   'Crea la consulta
	   sSQL = "SELECT    TOP 1 ci.*, pa.nombre as nombre_pais, pa.ddi as ddi_pais " & _
	          "FROM      Ciudad ci " & _
	          "JOIN      Pais Pa ON pa.codigo = ci.codigo_pais " & _
	          "WHERE     1 = 1 "
	   If NombreCiudad <> "" Then
	      sSQL = sSQL & " AND SOUNDEX(ci.nombre) = SOUNDEX('" & NombreCiudad & "') "
	   End If
	   If CodigoCiudad <> "" Then
	      sSQL = sSQL & " AND SOUNDEX(ci.codigo) = SOUNDEX('" & CodigoCiudad & "') "
	   End If

	   Set rsCiudad = EjecutarSQLCliente(Conexion, sSQL)

	   rsCiudad.MoveFirst
	   If Err.Number = 0 Then			
			CodigoCiudad = rsCiudad("codigo")
			NombreCiudad = rsCiudad("Nombre")
			CodigoPais = rsCiudad("codigo_pais")
			NombrePais = rsCiudad("nombre_pais")
			ddiAreaB = rsCiudad("ddi")
			ddiPaisB = rsCiudad("ddi_pais")
		   ObtenerUbicacion = True
		Else
			Err.Clear 
		End If
		
	End Function


%>
<!--#INCLUDE virtual="/compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" Content="VBScript">

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE=vbscript>
<!--
	
-->
</SCRIPT>


</HEAD>
<BODY>
Pruebas
<!--<P>Request.Form("contenido")</P>-->
<OBJECT RUNAT=server PROGID=Scripting.FileSystemObject id=OBJECT1> </OBJECT>
</BODY>
</HTML>
