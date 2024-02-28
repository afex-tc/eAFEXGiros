<%@ Language=VBScript %>
<% Response.Buffer = True %>
<!--#INCLUDE VIRTUAL="/Compartido/TimeOut.asp" -->
<%

	Server.ScriptTimeOut = 1500
	
	' Jonathan mIranda G. 11-01-2007
	private Function EjecutarSQLClienteLOCAL(ByVal Conexion, ByVal SQL)
		Dim rs
		Dim Cnn		
		Const adUseClient = 2
		Const adOpenStatic = 3
		Const adLockBatchOptimistic = 4
		
		'On Error Resume Next

	'	Set EjecutarSQLClienteLOCAL = Nothing
	   
		Set Cnn = server.CreateObject("ADODB.Connection")
		Cnn.CommandTimeout = 600
		Cnn.Open Conexion
	   
		If Err.number <> 0 Then
			Cnn.Close
			Set Cnn = Nothing			
			MostrarErrorMS "Ejecutar SQL 1"
		End If
				
		Set rs = server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3
		rs.Open SQL, Cnn, 3, 4		

		If Err.number <> 0 Then
			'rsESQL.Close
			Set rs = Nothing
			err.raise 50000,"EjecutarSQLClienteLOCAL", err.description
			exit function
		End If		

		Set rs.ActiveConnection = Nothing			
		
		if not rs.eof then
			EjecutarSQLClienteLOCAL = rs("codigogiro")
		end if

		rs.Close
		Set rs = Nothing
		
		Cnn.Close
		Set Cnn = Nothing

	End Function	
'-------------------------- Fin -------------------------


	Const afxIbero = "IB"
	Const afxMultiexpress = "ND"
	Const afxRapidEnvios = "TI"
	Const afxAfexLax = "AL"
	' Jonathan Miranda G. 18-12-2006
	Const afxDelgado = "DE"
	const afxRemittance="UR"
	const afxOmnex = "OG"
	const afxUnigiros ="UE"
	const afxLatinoenvios = "LT"
	const afxCsi = "CI"
	const afxLatinExpress = "LE"
	const afxVigo = "VG"
	'---------------------- Fin ----------------------------	

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
	Dim cMonto, sFecha, nRegistrosTransmitidos
	Dim sFono
	Dim sNombrePais
	Dim sNombreCiudad
	Dim sArchivo, bOkCAT, sNombreArchivo, nGiros, sArchivo1, sLog
	Dim Giros, GirosOk, GirosError, ListaGirosOk, ListaGirosError, sHtml
	' Jonathan MIranda G. 11-01-2007
	Dim cMontoPesos, cTipoCambio, sMoneda, sMonedaLocal
	'------------------ Fin ---------------------


	Giros = 0
	GirosOk = 0
	GirosError = 0	
	
	sArchivo = Replace(cSTR(date), "/", "")
	sArchivo = Replace(sArchivo, "-", "")
	sArchivo = sArchivo & Replace(Time, ":", "")
	sArchivo1 = Session("CodigoAgente") & sArchivo
	sNombreArchivo = Session("CodigoAgente") & sArchivo & ".txt"
	'*************carga archivos desde equipo Patty************************************************************
	'sLog = "C:\CARPETA\"  & Session("CodigoAgente") & "\" & sArchivo & ".log"
	'sArchivo = "C:\RESPALDO\" & Session("CodigoAgente") & "\" & sNombreArchivo	
	'*********************************************************************************************************
	sLog = "d:\sitios\Archivos\" & Session("CodigoAgente") & "\" & sArchivo & ".log"
	sArchivo = "d:\sitios\Archivos\" & Session("CodigoAgente") & "\" & sNombreArchivo	

	EscribirEnArchivo
	
	' Jonathan Miranda G. 12-12-2006	
	'if Session("CodigoAgente") = afxDelgado or Session("CodigoAgente") = afxRemittance or _
	'	Session("CodigoAgente") = afxOmnex or _ 
	'	Session("CodigoAgente")= afxUnigiros  or Session("CodigoAgente") = afxLatinoenvios or _
	'	Session("CodigoAgente") = afxRapidEnvios or Session("CodigoAgente") = afxCsi or  _
	'	Session("CodigoAgente") = afxLatinExpress or afx then
			'response.Write sArchivo & ", " & Session("CodigoAgente") & ", " & Date & ", I"
			'response.End 	
	'		CargarArchivosDLL sArchivo, Session("CodigoAgente"), Date, "I", Session("NombreUsuario")	
	'	
	'else
'			bOkCAT = CargarArchivoTXT(Session("afxCnxAFEXpress"), Session("afxCnxCorporativa"), sArchivo, Session("CodigoAgente"), Session("NombreUsuario"))
'	end if
	'--------------------- Fin ------------------------------COMENTADO PSS 11-02-2009
	
	
	' PSS 11-02-2009  NUEVO 
	
	If Session("CodigoAgente")= afxIbero or Session("CodigoAgente") = afxMultiexpress or _
	 Session("codigoAgente") = afxAfexLax then
		bOkCAT = CargarArchivoTXT(Session("afxCnxAFEXpress"), Session("afxCnxCorporativa"), sArchivo, Session("CodigoAgente"), Session("NombreUsuario"))
	Else 
		CargarArchivosDLL sArchivo, Session("CodigoAgente"), Date, "I", Session("NombreUsuario")	
	End IF
	
	sCorreo = vbCrLf & "Resumen de Carga de Giros" & vbCrLf & vbCrLf
	sCorreo = sCorreo & "Nombre del archivo: " & Request.Form("fileadjunto") & vbCrLf
	sCorreo = sCorreo & "Código confirmación AFEX: " & sArchivo1 & vbCrLf
	sCorreo = sCorreo & "Total de Giros: " & Giros & vbCrLf
	sCorreo = sCorreo & "Giros cargados exitosamente: " & GirosOk & vbCrLf
	If GirosError > 0 Then
		sCorreo = sCorreo & "Giros no cargados por error: " & GirosError & vbCrLf
		sCorreo = sCorreo & ListaGirosError & vbCrLf & vbCrLf
	End If
	'--------------------------------------------------------------
	
	
	
	' JFMG 07-08-2008 Omitido para que no se envíe un mail el cual puede ser tomado como espam
	'EnviarCorreo sCorreo
	' *************************** FIN *************************


	sHtml = Replace(sCorreo, vbCrlf, "<br>")
	sHtml = Replace(sHtml, "          ", "- - - - - > ")
	'Response.Redirect "http:../compartido/informacion.asp?detalle=" & sHtml
	
'	If Not bOkCat Then
'		EnviarCorreo "No se pudo cargar el archivo " & Request.Form("fileadjunto") & " (" & sArchivo1 & "), o parte de el, debido a un error desconocido."
'		response.Redirect "http:../compartido/informacion.asp?detalle=No se pudo cargar el archivo, o parte de el, debido a un error desconocido"	
'	End If
'	EnviarCorreo "Se han cargado " & nGiros & " giros desde el archivo " & Request.Form("fileadjunto") & " (" & sArchivo1 & ")."
'	response.Redirect "http:../compartido/informacion.asp?titulo=Carga de Archivos&detalle=El archivo se cargó con éxito"	
	

	Function EscribirEnArchivo
	  Const ParaLeer = 1, ParaEscribir = 2
	  Dim fso, f
	  Set fso = CreateObject("Scripting.FileSystemObject")
	  Set f = fso.OpenTextFile(sArchivo, 2, True)
	  f.Write Request.Form("contenido")
	  f.Close
	  
	End Function

	Function EscribirEnLog
	  Const ParaLeer = 1, ParaEscribir = 2
	  Dim fso, f
	  Set fso = CreateObject("Scripting.FileSystemObject")
	  Set f = fso.OpenTextFile(sLog, 2, True)
	  f.Write "Se produjeron los siguientes errores en la carga de giros: "
	  f.Write ListaGirosError
	  f.Close
	End Function


	Function CargarArchivoTXT(ByVal Conexion, _
	                          ByVal ConexionCP, _
	                          ByVal Archivo, _
	                          ByVal Agente, _
	                          ByVal Usuario)
		Dim afxGiroCAT, fs, f, sLinea, i, aLinea1, sGiro, nGirosOk, nGirosError, nInicio
		dim rsGiro
		
	   CargarArchivoTXT = False
	   On Error Resume Next
		' Jonathan Miranda G. 09-01-2007
		'Set afxGiroCAT = server.CreateObject("AFEXGiro.Giro")
		'-------------------------- Fin ------------------------
		Set  fs = CreateObject("Scripting.FileSystemObject")
		Set f = fs.OpenTextFile(Archivo, 1, False)
		sLinea = f.readall				
		aLinea1 = Split(sLinea, vbCrLf)
		
		Select Case Agente
			Case afxMultiexpress
				sFecha = Mid(aLinea1(0), 4, 2) & "-" & Mid(aLinea1(0), 1, 2) & "-" & Mid(aLinea1(0), 7, 2)
				nRegistrosTransmitidos = cCur(0 & ObtenerNumero(Mid(aLinea1(0), 23, 5)))
				nInicio = 1
					
			'Case afxRapidEnvios
			'	nInicio = 1
						
			Case afxAfexLax
				' Jonathan Miranda G. 11-12-2006 Lineas comentadas por cambio de formato.
				sFecha = Mid(aLinea1(0), 8, 2) & "-" & Mid(aLinea1(0), 6, 2) & "-" & Mid(aLinea1(0), 2, 4)
				nRegistrosTransmitidos = cCur(0 & ObtenerNumero(Mid(aLinea1(0), 27, 4)))
				nInicio = 1
				'nInicio = UBound(aLinea1)
				'------------------------ Fin ---------------------------------				
				
			Case Else
				nInicio = 0
				
		End Select
		
	   For i = nInicio To UBound(aLinea1)
			' Jonathan Miranda G. 09-01-2007
			sGiro = empty
			cMontoPesos = "null"
			cTipoCambio = "null"
			sMoneda = empty
			sMonedaLocal = 0
			'------------------- Fin ---------------------

	      sLinea = aLinea1(i)
	      If Trim(sLinea) = "" Then Exit For	      
	      SplitLinea ConexionCP, sLinea, Agente
	      'response.Write i & ", " & sInvoice & ", " & cMonto & ", " & sNombreR & ", " & sApellidoR & ", " & sNombreB & ", " & sApellidoB & ", " & sDireccionB & ", " & sCiudadB & ", " & sPaisB & ", " & ddiPaisB & ", " & ddiAreaB & "<BR><BR>"
	      'response.Redirect "http:../compartido/informacion.asp?detalle=" & i & ", " & sInvoice & ", " & cMonto & ", " & sNombreR & ", " & sApellidoR & ", " & sNombreB & ", " & sApellidoB & ", " & sDireccionB & ", " & sCiudadB & ", " & sPaisB & ", " & ddiPaisB & ", " & ddiAreaB & "<BR><BR>"
			'					 0, 305391, 526, CRISTIAN, DI SILVESTRO, BEATRIZ ZULMA, AROMA, 11B ANTARTIDA MG C7, MDZ, AR, 54, 261
		   'response.Redirect "http:../compartido/informacion.asp?detalle=" & Conexion & ", " &  Agente & ", " &  Session("CodigoMatriz") & ", " &  cMonto & ", " &  0 & ", " &  afxGiroNormal & ", " &  afxPagoSucursal & ", " &  afxPagoEfectivo & ", " &  "USD" & ", " &  "USD" & ", " &  sMensaje & ", " &  sNota & ", " &  "" & ", " &  "" & ", " &  "" & ", " &  sNombreB & " " & sApellidoB & ", " &  "" & ", " &  sDireccionB & ", " &  sCiudadB & ", " &  "" & ", " &  sPaisB & ", " &  ddiPaisB & ", " &  ddiAreaB & ", " &  cFonoB & ", " &  "" & ", " &  "" & ", " &  "" & ", " &  sNombreR & ", " &  sApellidoR & ", " &  "" & ", " &  "" & ", " &  "" & ", " &  "" & ", " &  0 & ", " &  0 & ", " &  0 & ", " &  Usuario & ", " & "" & ", " & "" & ", " &  sInvoice & ", " & "" & ", " & "" & ", " & "" & ", " & "" & ", " & "" & ", " &  "False" & ", " &  0 & ", " &  0 & ", " &  0 & ", " &  0 & ", " &  0 & ", " &  0 & ", " &  Session("Categoria")
		   'response.Redirect "http:../compartido/informacion.asp?detalle=" & _
		   'Conexion & ", " & Agente & ", " & Session("CodigoMatriz") & ", " & cMonto & ", 0, " & afxGiroNormal & ", " & afxPagoSucursal & ", " & afxPagoEfectivo & ", " & "USD" & ", " & "USD" & ", " & sMensaje & ", " & sNota & ", " & "" & ", " & "" & ", " & "" & ", " & sNombreB & ", " & sApellidoB & ", " & Left(sDireccionB, 255) & ", " & sCiudadB & ", " & "" & ", " & sPaisB & ", " & ddiPaisB & ", " & ddiAreaB & ", " & cCur(cFonoB) & ", " & "" & ", " & "" & ", " & "" & ", " & sNombreR & ", " & sApellidoR & ", " & "" & ", " & "" & ", " & "" & ", " & "" & ", 0, 0, 0, " & Usuario & ", , , " & sInvoice & ", , , , , " & True & ", " & False & ", 0, 0, 0, 0, 0, 0, " & Session("Categoria")
		
					   
			' Jonathan Miranda G. 09-01-2007
'	      sGiro = afxGiroCAT.Enviar(Conexion, Agente, Session("CodigoMatriz"), cMonto, 0, afxGiroNormal, 'afxPagoSucursal, afxPagoEfectivo, "USD", "USD", sMensaje, sNota, "", "", "", sNombreB, sApellidoB, 'Left(sDireccionB, 255), sCiudadB, "", sPaisB, ddiPaisB, ddiAreaB, cCur(cFonoB), "", "", "", sNombreR, 'sApellidoR, "", "", "", "", 0, 0, 0, Usuario, , ,  sInvoice, , , , , True, False, 0, 0, 0, 0, 0, 0, 'Session("Categoria"))

		  
			sSQL = " execute enviargiro " & _
			EvaluarSTR(Agente) & ", " & EvaluarSTR(Session("CodigoMatriz")) & ", " & _
			FormatoNumeroSQL(cCur(cDbl(cMonto))) & ", 0, 0, 1, 0, 'USD', 'USD', " & _
			EvaluarSTR(sMensaje) & ", " & EvaluarSTR(sNota) & ", " & " NULL, NULL, NULL, " & _
			EvaluarSTR(sNombreB) & ", " & EvaluarSTR(sApellidoB) & ", " & _
			EvaluarSTR(Left(sDireccionB, 255)) & ", " & EvaluarSTR(sCiudadB) & ", " & " NULL, " & _
			EvaluarSTR(sPaisB) & ", " & cInt(0 & ddiPaisB) & ", " & cInt(0 & ddiAreaB) & ", " & _
			cCur(0 & cCur(cFonoB)) & ", null, null, null, " & EvaluarSTR(Trim(sNombreR)) & ", " & _
			EvaluarSTR(sApellidoR) & ", null , null, null, null, 0, 0, 0, " & _
			EvaluarSTR(Session("NombreUsuarioOperador")) & ", null, null, " & EvaluarSTR(sInvoice) & ", " & _
			"null, null, null, null, null, null, null, null, null, null, null, null," & sMonedaLocal & ", " & _
			cTipoCambio & ", " & cMontoPesos

'Response.Write ssql
'Response.End 
			'set rsGiro = EjecutarSqlClienteLOCAL(Session("afxCnxAFEXpress"), sSQL)
			sGiro = EjecutarSqlClienteLOCAL(Session("afxCnxAFEXpress"), sSQL)
		  '---------------------- Fin --------------------------
	      'limpia variables			
				cMonto = ""
				sNombreR = ""
				sApellidoR = ""
				sNombreB = ""
				sApellidoB = ""
				sDireccionB = ""
				sNombrePais = ""
				sFono = ""
				sMensaje = ""
				sNota = ""
				sNombreCiudad = ""
				sNombreCiudad = ""
				sCiudadB = ""
				sPaisB = ""
				ddiPaisB = 0
				ddiAreaB = 0
				' Jonathan Miranda G. 11-01-2007
				cMontoPesos = 0
				cTipoCambio = 0
				sMoneda = empty
				sMonedaLocal = 0
				'--------------- Fin ---------------------
			'************
			If Err.number <> 0 Then
				'Set afxGiroCAT = Nothing
				'MostrarErrorMS "Grabar Envio Giro 1"
				GirosError = GirosError + 1
				ListaGirosError = ListaGirosError & "          " & sInvoice & " " & Err.Description & vbCrLf

			' Jonathan Miranda G. 09-01-2007
			'ElseIf afxGiroCAT.ErrNumber <> 0 Then
				'MostrarErrorAFEX afxGiroCAT, "Grabar Envio Giro 2"
			'	GirosError = GirosError + 1
			'	If afxGiroCAT.ErrNumber = 10263 Then
			'		ListaGirosError = ListaGirosError &  "          " & sInvoice & " El Invoice ya existe" & 'vbCrLf
			'	Else
			'		ListaGirosError = ListaGirosError & "          " & sInvoice & " " & 'afxGiroCAT.ErrDescription & vbCrLf
			'	End If
			'----------------------------- Fin ----------------------
			ElseIf sgiro = "" Then
				'MostrarErrorMS "Grabar Envio Giro 3"
				GirosError = GirosError + 1
				ListaGirosError = ListaGirosError  & "          " &  sInvoice & " Error desconocido" & vbCrlf
				
			Else
				'sGiro = rsGiro("codigogiro")
				GirosOk = GirosOk + 1
				ListaGirosOk = ListaGirosOk & sInvoice & vbCrlf			
			End If			
			Err.Clear 
	   Next
	   Giros = i - nInicio
	   CargarArchivoTXT = True
	   
		' Jonathan Miranda G.
	   'Set afxGiroCAT = Nothing
	   set rsGiro = nothing
		'--------------------- Fin -------------------------------
	   f.Close

	End Function

	Sub SplitLinea(ByVal Conexion, ByVal Linea, ByVal Agente)
		Dim aLinea
		
		'response.Write Agente
		Select Case Agente
			Case afxIbero         'Ibero
				Linea = UCase(Linea)
				Linea = LimpiarCampo(Linea)
				Linea = Replace(Linea, Chr(126), "")
				Linea = Replace(Linea, Chr(34), "")
				aLinea = Split(Linea, vbTab)
				sInvoice = Replace(UCase(aLinea(3)), "AFEX", "")
				cMonto = CCur(0 & Trim(aLinea(4)))
				' Jonathan Miranda G. 11-01-2007
				sMoneda = Trim(aLinea(2))
				cMontoPesos = CCur(0 & Trim(aLinea(6)))
				cTipoCambio = CCur(0 & Trim(aLinea(5)))
				if instr(ucase(sMoneda), "PESOS CHILENOS") > 0 then
					sMonedaLocal = 1
					cMontoPesos = round(replace(ccur(cMontoPesos), ",", "."), 0)
					cTipoCambio = replace(ccur(cTipoCambio), ",", ".")
				else
					sMonedaLocal = 0
					cMontoPesos = "null"
					cTipoCambio = "null"
				end if				
				'------------------- Fin -------------------------

				sNombreR = Trim(aLinea(7))
				sApellidoR = Trim(aLinea(8))
				sNombreB = Trim(aLinea(12))
				sApellidoB = Trim(aLinea(13))
				sDireccionB = Trim(aLinea(14)) & ", " & Trim(aLinea(23))
				sNombrePais = Trim(aLinea(17))
				sFono = ObtenerFono(Trim(aLinea(18)))
				sMensaje = Trim(aLinea(19))
				sNota = Trim(aLinea(20))
				sNombreCiudad = Trim(Replace(uCase(Trim(aLinea(23))), "DE CHILE", ""))
				sNombreCiudad = Trim(Replace(sNombreCiudad, "CHILE", ""))
				sCiudadB = ""
				sPaisB = ""
				ddiPaisB = 0
				ddiAreaB = 0
				ObtenerUbicacion Conexion, sCiudadB, sPaisB, sNombreCiudad, sNombrePais, ddiPaisB, ddiAreaB
	      
				'562 2819283
				' Jonathan Miranda G. 11-12-2006
				sNota = sNota & " Teléfono: " & sFono
				'--------------------- Fin ---------------------------
				Select Case sPaisB
					Case Session("PaisMatriz"), ""
					'******************** modificacion	   
						If Left(sFono, 3) = "569" Then
							ddiPaisB = 56
							ddiAreaB = 9
							cFonoB = CCur(Mid(sFono, 4))
						ElseIf Left(sFono, 3) = "568" Then
							ddiPaisB = 56
							ddiAreaB = 8
							cFonoB = CCur(Mid(sFono, 4))
					'******************** fin 				
						ElseIf sCiudadB = Session("CiudadMatriz") Or sCiudadB = "" Then
							cFonoB = CCur(0 & Mid(sFono, 4, 7))
						Else
							cFonoB = CCur(0 & Mid(sFono, 5, 6))
						End If
	         
					Case Else
						sFono = mid(sFono, len(Trim(cStr(ddiPaisB)))+1)
						sFono = mid(sFono, len(Trim(cStr(ddiAreaB)))+1)
						cFonoB = CCur(0 & sFono)	         
				
				End Select
				'response.Write 1 & ", " & sNombreB

			Case afxMultiexpress
				CargarMultiexpress Conexion, Linea

		
			Case afxAfexLax
				CargarAfexLax Conexion, Linea
				
			Case Else
				CargarEstandar Conexion, Linea
			
		End Select
	End Sub


	Sub CargarAfexLax(ByVal Conexion, ByVal Linea)
		Linea = UCase(Linea)
		Linea = LimpiarCampo(Linea)
		'Linea = Replace(Linea, Chr(34), " ")	      
		'Linea = Replace(Linea, Chr(126), " ")
		sInvoice = cCur(0 & Trim(Mid(Linea, 2, 15)))
		sInvoice = "0" & sInvoice 
		cMonto = cDbl(0 & Replace(Trim(mid(Linea, 17, 10)), ".", ","))	   
		sNombreR = Trim(mid(Linea, 243, 30))
		sApellidoR = Trim(mid(Linea, 213, 30))
		sNombreB = Trim(mid(Linea, 66, 30))
		sApellidoB = Trim(mid(Linea, 36, 30))
		ddiPaisB = 0
		ddiAreaB = 0
		
		sFono = ObtenerFono(Mid(Linea, 103, 10))
		
		If Mid(Trim(sFono), 1, 2) = "56" Then
			ddiPaisB = 56
			If Mid(Trim(sFono), 3, 1) = "2" Then
				ddiAreaB = 2
			Else
				ddiAreaB = 0
			End If
			cFonoB = Mid(Trim(sFono), 4)
		Else
			cFonoB = cCur(sFono)
		End If
		
		'cFonoB = TransformarFono
		
		sNombreCiudad = CorregirCiudad(Replace(uCase(Trim(Mid(Linea, 163, 25))), "DE CHILE", ""))
		sNombreCiudad = Trim(Replace(sNombreCiudad, "CHILE", ""))
		If InStr(1, uCase(sNombreCiudad), "SANTIAGO") > 0 Then 
			sNombreCiudad = "SANTIAGO"
		End If
		sDireccionB = left(Trim(mid(Linea, 113, 50)) & ", " & trim(sNombreCiudad), 50)
		sPaisB = Trim(mid(Linea, 188, 25))
		
		sCiudadB = ""
		sNombrePais = ""
		sMensaje = ""
		sNota = ""
	   
		'If Not ProcesarFono(sFono, ddiPaisB, ddiAreaB) Then
		'	'sMensaje = sMensaje & ", Telefono: " & cFonoB
		'	cFonoB = 0
		'Else
		'	cFonoB = cCur(sFono)
		'End If

		ObtenerUbicacion Conexion, sCiudadB, sPaisB, sNombreCiudad, sNombrePais, ddiPaisB, ddiAreaB
	End Sub

	Sub CargarMultiexpress(ByVal Conexion, ByVal Linea)
	   Linea = UCase(Linea)
	   Linea = LimpiarCampo(Linea)
	   Linea = Replace(Linea, Chr(34), " ")	      
	   Linea = Replace(Linea, Chr(126), " ")
	   sInvoice = Trim(Mid(Linea, 181, 8))
	   cMonto = cDbl(0 & Replace(Trim(mid(Linea, 229, 15)), ".", ","))	   
	   sNombreR = Trim(mid(Linea, 189, 40))
	   sApellidoR = ""
	   sNombreB = Trim(mid(Linea, 1, 40))
	   sApellidoB = ""
	   sNombreCiudad = CorregirCiudad(Trim(Mid(Linea, 513, 40)))
	   sDireccionB = left(Trim(mid(Linea, 61, 40)) & ", " & trim(sNombreCiudad), 40)
	   sNombrePais = ""
	   sFono = ObtenerFono(Trim(mid(Linea, 41, 20)))
	   'MS 19-05-2014 APPL-5913: El comentario se pasa al campo nota al agente para que no sea impreso en el comprobante
	   sNota = Trim(mid(Linea, 259, 254))
	   sMensaje = ""
	   'MS 19-05-2014 APPL-5913
	   sCiudadB = ""
	   sPaisB = ""
	   ddiPaisB = 0
	   ddiAreaB = 0
	   ObtenerUbicacion Conexion, sCiudadB, sPaisB, sNombreCiudad, sNombrePais, ddiPaisB, ddiAreaB
	   'cFonoB = 0
	   'cFonoB = cCur(sFono)
	   cFonoB = TransformarFono
	   If Not ProcesarFono(sFono, ddiPaisB, ddiAreaB) Then
			sNota = sNota & ", Telefono: " & cFonoB 'MS 19-05-2014 APPL-5913
			cFonoB = 0
		Else
			cFonoB = cCur(sFono)
	   End If
	End Sub

	Sub CargarEstandar(ByVal Conexion, ByVal Linea)
	   Linea = UCase(Linea)
	   Linea = LimpiarCampo(Linea)
	   Linea = Replace(Linea, Chr(34), " ")	      
	   Linea = Replace(Linea, Chr(126), " ")
	   Linea = Replace(Linea, Chr(39), " ")	   
	   sInvoice = Trim(Mid(Linea, 2, 15))
	   cMonto = cDbl(0 & Replace(Trim(mid(Linea, 17, 10)), ".", ","))	   
	   sNombreR = Trim(mid(Linea, 243, 30))
	   sApellidoR = Trim(mid(Linea, 213, 30))
	   sNombreB = Trim(mid(Linea, 66, 30))
	   sApellidoB = Trim(mid(Linea, 36, 30))
	   sNombreCiudad = CorregirCiudad(Trim(Mid(Linea, 163, 25)))
	   sDireccionB = left(Trim(mid(Linea, 113, 50)) & ", " & trim(sNombreCiudad), 25)
	   sNombrePais = Trim(mid(Linea, 188, 25))
	   sFono = ObtenerFono(Trim(mid(Linea, 103, 10)))
	   sMensaje = Trim(mid(Linea, 273, 240))
	   sNota = Trim(mid(Linea, 513, 240))
	   sCiudadB = ""
	   sPaisB = ""
	   ddiPaisB = 0
	   ddiAreaB = 0
	   ObtenerUbicacion Conexion, sCiudadB, sPaisB, sNombreCiudad, sNombrePais, ddiPaisB, ddiAreaB
	   'cFonoB = 0
	   'cFonoB = cCur(sFono)
	   cFonoB = TransformarFono
	   If Not ProcesarFono(sFono, ddiPaisB, ddiAreaB) Then
			sMensaje = sMensaje & ", Telefono: " & cFonoB
			cFonoB = 0
		Else
			cFonoB = cCur(sFono)
	   End If
	End Sub


	Function TransformarFono()
		TransformarFono = 0
		
	   '562 2819283
	   cFonoB = 0
	   Select Case sPaisB
	   Case Session("PaisMatriz"), ""	   
			If (sCiudadB = Session("CiudadMatriz") Or sCiudadB = "") Then
				If Len(sFono) >= clng(10) then
					cFonoB = CCur(Mid(sFono, 4, 7))
				Else
					cFonoB = CCur(sFono)
				End If
	      ElseIf Len(sFono) >= clng(10) then
	         cFonoB = CCur(Mid(sFono, 5, 6))
	      Else
 				cFonoB = CCur(sFono)	         
	      End If
	         
	   Case Else
			If left(sFono, len(Trim(cStr(ddiPaisB)))) = cStr(ddiPaisB) Then
				sFono = mid(sFono, len(Trim(cStr(ddiPaisB)))+1)
			End If
			If left(sFono, len(Trim(cStr(ddiAreaB)))) = cStr(ddiAreaB) Then
				sFono = mid(sFono, len(Trim(cStr(ddiAreaB)))+1)
			End If
			If sFono = "" Then sFono = 0
			cFonoB = CCur(sFono)	         
	   End Select
		TransformarFono = cFonoB
	End Function

	Function ObtenerFono(ByVal Cadena)
	   Dim i
	   Dim sChar
	   Dim sFono

	   For i = 1 To Len(Cadena)
	      sChar = Mid(Cadena, i, 1)
	      If sChar = "0" And Trim(sFono) = "" Then 
	      
	      ElseIf sChar >= "0" And sChar <= "9" Then 
				sFono = sFono + sChar
			End If
	   Next
	   If Trim(sFono) = "" Then sFono = "0"
	   ObtenerFono = sFono
	End Function

	Public Function LimpiarCampo(ByVal Cadena)
	   Cadena = Replace(Cadena, Chr(10), "")
	   Cadena = Replace(Cadena, Chr(13), "")
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


	Function ObtenerUbicacion(ByVal Conexion, _
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
	      sSQL = sSQL & " AND ci.nombre LIKE '%" & NombreCiudad & "%' "
	   End If
	   If CodigoCiudad <> "" Then
	      sSQL = sSQL & " AND ci.codigo = '" & CodigoCiudad & "' "
	   End If

	   Set rsCiudad = EjecutarSQLCliente(Conexion, sSQL)

	   rsCiudad.MoveFirst
	   If Err.Number = 0 Then			
			CodigoCiudad = rsCiudad("codigo")
			NombreCiudad = rsCiudad("Nombre")
			CodigoPais = rsCiudad("codigo_pais")
			NombrePais = rsCiudad("nombre_pais")
			ddiAreaB = cInt(rsCiudad("ddi"))
			ddiPaisB = cInt(rsCiudad("ddi_pais"))
		   ObtenerUbicacion = True
		Else
			Err.Clear 
		End If
		Set rsCiudad = Nothing
	End Function

	Function CorregirCiudad(ByVal NombreCiudad)
	
		Select Case NombreCiudad
			Case "VINA DEL MAR"
				CorregirCiudad = "VIÑA DEL MAR"
			
			Case Else
				CorregirCiudad = NombreCiudad		
				
		End Select
		
	End Function

	Sub MostrarErrorAFEX(ByRef objAFEX, ByVal Titulo)
		
		If Titulo = "" Then Titulo = "Associated Foreign Exchange"
		Response.Redirect  "http:../Compartido/Error.asp?Titulo=" & Titulo & _
					"&Number=" & objAFEX.ErrNumber  & _
					"&Source=" & objAFEX.ErrSource & _
					"&Description=" & replace(objAFEX.ErrDescription, vbCrLf , "^")			
		Set objAFEX = Nothing	
	End Sub


	Function MostrarErrorMS(ByVal Titulo)
		
		If Titulo = "" Then Titulo = "Associated Foreign Exchange"
		Response.Redirect "http:../Compartido/Error.asp?Titulo=" & Titulo & _
					"&Number=" & Err.Number  & _
					"&Source=" & Err.Source & _
					"&Description=" & Err.Description
		
	End Function
	
	Sub EnviarCorreo(ByVal Mensaje)		
		EnviarEmail "AfexMoneyWeb <transmisiones@afex.cl>", Session("NombreCliente") & " <" & Session("emailCliente") & ">", _
					"Cynthia Peña <transmisiones@afex.cl>", _
					"Carga de Giros." & Session("NombreCliente"), _
					"Sres." & vbCrLf & Session("NombreCliente") & vbCrLf & _
					Mensaje & vbCrLf & vbCrLf & "AFEX Ltda." & vbCrLf & "Santiago - Chile", 0
	End Sub


	Function ProcesarFono(ByRef s_FonoB, ByVal I_DdiPais, _
	                              ByVal I_DdiCiudad)
		Dim S_Aux
		Dim K
	  
		ProcesarFono = False
	   'procesar fono del beneficiario
		s_FonoB = Trim(s_FonoB)
	   s_FonoB = ObtenerFono(s_FonoB)
	   If s_FonoB = "" Then s_FonoB = "0"
	      s_FonoB = CStr(CCur(s_FonoB))
	    If cint(I_DdiPais) = cInt(56) Then     ' Chile
	         If cint(I_DdiCiudad) = cint(0) Then
	            If Len(s_FonoB) > 7 Then
	              If "56" = Left(s_FonoB, K) Then s_FonoB = Mid(s_FonoB, 3)
	              If Left(s_FonoB, 1) = "2" Then
	                 s_FonoB = Right(s_FonoB, 7)
	               Else
	                 s_FonoB = Right(s_FonoB, 6)
	               End If
	            End If
	         ElseIf I_DdiCiudad = 2 Then
	            s_FonoB = Right(s_FonoB, 7)
	         Else
	            s_FonoB = Right(s_FonoB, 6)
	         End If
	    Else
	       If Len(s_FonoB) > 7 Then
	           If I_DdiPais > 0 Then
	              S_Aux = CStr(I_DdiPais)
	              K = Len(S_Aux)
	              If S_Aux = Left(s_FonoB, K) Then s_FonoB = Mid(s_FonoB, K + 1)
		           If I_DdiCiudad > 0 Then
						  S_Aux = CStr(I_DdiCiudad)
						  K = Len(S_Aux)
						  If S_Aux = Left(s_FonoB, K) Then s_FonoB = Mid(s_FonoB, K + 1)
     	           End If
     	           If Len(s_FonoB) > 7 Then s_FonoB = Right(s_FonoB, 7)
     	        Else
     				  Exit Function
	           End If
	           
	       End If
	    End If	    
	   If Len(s_FonoB) < 9 Then
	     ProcesarFono = True
	   End If	 

	End Function
	
	
	'Creado:       11-12-2006  Jonathan Miranda G.
	'Objetivo:     Cargar archivos con giros usando la dll AFEXArchivos
	'Parámetros:   Archivo, dirección del archivo que se carga
	'              Agente, código del agente captador
	'              FechaTransmision
	'              NumeroTransmision
	function CargarArchivosDLL(ByVal Archivo, ByVal Agente, ByVal FechaTransmision, ByVal NumeroTransmision, Byval Usuario)
	   Dim dllArchivos
	   Dim sMensaje	   
	   Dim sMensajeError
	   Dim i 
	   
	   CargarArchivosDLL = False
	   
	   On Error Resume Next
	   
	   Set dllArchivos = server.CreateObject("AFEXArchivos.Archivo")	   
	   sMensaje = dllArchivos.CargarArchivo(Session("afxCnxAFEXpress") ,  Archivo, Agente, Session("CodigoMatriz"), _
	                                             FechaTransmision, NumeroTransmision,  1, Usuario)	   
		   
	   If Err.Number <> 0 Then
			close
			Set dllArchivos = Nothing
	      MostrarErrorMS "Cargar DLL" & agente 
	      
	   ElseIf dllArchivos.ErrNumber <> 0 Then
	      If sMensaje <> Empty Then
	         i = instr(sMensaje, ";")
	         Giros = left(sMensaje, i - 1)
	         sMensaje = mid(sMensaje, i + 1)
	         i = instr(sMensaje, ";")
	         GirosOK  = left(sMensaje, i - 1)
	         sMensaje = mid(sMensaje, i + 1)
	         i = instr(sMensaje, ";")
	         GirosError = left(sMensaje, i - 1)
	         sMensaje = mid(sMensaje, i + 1)	         
	         ListaGirosError = sMensaje
	         
	         'If CCur(sGirosError) > 0 Then
	            'GLB_Parametros = sMensaje
	            'GPRC_Parametros "", "", "", sDato
	         '   GPRC_Error Me.Caption, "Ocurrió un error al cargar el archivo " & Archivo & ". " & vbCrLf & _
	         '                          "De un total de " & Trim(sTotalGiros) & " giros, se cargaron " & Trim(sGirosCargados) & _
	         '                          " y se encontraron " & Trim(sGirosError) & " con problemas. Lea el siguiente mensaje: " & vbCrLf & _
	         '                          sMensajeError
	            'GLB_Parametros = Empty
	         'End If
	      Else
				close
				sMensajeError = dllArchivos.ErrDescription
				Set dllArchivos = Nothing
	         MostrarErrorMS "Ocurrió un error al cargar el archivo " & Archivo & ". " & sMensajeError
	      End If
	      
	   ElseIf sMensaje <> Empty Then
	      i = instr(sMensaje, ";")
	      Giros = left(sMensaje, i - 1)
	      sMensaje = mid(sMensaje, i + 1)
	      i = instr(sMensaje, ";")
	      GirosOK  = left(sMensaje, i - 1)
	      sMensaje = mid(sMensaje, i + 1)
	      i = instr(sMensaje, ";")
	      GirosError = left(sMensaje, i - 1)
	      sMensaje = mid(sMensaje, i + 1)	         
	      ListaGirosError = sMensaje
	   End If
	   
	   CargarArchivosDLL = True	
	   Err.Clear
	   Set dllArchivos = Nothing
	End Function
	

%>
<!--#INCLUDE VIRTUAL="/Compartido/Constantes.asp" -->
<!--#INCLUDE VIRTUAL="/Compartido/Rutinas.asp" -->

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" Content="VBScript">

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE=vbscript>
<!--
	sub window_onload()		
		<%if sHTML <> "" then%>
			giro.shtml.value = "<%=sHTML%>"			
			giro.action = "../compartido/informacion.asp"
			giro.submit()
			giro.action = "" 			
		<%end if%>
	end sub
-->
</SCRIPT>


</HEAD>
<BODY>
<OBJECT RUNAT=server PROGID=Scripting.FileSystemObject id=OBJECT1> </OBJECT>

<form name="giro" method="post">
	<input type="hidden" name="sHTML">
</form>
</BODY>
</HTML>