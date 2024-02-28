<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%

	Dim nAccion, sPais, sCiudad, sComuna, sCaptador, sMoneda, nMonto,sMonedaNacional
	Dim nTarifaSugerida, nTarifaCobrada, bCliente, bExtranjero
	Dim nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador
	Dim nComisionMatriz, nAfectoIva
	Dim sNombreR, sApellidoR, sDireccionR, sAFEXpress, sAFEXchange
	Dim sNombres, sApellidoP, sApellidoM, sDireccion, sPaisR, sCiudadR
	Dim sPaisFono, sAreaFono, sFono, sRut
	Dim sNombrePais, sNombreCiudad, sNombreComuna, sNombrePaisPass
	Dim tipoP,sColorMoneda, sMonedaMG
	'*************************************************************
	Dim sTipoAgente, sTitulo, sDisplay
	'*************************************************************
	
	' JFMG 13-05-2009 se agrega la declaración, ya que no existía pero si se utilizaba la variable
	Dim sPasaporte, sPaisPass
	' *************** FIN 13-05-2009 ********************
	
	
	sTipoAgente = Request("ag")
	
	Select Case sTipoAgente
		Case Trim(Session("CodigoMGPago"))
			sTitulo = "Pago MoneyGram"
			sDisplay = "none"
			tipoP=" "
		Case Trim(Session("CodigoTXPago"))
			sTitulo = "MoneyBroker"
			sDisplay = ""
			tipoP="TMT "
		Case Trim(Session("CodigoUTPago"))
			sTitulo = "U n i t e l l e r"
			sDisplay = ""
			tipoP=" "			
		Case Else
			sTitulo = ""
			
	End Select
	
	nAccion = cInt(0 & Request("Accion"))
	nTarifa = cCur(0)
	nTarifaCobrada = cCur(0)
	nMonto = cCur(0)
	nGastoTransfer = cCur(0)
	nComisionCaptador = cCur(0)
	nComisionPagador = cCur(0)
	nComisionMatriz = cCur(0)
	nAfectoIva = cCur(0)
	sCaptador = Session("CodigoMGPago")
	'sMoneda = Session("MonedaExtranjera")
	'sMonedaNacional=session("MonedaNacional")
	sPaisR = Session("PaisCliente")
	sCiudadR = Session("CiudadCliente")
	sAFEXpress = ""
	sAFEXchange = ""
	
	Select Case nAccion
	Case afxAccionPais, afxAccionCiudad
			Cargar
			CalcularComision
			
	Case afxAccionMonto	
			Cargar
			CalcularComision
			
	Case Else
			If Not CargarCliente() Then
				Response.Redirect "AtencionClientes.asp"
			End If
			
			If CargarUltimoRecibo() Then
				'CalcularComision
			End If			
			
	End Select
	
	Function CargarCliente()
		Dim nCampo, rs, sArgumento, sArgumento2, sArgumento3
		
		CargarCliente = False
		nCampo = cInt(0 & request("Campo"))
		sArgumento = request("Argumento")
		sArgumento2 = request("Argumento2")
		sArgumento3 = request("Argumento3")		
		
		'Response.Write nCampo & ", " & sArgumento & ", " & sArgumento2 & ", " & sArgumento3
		'Response.end
		
		Set rs = BuscarCliente(nCampo, sArgumento, sArgumento2, sArgumento3)
		If rs.EOF Then
			Set rs = Nothing
			Exit Function
		End If
		
		If rs.RecordCount > 1 Then
			Set rs = Nothing
			Response.Redirect "ListaClientes.asp?Accion=" & afxAccionIngresarMG & "&Campo=" & nCampo & _
									"&Argumento=" & sArgumento & _
								   "&Argumento2=" & sArgumento2 & _
									"&Argumento3=" & sArgumento3 & _
									"&Titulo=Lista de Clientes" & _
									"&ag=" & sTipoAgente
		End If
		
		'Response.Redirect "../Compartido/Error.asp?Titulo=" & rs.eof
		If Not rs.EOF Then
			nTipoCliente = cInt(0 & rs("tipo"))
			If nTipoCliente = 1 Then
				sApellidoP = MayMin(EvaluarVar(rs("paterno"), ""))
				sApellidoM = MayMin(EvaluarVar(rs("materno"), ""))
				sNombres = MayMin(EvaluarVar(rs("nombre"), ""))
			Else
				sRazonSocial = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			End If	
			sNombreCompleto = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			sRut = FormatoRut(EvaluarVar(rs("rut"), ""))
			sPasaporte = EvaluarVar(rs("pasaporte"), "")
			sPaisPass = EvaluarVar(rs("codigo_paispas"), "")
			sDireccion = MayMin(EvaluarVar(rs("direccion"), ""))
			sPais = EvaluarVar(rs("codigo_pais"), "")
			sCiudad = EvaluarVar(rs("codigo_ciudad"), "")
			sComuna = EvaluarVar(rs("codigo_comuna"), "")
			sPaisFono = EvaluarVar(rs("ddi_pais"), "")
			sAreaFono = EvaluarVar(rs("ddi_area"), "")
			sFono = EvaluarVar(rs("telefono"), "")
			sNombrePais = MayMin(EvaluarVar(rs("pais"), ""))
			sNombreCiudad = MayMin(EvaluarVar(rs("ciudad"), ""))
			sNombreComuna = MayMin(EvaluarVar(rs("comuna"), ""))
			sNombrePaisPass = MayMin(EvaluarVar(rs("paispas"), ""))
			sAFEXchange = rs("Exchange")
			sAFEXpress = rs("Express")
			If Err.number <> 0 Then
				MostrarErrorMS ""
			End If
		End If
		Set rs = Nothing
		CargarCliente = True
	End Function

	Sub Cargar()
		sAFEXpress = Request.form("txtExpress")
		sAFEXchange = Request.form("txtExchange")
		sNombreCompleto = Request.form("txtNombreCompleto")
		sRut = request.Form("txtRut")
		sPasaporte = request.Form("txtPasaporte")
		sNombres = request.Form("txtNombres")
		sApellidoP = request.Form("txtApellidoP")
		sApellidoM = request.Form("txtApellidoM")
		sDireccion = request.Form("txtDireccion")
		sPaisPass = Request.Form("cbxPaisPasaporte")
		sPais = Request.Form("cbxPais")
		sCiudad = Request.form("cbxCiudad")
		sComuna = Request.form("cbxComuna")			
		sPaisFono = request.Form("txtPaisFono")
		sAreaFono = request.Form("txtAreaFono")
		sFono = request.Form("txtFono")
		If Request.Form("optPersona") = "on" Then
			nTipoCliente = 1
		Else
			nTipoCliente = 2
		End If
		sPaisR = Request("PaisR")
		sCiudadR = Request("CiudadR")
		If sPaisR = "" Then
			sPaisR = Request.Form("cbxPaisR")
			sCiudadR = Request.Form("cbxCiudadR")
		End If
		
		'If Request.Form("cbxMonedaGiro") = "" then
		If Request.Form("cbxMonedaPago") = "" then
			sMoneda = session("MonedaNacional")
			sColormoneda="DodgerBlue"
		else
		
			'If Request.Form("cbxMonedaGiro")= sMonedaNacional Then
			If Request.Form("cbxMonedaPago")= session("MonedaNacional") Then
				sMoneda = session("MonedaNacional")
				nMonto = round(cDbl(0 & Request.Form("txtMonto")))
				sColormoneda="DodgerBlue"
			else		
				sMoneda = session("MonedaExtranjera")
				nMonto = cDbl(0 & Request.Form("txtMonto"))		
				scolormoneda= "#4dc087"
			End if
		
		End if
		
		sNombreR = Request.Form("txtNombreR")
		sApellidoR = Request.Form("txtApellidoR")
		'sDireccionR = Request.Form("txtDireccionR")
		sNombrePais = Request.Form("txtNombrePais")
		sNombreCiudad = Request.Form("txtNombreCiudad")
		sNombreComuna = Request.Form("txtNombreComuna")
		sNombrePaisPass = Request.Form("txtNombrePaisPass")
		'CalcularComision
	End Sub
	
	Function CargarUltimoRecibo()
		Dim rsUG
		
		CargarUltimoRecibo = False		
		If IsNull(sAFEXpress) Then Exit Function
		If Trim(sAFEXpress) = "" Then Exit Function
				
		'Set rsUG = ObtenerListaGiros(afxGirosRecibidos, sAFEXpress,  _
		'							 Session("CodigoMGPago"), 1)
		Set rsUG = ObtenerListaGiros(afxGirosRecibidos, sAFEXpress,  _
									 sTipoAgente, 1)
		If Not rsUG.EOF Then
			'If Trim(rsUG("pais_remitente")) = Session("PaisCliente") _
			'And Trim(rsUG("ciudad_remitente")) = Session("CiudadCliente") Then
				sNombreR = MayMin(rsUG("nombre_remitente"))
				sApellidoR = MayMin(rsUG("apellido_remitente"))
				sDireccionR = MayMin(rsUG("direccion_remitente"))
				'nMonto = cCur(0 & rsUG("monto_giro"))
				If IsNull(rsUG("pais_remitente")) Then
					sPaisR = Empty
				Else
					sPaisR = rsUG("pais_remitente")
				End If
				If IsNull(rsUG("ciudad_remitente")) Then
					sCiudadR = Empty
				Else
					sCiudadR = rsUG("ciudad_remitente")
				End If
				CargarUltimoRecibo = True			
			'End If
		End If
		Set rsUG = Nothing
	End Function
		

	Sub CalcularComision
		
		If sPais <> "" And sCiudad <> "" And nMonto <> 0 Then
			'response.redirect "../compartido/error.asp?description=" & Session("PaisCliente") & ", " & Session("CiudadCliente") & ", " & nMonto
			nComisionPagador = ObtenerComisionMG(Session("PaisCliente"), Session("CiudadCliente"), nMonto,Request.Form("cbxMonedaPago"))
		End If				
		
	End Sub
	
	If Request.Form("cbxMonedaPago") = "" Then
		sMONEDA = session("MonedaNacional")
		sColormoneda="DodgerBlue"
	else 
		
		If Request.form("cbxMonedaPago") = session("MonedaNacional") then
		'If Request.form("cbxMonedaGiro") = sMonedaNacional then
			sMoneda = session("MonedaNacional")
			sColormoneda="DodgerBlue"
		else
			sMoneda = session("MonedaExtranjera")
			scolormoneda= "#4dc087"
		End IF
	End if 
	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<script LANGUAGE="VBScript">
<!--
	Sub window_onLoad()
		'CargarDatos
		CalcularTotal
		ActivarFoco	
		
		If "<%=sPasaporte%>" <> Empty Then ' JFMG 13-05-2009 se comenta ya que esta variable estaba mal usada. If "<%=sPass%>" <> Empty Then
			optPasaporte_onClick 
		Else
			optRut_onClick
		End If
		colormoneda "<%=sColorMoneda%>"
		If frmGiro.txtMonto.value <> 0  Then
			'If frmGiro.cbxMonedaPago.value = "<%=sMonedaExtranjera%>" Then
			If frmGiro.cbxMonedaPago.value = "<%=sMoneda%>" Then
				frmGiro.txtMonto.value = FormatNumber(frmGiro.txtMonto.value,2)
			Else
				frmGiro.txtMonto.value = frmGiro.txtMonto.value
			End If
		End IF
	End Sub

	Sub optRut_onClick()
		frmGiro.optpasaporte.checked = False
		frmGiro.optRut.checked = True
		frmGiro.txtpasaporte.style.display="none"
		frmGiro.txtNombrePaisPass.style.display="none"
		lblPaisPasaporte.style.display="none"
		frmGiro.txtRut.style.display = ""		
	End Sub
	
	Sub optPasaporte_onClick()
		frmGiro.optpasaporte.checked = True
		frmGiro.optRut.checked=False
		frmGiro.txtRut.style.display = "none"		
		frmGiro.txtpasaporte.style.display=""
		frmGiro.txtNombrePaisPass.style.display=""
		lblPaisPasaporte.style.display=""
	End Sub
	
	Sub CargarDatos()
		frmGiro.txtExchange.value = "<%=sAFEXchange%>"
		frmGiro.txtExpress.value = "<%=sAFEXpress%>"
		frmGiro.txtNombreCompleto.value = "<%=sNombreCompleto%>"
		frmGiro.txtApellidoM.value = "<%=sApellidoM%>"
		frmGiro.txtApellidoP.value = "<%=sApellidoP%>"
		frmGiro.txtNombres.value = "<%=sNombres%>"
		frmGiro.txtDireccion.value = "<%=sDireccion%>"
		frmGiro.txtPaisFono.value = "<%=ObtenerDDI(1, sPais)%>" 
		frmGiro.txtAreaFono.value = "<%=ObtenerDDI(2, sCiudad)%>"
		frmGiro.txtFono.value = "<%=sFono%>"
		frmGiro.txtRut.value = "<%=sRut%>"
		frmGiro.txtPasaporte.value = "<%=sPasaporte%>"		
	End Sub
	
	Sub ActivarFoco()
		<% Select Case nAccion %>
		<%	Case afxAccionPais %>
				'frmGiro.txtMonto.select
				'frmGiro.cbxMonedaGiro.select
				
		<%	Case afxAccionMonto %>
				frmGiro.txtInvoiceMG.select				
		<% Case Else %>
				If frmGiro.txtNombreR.value <> "" Then
					'frmGiro.txtMonto.focus()
					'frmGiro.txtMonto.select()
					frmGiro.cbxMonedaPago.focus()
					
				Else
					frmGiro.txtNombreR.focus()
					frmGiro.txtNombreR.select()
				End If
		<% End Select %>		
	End Sub 

	Sub txtMonto_onBlur()
		Dim nPos
		
		If Trim(frmGiro.txtMonto.value) = "" Then
			frmGiro.txtMonto.value = "0"
		End If
		
		'If frmGiro.cbxMonedaPago.value = "<%=sMonedaNacional%>" Then
		
		If frmGiro.cbxMonedaPago.value = "<%=session("MonedaNacional")%>" Then
			frmGiro.txtMonto.value = round(frmGiro.txtMonto.value)
			Colormoneda "DodgerBlue"
			sMoneda ="<%=session("MonedaNacional")%>"
		Else
			 If frmGiro.txtMonto.value = FormatNumber(frmGiro.txtMonto.value, 2) Then
				colormoneda "#4dc087"
				sMoneda = "<%=session("MonedaExtranjera")%>"
			End if
		End iF
		If cCur(frmGiro.txtMonto.value) <> cCur(0 & "<%=nMonto%>") Then
			HabilitarControles
			'frmGiro.action = "IngresarGiroMG.asp?Accion=<%=afxAccionMonto%>"
			frmGiro.action = "IngresarGiroMG.asp?Accion=<%=afxAccionMonto%>&ag=<%=sTipoAgente%>"
			frmGiro.submit 
			frmGiro.action = ""
		End IF
	End Sub 
	
	Sub txtComisionPagada_OnBlur()
		If Trim(frmGiro.txtComisionPagada.value) = "" Then
			frmGiro.txtComisionPagada.value = "0,00"
		Else
			frmGiro.txtComisionPagada.value = FormatNumber(frmGiro.txtComisionPagada.value, 2)
		End If
		CalcularTotal	
	End Sub 
	
	Sub imgCalcular_onClick()
		CalcularTotal
	End Sub

	Sub CalcularTotal()
		Dim nCobrada, nMonto
		If frmGiro.cbxMonedaPago.value = "<%=session("MonedaNacional")%>" Then
		    'msgbox "<%=session("MonedaNacional")%>" & "," & frmGiro.cbxmonedapago.value
			frmGiro.txtComisionGanada.value = round(FormatNumber("<%=nComisionPagador%>",2))
			nMonto = (cDbl(0 & round(frmGiro.txtMonto.value)))
			nCobrada = round(cDbl(0 & frmGiro.txtComisionGanada.value))
			frmGiro.txtTotal.value = round(FormatNumber(nMonto + nCobrada, 2))
		Else
			if frmGiro.cbxmonedapago.value = "<%=session("MonedaExtranjera")%>" Then
			'	msgbox "<%=session("MonedaExtranjera")%>" & "," & frmGiro.cbxmonedapago.value
				frmGiro.txtComisionGanada.value = FormatNumber("<%=nComisionPagador%>", 2)
				nMonto = cDbl(0 & frmGiro.txtMonto.value)
				nCobrada = cDbl(0 & frmGiro.txtComisionGanada.value)
				frmGiro.txtTotal.value = FormatNumber(nMonto + nCobrada, 2)
			end if 
		End If
	End Sub
 
	Sub cbxPaisR_onblur()
		Dim sCiudad
		
		If frmGiro.cbxPaisR.value = "" Then Exit Sub
		If frmGiro.cbxPaisR.value = "<%=sPaisR%>" Then Exit Sub
		'If cDbl(0 & frmGiro.txtmonto.value) = 0 Then Exit Sub
		HabilitarControles
		frmGiro.cbxCiudadR.value = ""
		'frmGiro.action = "IngresarGiroMG.asp?Accion=<%=afxAccionPais%>"
		frmGiro.action = "IngresarGiroMG.asp?Accion=<%=afxAccionPais%>&ag=<%=sTipoAgente%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	Sub cbxMonedaPago_onchange()
		'if frmGiro.cbxMonedaPago.value ="<%=sMonedaNacional%>" then
		'msgbox frmGiro.cbxMonedaPago.value  &"," & "<%=session("Monedanacional")%>"
		if frmGiro.cbxMonedaPago.value ="<%=session("MonedaNacional")%>" then
			Colormoneda "DodgerBlue"
		else
			colormoneda "#4dc087"
		End IF
		frmGiro.txtMonto.value = ""
		frmGiro.txtComisionGanada.value= ""
		frmGiro.txtTotal.value = ""
	end sub
	
	Sub ColorMoneda (byval color)
			frmGiro.cbxMonedaPago.style.backgroundColor =color
			frmGiro.txtMonto.style.backgroundColor=color
			frmGiro.txtBoleta.style.BackgroundColor=color
	End sub

	Sub cbxCiudadR_onblur()
		Dim sComuna
		
		If frmGiro.cbxCiudadR.value = "" Then Exit Sub
		If frmGiro.cbxCiudadR.value = "<%=sCiudadR%>" Then Exit Sub		
		HabilitarControles
		frmGiro.action = "IngresarGiroMG.asp?Accion=<%=afxAccionCiudad%>&ag=<%=sTipoAgente%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	
	Sub UltimosGiros()
		Dim sString, aGiro, sNombre, sCliente
	
		sCliente = Trim(frmGiro.txtExpress.value)
		sNombre = Trim(Trim(frmGiro.txtnombres.value) & " " & Trim(frmGiro.txtApellidos.value))
		
		sString = Empty
		sString = window.showModalDialog("../Compartido/UltimosGiros.asp?CodigoCliente=" & sCliente & _
																	"&NombreCliente=" & sNombre & _
																	"&CodigoMoneda=" & sMoneda & _
																	"&TipoGiro=<%=afxListaGirosRecibidos%>")
		If sString <> Empty Then
			' crea un arreglo con los datos de la transferencia seleccionada
			aGiro = Split(sString, ";", 13)
			
			' asigna los datos al envio
			window.frmGiro.txtNombreR.value = aGiro(0)
			window.frmGiro.txtApellidoR.value = aGiro(1)
			window.frmGiro.txtDireccionR.value = aGiro(2)
			window.frmGiro.cbxPaisR.value = aGiro(3)
			window.frmGiro.cbxCiudadR.value = aGiro(4)
			'window.frmGiro.txtPaisFonoB.value = aGiro(5)
			'window.frmGiro.txtAreaFonoB.value = aGiro(6)
			'window.frmGiro.txtFonoB.value = aGiro(7)
			window.frmGiro.txtMonto.value = aGiro(8)
			
			HabilitarControles
			'frmGiro.action = "IngresarGiroMG.asp?Accion=<%=afxAccionMonto%>"
			frmGiro.action = "IngresarGiroMG.asp?Accion=<%=afxAccionMonto%>&ag=<%=sTipoAgente%>"
			frmGiro.submit 
			frmGiro.action = ""
		End If		
	End Sub


	Sub imgAceptar_OnClick()

		'Validaciones 
		If Not ValidarDatos Then
			Exit Sub
		End If
		HabilitarControles
		<% If Session("ModoPrueba") Then 
				If sTipoAgente = Trim(Session("CodigoMGPago")) Then %>
				
					frmGiro.action = "GrabarIngresoGiroMG.asp"
					
			<%  Elseif sTipoAgente = Trim(Session("CodigoUTPago")) Then%>
			
					frmGiro.action = "GrabarIngresoGiroUT.asp"
			<%else%>
					frmGiro.action = "GrabarIngresoGiroTX.asp"
			<%  End If %>
		<% Else 
				If sTipoAgente = Trim(Session("CodigoMGPago")) Then %>
					frmGiro.action = "GrabarIngresoGiroMG.asp"
					
				<%Elseif sTipoAgente = Trim(Session("CodigoUTPago")) Then%>
				
					frmGiro.action = "GrabarIngresoGiroUT.asp"
				<%else%>
					frmGiro.action = "GrabarIngresoGiroTX.asp"
		<%  End If %>
		<% End If %>
		frmGiro.submit 
		frmGiro.action = ""
		
	End Sub

	Function ValidarDatos()
		
		ValidarDatos = False
		If Trim(frmGiro.txtNombres.value) = "" Then
			MsgBox "Debe ingresar el nombre del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtApellidos.value) = "" Then
			MsgBox "Debe ingresar apellidos del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtNombreR.value) = "" Then
			MsgBox "Debe ingresar el nombre del remitente",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtApellidoR.value) = "" Then
			MsgBox "Debe ingresar apellidos del remitente",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.cbxPaisR.value) = "" Then
			MsgBox "Debe ingresar el pais del remitente",,"AFEX"
			Exit Function
		End If
		'If Trim(frmGiro.cbxCiudadR.value) = "" Then
		'	MsgBox "Debe ingresar la ciudad del remitente",,"AFEX"
		'	Exit Function
		'End If
		If cCur(0 & Trim(frmGiro.txtMonto.value)) = 0 Then
			MsgBox "Debe ingresar el monto del giro",,"AFEX"
			Exit Function
		End If			

		If ( trim(frmGiro.txtMonto.value))>10000 and frmGiro.cbxMonedaPago.value = "<%=session("MonedaExtranjera")%>" Then
				msgbox "El monto ingresado no corresponde" , , "AFEX"
				Exit Function
		End If			

		If cCur(0 & Trim(frmGiro.txtComisionGanada.value)) = 0 Then
			MsgBox "Se debe calcular la comision",,"AFEX"
			Exit Function
		End If
		<%	If sTipoAgente = Trim(Session("CodigoMGPago")) Then %>
				If Len(Trim(frmGiro.txtInvoiceMG.value)) <> 8 Then
					MsgBox "El invoice debe ser de 8 dígitos", ,"AFEX"
					Exit Function
				End If
		<%elseif sTipoAgente = Trim(Session("CodigoUTPago")) Then%>		
				If Len(Trim(frmGiro.txtInvoiceMG.value)) <> 10 Then
					MsgBox "El invoice debe ser de 10 dígitos", ,"AFEX"
					Exit Function
				End If
		<%	Else %>
			cont = (Len(Trim(frmGiro.txtInvoiceMG.value)))
			j=0
				for i=1 to cont
					if (isnumeric(Mid(Trim(frmGiro.txtInvoiceMG.value), i,i))= false )then  
					   j=j+1 					
 					end if
 				next		
 			    if j>0 then 
 				  MsgBox "Debe ingresar solo Números en el invoice", ,"AFEX"
 				  Exit Function
 				end if  
				If Len(Trim(frmGiro.txtInvoiceMG.value)) <> 9 Then
					MsgBox "El invoice debe ser de 9 dígitos", ,"AFEX"
					Exit Function
				End If
		<%	End If %>
		If Trim(frmGiro.txtInvoiceMG.value) = "" Then
			MsgBox "Debe ingresar el invoice del giro",,"AFEX"
			Exit Function
		End If				
		ValidarDatos = True
	End Function	
-->
</script>

<body>
<table><tr><td>
<marquee STYLE="HEIGHT: 60px; LEFT: 2px; POSITION: absolute; TOP: 0px; WIDTH: 403px" BEHAVIOR="slide" DIRECTION="right" SCROLLAMOUNT="2000" SCROLLDELAY="1">
	<h6 class="TituloFondo" STYLE="FONT-SIZE: 60px; top: -22px">Transacciones</h6>
</marquee>		
<marquee STYLE="HEIGHT: 55px; LEFT: 21px; POSITION: absolute; TOP: 3px; WIDTH: 215px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="2000" SCROLLDELAY="1">		
	<!--<h1 STYLE="COLOR: #cfcfcf; FONT-SIZE: 25px">Pago MoneyGram</h1>-->
	
	<h1 STYLE="COLOR: #cfcfcf; FONT-SIZE: 25px"><%=sTitulo%></h1>
	
</marquee>
<marquee STYLE="HEIGHT: 55px; LEFT: 20px; POSITION: absolute; TOP: 2px; WIDTH: 215px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="2000" SCROLLDELAY="1">		
	<!--<h1 STYLE="COLOR: steelblue; FONT-SIZE: 25px">Pago MoneyGram</h1>-->
	<h1 STYLE="COLOR: steelblue; FONT-SIZE: 25px"><%=sTitulo%></h1>
</marquee>
</td></tr></table>
<form id="frmGiro" method="post">
<input type="hidden" name="txtExchange" value="<%=sAFEXchange%>">
<input type="hidden" name="txtExpress" value="<%=sAFEXpress%>">
<input type="hidden" name="txtApellidoP" value="<%=sApellidoP%>">
<input type="hidden" name="txtApellidoM" value="<%=sApellidoM%>">
<input type="hidden" name="txtDireccion" value="<%=sDireccion%>">
<input type="hidden" name="cbxComuna" value="<%=sComuna%>">
<input type="hidden" name="cbxCiudad" value="<%=sCiudad%>">
<input type="hidden" name="cbxPais" value="<%=sPais%>">
<input type="hidden" name="cbxPaisPasaporte" value="<%=sPaisPass%>">
<input type="hidden" name="txtPaisFono" value="<%=sPaisFono%>">
<input type="hidden" name="txtAreaFono" value="<%=sAreaFono%>">
<input type="hidden" name="txtFono" value="<%=sFono%>">
	<table class="borde" ID="tabPaso1" cellspacing="0" cellpadding="1" style="position: relative; top: 20; left: 4;">
		<tr HEIGHT="10"><td colspan="5" class="Titulo">Datos del Beneficiario</td></tr>
		<tr><td colspan="5">
			<table cellspacing="0" cellpadding="0">
			<tr HEIGHT="15">
				<td></td>
				<td VALIGN="center" colspan="3">
					<table border="0" cellpadding="1">
					<tr style="display: none">
						<td>
							<input TYPE="radio" name="optRut" disabled>Rut
							<input TYPE="radio" name="optPasaporte" disabled>Pasaporte
						</td>
						<td id="lblPaisPasaporte" style="display: none">Pais</td>
					</tr>
					<tr>
						<td>
							<input name="txtRut" style="width: 150px; text-align:right" value="<%=sRut%>" disabled>
							<!-- ' JFMG 13-05-2009 se comenta ya que esta variable estaba mal usada. -->
							<!--<input name="txtPasaporte" style="width: 150px; display: none" value="<%=sPass%>" disabled>-->
							<input name="txtPasaporte" style="width: 150px; display: none" value="<%=sPasaporte%>" disabled>
							<!-- ************************ FIN 13-05-2009 ************************** -->
							<input name="txtNombrePaisPass" style="width: 150px; display: none" value="<%=sNombrePaisPass%>" disabled>
						</td>
						<td><img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" style="cursor: hand" WIDTH="19" HEIGHT="22" onClick="UltimosGiros"></td>
						<td style="cursor: hand" onClick="UltimosGiros">Ultimos Giros</td>
					</tr>
					</table>
				</td>
			</tr>		
			<tr HEIGHT="15">
				<td></td>
				
				<td VALIGN="center" colspan="3">
				<table border="0" cellpadding="0">
				<tr>
					<td></td>
					<td colspan="2">Nombres<br><input NAME="txtNombres" style="HEIGHT: 22px; WIDTH: 280px" disabled value="<%=sNombres%>"></td>
					<td>Apellidos<br><input NAME="txtApellidos" style="HEIGHT: 22px; WIDTH: 200px" disabled value="<%=Trim(sApellidoP & " " & sApellidoM)%>"></td>
				</tr>
				<tr HEIGHT="15">
					<td></td>
					<td>Pais<br><input NAME="txtNombrePais" style="HEIGHT: 22px; WIDTH: 150px" disabled value="<%=sNombrePais%>"></td>
					<td>Ciudad<br><input NAME="txtNombreCiudad" style="HEIGHT: 22px; WIDTH: 150px" disabled value="<%=sNombreCiudad%>"></td>
					<td>Comuna<br><input NAME="txtNombreComuna" style="HEIGHT: 22px; WIDTH: 150px" disabled value="<%=sNombreComuna%>"></td>
				</tr>
				</table>
				</td>
			</tr>
			</table>
		</td></tr>
		<tr HEIGHT="10"><td colspan="5" class="Titulo">Datos del Remitente</td>
		</tr>
		<tr><td colspan="5">
		<table cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td></td>
			<td colspan="2">Nombres<br><input NAME="txtNombreR" SIZE="25" style="HEIGHT: 22px; WIDTH: 300px" onKeyPress="IngresarTexto(2)" onBlur="window.frmGiro.txtNombreR.value=MayMin(Trim(window.frmGiro.txtNombreR.value))" value="<%=sNombreR%>"></td>
			<td colspan="2">Apellidos<br><input NAME="txtApellidoR" SIZE="25" style="HEIGHT: 22px; WIDTH: 200px" onKeyPress="IngresarTexto(2)" onBlur="window.frmGiro.txtApellidoR.value=MayMin(Trim(window.frmGiro.txtApellidoR.value))" value="<%=sApellidoR%>"></td>
		</tr>
		</table>
		</td></tr>
		<tr><td colspan="5">
		<table border="0" cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td></td>
			<td COLSPAN="2" style="display: none">Dirección<br>
			<input NAME="txtDireccionR" SIZE="50" style="HEIGHT: 22px; WIDTH: 280px" onBlur="window.frmGiro.txtDireccionR.value=MayMin(Trim(window.frmGiro.txtDireccionR.value))" value="<%=sDireccionR%>" disabled></td>
			<td>Pais<br>
				<select name="cbxPaisR" style="width: 130px">
					<%	
						CargarUbicacion 1, "", sPaisR
					%>
				</select>
			</td>
			<td colspan="1" style="display: <%=sDisplay%>">Ciudad<br>
				<select name="cbxCiudadR" style="width: 130px">
					<%	
						CargarCiudadesPais sPaisR, sCiudadR
					%>
				</select>
			</td>
			
		</tr>
		</table>
		</td></tr>
		<tr><td colspan="5">
		<table border="0" cellspacing="0" cellpadding="1">
		<tr HEIGHT="15" style="display: none">
			<td></td>
			<td>Mensaje al Beneficiario<br><input name="txtMensajeB" style="font-family: verdana; font-size: 9pt; HEIGHT: 21px; WIDTH: 540px" SIZE="255" value="<%=Request.Form("txtMensajeB")%>">
			</td>
		</tr>
		</table>
		</td></tr>
		<tr HEIGHT="15">
			<td colspan="5" class="Titulo">Datos del Agente Pagador</td>
		</tr>
		<tr><td colspan="5">
		<table cellspacing="0" cellpadding="1" id="tbPagador">
		<tr><td colspan="5">
		<table cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td></td>
			<td valign="bottom" >Moneda de Pago<br>
				<select name="cbxMonedaPago" style="width: 130px background-color:<%=sColorMoneda%>" >
					<%	
							CargarMonedaGiro session("CodigoAgente"), "CL",sCiudad, sMoneda
					%>
				</select>
			</td>
			<td>Monto<br>
				<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" NAME="txtMonto" SIZE="15" onkeypress="IngresarTexto(1)" value="<%=nMonto%>">
			</td>
			<td>Comision Ganada<br>
				<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" NAME="txtComisionGanada" SIZE="15" disabled value="<%=nComisionPagador%>">
			</td>
			<td style="display:none">Comision Pagada<br>
				<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" NAME="txtComisionPagada" SIZE="15" value="<%=nComisionCaptador%>">
			</td>
			<td>Total<br>
				<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 103px" NAME="txtTotal" SIZE="15" disabled>
			</td>
			<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Invoice<br>								
			&nbsp;&nbsp;
				<%=tipoP%>&nbsp;<input STYLE="HEIGHT: 22px; TEXT-ALIGN: left; WIDTH: 107px" NAME="txtInvoiceMG" SIZE="15" maxlength="10" value="<%=Request.Form("txtInvoiceMG")%>">
				
			</td>
			<td style="display: none">Nº Boleta<br>
				<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" name="txtBoleta" SIZE="15" value="<%=Request.Form("txtBoleta")%>" disabled>
			</td>
		</tr>
		</table>
		</td></tr>
		<tr><td colspan="5">
		<table cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td></td>
			<td>Nota al Agente<br><input name="txtMsjPagador" style="font-family: verdana; font-size: 9pt; HEIGHT: 21px; WIDTH: 450px" SIZE="255" value="<%=Request.Form("txtMsjPagador")%>">
			</td>
			<td width="100%"></td>
			<td align="rigth"><img border="0" id="imgAceptar" onclick src="../images/BotonAceptar.jpg" style="cursor: hand" WIDTH="70" HEIGHT="20"></td>
		</tr>
		</table>
		</td></tr>
		<tr HEIGHT="0">
		</tr>
	</table>	
</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</html>
