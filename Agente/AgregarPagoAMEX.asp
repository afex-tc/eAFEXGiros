<%@ Language=VBScript%>
<!-- #INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	dim sSQL
	dim rs
	dim iFormaPago, iMonedaPago, sNumeroTarjeta, cMontoDolares, cMontoPesos, cTipoCambio, sNombreSucursal, sNombreTitular, sTelefonoTitular, sMensajeInicio, _
		sFecha
	
	on error resume next
	
	' setea algunos valores			
		' forma pago
	if Request.Form("rbtEfectivo") = "on" then
		iFormaPago = 1	' efectivo
	elseif Request.Form("rbtCheque") = "on" then
		iFormaPago = 2	' cheque
	end if
		' moneda pago
	if Request.Form("rbtDolar") = "on" then
		iMonedaPago = 1	' dolar
	elseif Request.Form("rbtPeso") = "on" then
		iMonedaPago = 2	' peso
	end if
	cMontoDolares = Request.Form("txtMontoDolares")
	cMontoPesos = Request.Form("txtMontoPesos")
	cTipoCambio = Request.Form("txtTipoCambio")
	sNombreTitular = Request.Form("txtNombreTitular")
	sTelefonoTitular = Request.Form("txtTelefonoTitular")
	sfecha = Request.Form ("txtFecha")
	
	
	if request("Accion") <> "" then
		' setea algunos valores			
		sNumeroTarjeta = trim(Request.Form("txtNumeroTarjeta"))
	
		' verifica que hace con el rellamado
		select case request("Accion") 
			case 1	' buscar el detalle de un pago
				sSQl = "exec MostrarPagoAMEXCodigo " & request("Pago")
				set rs = ejecutarsqlcliente(Session("afxCnxAFEXchangeMCF"), sSQL)
				if err.number <> 0 then
					set rs = nothing
					MostrarErrorMS "Error al buscar el PagoAMEX. "
				else
					if not rs.eof then
						iFormaPago = rs("formapago")
						iMonedaPago = rs("monedapago")
						cMontoDolares = rs("montodolares")
						cMontoPesos = rs("montopesos")
						cTipoCambio = rs("tipocambio")
						sNombreTitular = rs("nombretitular")
						sTelefonoTitular = rs("telefonotitular")
						sNumeroTarjeta = rs("numerotarjeta")
						sFecha = rs("fecha")
						rs.close						
					end if
					set rs = nothing					
				end if
				
			case 2	' graba el pago			
				' verifica el cliente 
				'sCodigoCliente = AgregarClienteMCF					
				
				sfecha = Request.Form ("txtfecha")
					
				' verifica si existe el pago
				if not VerificarExistenciaPago then
			
					' graba el pago
					sSQL = "exec InsertarPagoAMEX " & evaluarstr(sNumeroTarjeta) & ", " & iFormaPago & ", " & iMonedaPago & ", " & _
														formatonumerosql(ccur(Request.Form("txtMontoDolares"))) & ", " & formatonumerosql(ccur(Request.Form("txtMontoPesos"))) & ", " & _
														formatonumerosql(ccur(Request.Form("txtTipoCambio"))) & ", " & evaluarstr(Request.Form("txtNombreSucursal")) & ", " & _
														evaluarstr(Request.Form("txtNombreTitular")) & ", " & evaluarstr(Request.Form("txtTelefonoTitular")) & ", " & _
														evaluarstr(Session("CodigoAgente")) & ", " & Evaluarstr(sFecha)
'esponse.Write sSQL
'esponse.End 														
					set rs = ejecutarsqlcliente(Session("afxCnxAFEXchangeMCF"), sSQL)
					if err.number <> 0 then
						set rs = nothing
						MostrarErrorMS "Error al grabar el PagoAMEX. " & ssql
					else			
						set rs = nothing
						'sFecha = date
						'Response.Redirect "AtencionClientes.asp"
					end if
				end if
				
			case 3	' busca el último pago para la tarjeta
				
				sSQl = "exec MostrarUltimoPagoAMEXTarjeta " & evaluarstr(sNumeroTarjeta)
				set rs = ejecutarsqlcliente(Session("afxCnxAFEXchangeMCF"), sSQL)
				if err.number <> 0 then
					set rs = nothing
					MostrarErrorMS "Error al buscar el último PagoAMEX. "
				else
					if not rs.eof then
						iFormaPago = rs("formapago")
						iMonedaPago = rs("monedapago")
						cMontoDolares = rs("montodolares")
						cMontoPesos = "0"
						cTipoCambio = "0,00"						
						sNombreTitular = rs("nombretitular")
						sTelefonoTitular = rs("telefonotitular")
						'sFecha = rs("Fecha")
						rs.close						
					end if
					set rs = nothing					
				end if

			case 4	' graba el número de tarjeta rechazada				
				' graba el número				

				sSQL = " exec InsertartarjetaAMEXnoconsiderada " & evaluarstr(sNumeroTarjeta) & ", " & evaluarstr(Session("CodigoAgente"))
				set rs = ejecutarsqlcliente(Session("afxCnxAFEXchangeMCF"), sSQL)
				if err.number <> 0 then
					set rs = nothing
					MostrarErrorMS "Error al grabar el número. "
				else			
					set rs = nothing
					'Response.Redirect "AtencionClientes.asp"
				end if

		end select
	
	end if
		If sFecha = empty then
			sFecha = Date
		End If
	function VerificarExistenciaPago()
		dim sSQL, rs, swExiste
	
		VerificarExistenciaPago = True
	
		' verfiica si ya existe un pago con el mismo monto para la tarjeta ingresada el día de hoy				
		sSQl = "exec VerificarExistenciaPagoAMEX " & evaluarstr(sNumeroTarjeta) & ", " & formatonumerosql(ccur(Request.Form("txtMontoDolares")))
		set rs = ejecutarsqlcliente(Session("afxCnxAFEXchangeMCF"), sSQL)
		if err.number <> 0 then
			set rs = nothing
			MostrarErrorMS "Verificarexistencia."
		else
			if not rs.eof then
				swExiste = rs("swExistenciaPagoAMEX")				
				set rs = nothing			
				
				if swExiste = "1" then
					sMensajeInicio = "Ya existe un pago con el Monto ingresado, debe ingresar otro."
					cMontoDolares = ""
					cMontoPesos = ""
					cTipoCambio = ""
					
					exit function
				end if
			end if
		end if
					
		VerificarExistenciaPago = False	
	end function
	
	sNombreSucursal = Session("NombreCliente")
	
%>
<HTML>

	<head>
		<title>Pago AMEX</title>
	</head>	
	
	<BODY>	
			
		<!--
		<OBJECT id="ImpresoraTM295" style="LEFT: 0px; TOP: 0px; HIGHT: 0px; WIDTH: 0px;"  codebase="afxImpresorComprobante.CAB#version=1,0,0,0" 
			classid=CLSID:922A67FD-86CE-41E3-97FB-B85668474D0E VIEWASTEXT>
			<PARAM NAME="_ExtentX" VALUE="1958">
			<PARAM NAME="_ExtentY" VALUE="1640"></OBJECT>	
		-->
	
		<script language="VBScript">
		<!--			
			sub window_onload()			
				InicializarObjetos
				
				' verifica la accion
				if "<%=request("Accion")%>" = "2" and "<%=sMensajeInicio%>" = "" then
					' se grabó el pago asi es que ahora se manda a imprimir....
					cmdImprimir_onClick()
				end if
				
				' verifica si se muestra un mensaje al usuario
				if "<%=sMensajeInicio%>" <> "" then
					msgbox "<%=sMensajeInicio%>", ,"AFEX"
				end if
				
				' JFMG 20-07-2009 				
				if "<%=request("Accion")%>" = "4" then
					frmAMEX.txtNumeroTarjeta.value = ""
				end if
				' ********* FIN JFMG 20-07-2009 ***************
'/************** pss 12-08-2010
			'	If "<%=request("Accion")%>" = "5" then
			'		frmAMEX.txtFecha.value = ""
			'	End If

			end sub
			
			sub cmdImprimir_onClick()
				dim Printer
					
				On Error Resume Next
							
				If CajaPregunta("AFEX En Linea", "Coloque el comprobante AMEX en la impresora TM y haga click en Aceptar") Then					
					set Printer = CreateObject("afexImpresorComprobante.TM295")
					
					'ImpresoraTM295.ImprimirComprobanteAMEX 
					Printer.ImprimirComprobanteAMEX frmAMEX.txtNumeroTarjeta.value, "<%=iFormaPago%>", "<%=iMonedaPago%>", frmAMEX.txtMontoDolares.value, _
														frmAMEX.txtMontoPesos.value, frmAMEX.txtTipoCambio.value, frmAMEX.txtNombreSucursal.value, _
														frmAMEX.txtNombreTitular.value, frmAMEX.txtTelefonoTitular.value, "<%=sfecha%>"
				
					if err.number <> 0 then
						msgbox "Ocurrió un error al imprimir el comprobante. " & err.Description, , "AFEX"
					end if
					
					set Printer = nothing
					frmAMEX.txtfecha.value= date()
				End If
					
				if "<%=request("Accion")%>" = "1" then exit sub	
				
				frmAMEX.txtNumeroTarjeta.value = ""
				frmAMEX.txtMontoDolares.value = formatnumber(ccur("0"), 2)
				frmAMEX.txtMontoPesos.value = formatnumber(ccur("0"), 0)
				frmAMEX.txtTipoCambio.value = formatnumber(ccur("0"), 2)
				frmAMEX.txtNombreSucursal.value = "<%=sNombreSucursal%>"
				frmAMEX.txtNombreTitular.value = ""
				frmAMEX.txtTelefonoTitular.value = ""
				frmAMEX.txtFecha.Value = "<%=sFecha%>"
			end sub
			
			sub InicializarObjetos()
				
				select case cint("0" & "<%=iFormaPago%>")
					case 1	' efectivo
						rbtEfectivo_onClick()
						
					case 2	' cheque
						rbtCheque_onClick()
						
					case else
						rbtEfectivo_onClick()				
				end select
				
				select case cint("0" & "<%=iMonedaPago%>")
					case 1	' Dolar
						rbtDolar_onClick()
						
					case 2	' Peso
						rbtPeso_onClick()
						
					case else
						rbtDolar_onClick()				
				end select
				
				frmAMEX.txtMontoDolares.value = formatnumber(ccur("0" & "<%=cMontoDolares%>"), 2)
				frmAMEX.txtMontoPesos.value = formatnumber(ccur("0" & "<%=cMontoPesos%>"), 0)
				frmAMEX.txtTipoCambio.value = formatnumber(ccur("0" & "<%=cTipoCambio%>"), 4)
				frmAMEX.txtNombreSucursal.value = "<%=sNombreSucursal%>"
				frmAMEX.txtNombreTitular.value = "<%=sNombreTitular%>"
				frmAMEX.txtTelefonoTitular.value = "<%=sTelefonoTitular%>"		

				if frmAMEX.txtNumeroTarjeta.value = "" then
					frmAMEX.txtNumeroTarjeta.focus
					frmAMEX.txtNumeroTarjeta.select
				
				elseif ccur("0" & frmAMEX.txtMontoDolares.value) = 0 and not frmAMEX.txtMontoDolares.disabled then
					frmAMEX.txtMontoDolares.focus
					frmAMEX.txtMontoDolares.select
					
				elseif ccur("0" & frmAMEX.txtMontoPesos.value) = 0 and not frmAMEX.txtMontoPesos.disabled then
					frmAMEX.txtMontoPesos.focus
					frmAMEX.txtMontoPesos.select
					
				elseif ccur("0" & frmAMEX.txtTipoCambio.value) = 0 and not frmAMEX.txtTipoCambio.disabled then
					frmAMEX.txtTipoCambio.focus
					frmAMEX.txtTipoCambio.select
					
				elseif frmAMEX.txtNombreSucursal.value = "" then
					frmAMEX.txtNombreSucursal.focus
					frmAMEX.txtNombreSucursal.select
					
				elseif frmAMEX.txtNombreTitular.value = "" then
					frmAMEX.txtNombreTitular.focus
					frmAMEX.txtNombreTitular.select
				
				elseif frmAMEX.txtTelefonoTitular.value = "" then
					frmAMEX.txtTelefonoTitular.focus
					frmAMEX.txtTelefonoTitular.select
				end if
				
				if "<%=request("Accion")%>" = "1" then	' detalle de pago
					frmAMEX.txtNumeroTarjeta.value = "<%=sNumeroTarjeta%>"
				end if
			frmAMEX.txtFecha.value = "<%=sFecha%>"
			end sub
				
			sub txtMontoDolares_onBlur()
				CalcularDolares				
			end sub
			sub txtMontoPesos_onBlur()				
				CalcularDolares				
			end sub
			sub txtTipoCambio_onBlur()				
				CalcularDolares				
			end sub
				
			sub rbtEfectivo_onClick()
				frmAMEX.rbtCheque.checked = false
				frmAMEX.rbtCheque.value = ""				
				
				frmAMEX.rbtEfectivo.value = "on"	
				frmAMEX.rbtEfectivo.checked = true
			end sub
			
			sub rbtCheque_onClick()
				frmAMEX.rbtEfectivo.checked = false
				frmAMEX.rbtEfectivo.value = ""				
				
				frmAMEX.rbtCheque.value = "on"
				frmAMEX.rbtCheque.checked = true
			end sub
			
			sub rbtDolar_onClick()
				frmAMEX.rbtPeso.checked = false
				frmAMEX.rbtPeso.value = ""				
				
				frmAMEX.rbtDolar.value = "on"	
				frmAMEX.rbtDolar.checked = true
				
				CalcularDolares	
			end sub
			
			sub rbtPeso_onClick()
				frmAMEX.rbtDolar.checked = false
				frmAMEX.rbtDolar.value = ""
				
				frmAMEX.rbtPeso.value = "on"
				frmAMEX.rbtPeso.checked = true
				
				CalcularDolares
			end sub
			
			sub cmdAceptar_onClick()
				' validar el número de tarjeta
				if not cardValidation("37" & trim(frmAMEX.txtNumeroTarjeta.value)) then
					msgbox "El número de tarjeta no es correcto.", , "AFEX"
					frmAMEX.txtNumeroTarjeta.value = ""
					frmAMEX.txtNumeroTarjeta.focus
					frmAMEX.txtNumeroTarjeta.select
					exit sub
				end if
			
				' valida los demás campos
				if not ValidarCampos then exit sub			
			
				frmAMEX.action = "AgregarPagoAMEX.asp?Accion=2"
				frmAMEX.submit()
				frmAMEX.action = ""
			end sub
			
			
	Sub txtFecha_onBlur()
			
				if "<%=request("Accion")%>" = "1" then exit sub
				
				if triM(frmAMEX.txtFecha.value) = "" then exit sub
				
				IF NOT ISDATE(frmAMEX.txtFecha.value) Then
					msgbox "La fecha ingresada no es correcta, ingrese nuevamente.", ,"AFEX" 	
					frmAMEX.txtFecha.value =""
					frmAMEX.action = "AgregarPagoAMEX.asp?Accion=5"
					frmAMEX.submit()
					frmAMEX.action = ""
					
				End If
			End Sub
			
			
			
			sub txtNumeroTarjeta_onBlur()
				if "<%=request("Accion")%>" = "1" then exit sub
			
				if triM(frmAMEX.txtNumeroTarjeta.value) = "" then exit sub
				
				' validar el número de tarjeta
				if not cardValidation("37" & trim(frmAMEX.txtNumeroTarjeta.value)) then
					msgbox "El número de tarjeta no es correcto.", , "AFEX"

					
					' JFMG 20-07-2009 se agrega proc. para guardar el número de tarjeta rechazada
					'frmAMEX.txtNumeroTarjeta.value = ""
					'frmAMEX.txtNumeroTarjeta.focus
					'frmAMEX.txtNumeroTarjeta.select

					frmAMEX.action = "AgregarPagoAMEX.asp?Accion=4"
					frmAMEX.submit()
					frmAMEX.action = ""
					' *********** FIN JFMG 20-07-2009 *********************

					exit sub
				else
					
					frmAMEX.action = "AgregarPagoAMEX.asp?Accion=3"
					frmAMEX.submit()
					frmAMEX.action = ""
				end if			
				
			end sub
			
			function ValidarCampos
				ValidarCampos = False
			
				if frmAMEX.rbtPeso.checked then
					if ccur("0" & frmAMEX.txtMontoPesos.value) = 0 then
						msgbox "Debe ingresar el monto en Pesos.", , "AFEX"
						frmAMEX.txtMontoPesos.focus
						frmAMEX.txtMontoPesos.select
						exit function
					end if
								
					if ccur("0" & frmAMEX.txtTipoCambio.value) = 0 then
						msgbox "Debe ingresar el Tipo de Cambio.", , "AFEX"
						frmAMEX.txtTipoCambio.focus
						frmAMEX.txtTipoCambio.select
						exit function
					end if
				else
					if ccur("0" & frmAMEX.txtMontoDolares.value) = 0 then
						msgbox "Debe ingresar el monto en Dolares.", , "AFEX"
						frmAMEX.txtMontoDolares.focus
						frmAMEX.txtMontoDolares.select
						exit function
					end if
				end if
				
				if ccur("0" & frmAMEX.txtMontoDolares.value) > 100000 then
					msgbox "El monto máximo de pago no puede exceder US$ 100.000", , "AFEX"
					frmAMEX.txtMontoDolares.focus
					frmAMEX.txtMontoDolares.select
					exit function
				end if				
				
				if frmAMEX.txtNombreSucursal.value = "" then
					msgbox "Debe ingresar la sucursal.", , "AFEX"
					frmAMEX.txtNombreSucursal.focus
					frmAMEX.txtNombreSucursal.select
					exit function
				end if
				if trim(frmAMEX.txtNombreTitular.value) = "" then
					msgbox "Debe ingresar el nombre del titular.", , "AFEX"
					frmAMEX.txtNombreTitular.focus
					frmAMEX.txtNombreTitular.select
					exit function
				end if
				if trim(frmAMEX.txtTelefonoTitular.value) = "" then
					msgbox "Debe ingresar el teléfono del titular.", , "AFEX"
					frmAMEX.txtTelefonoTitular.focus
					frmAMEX.txtTelefonoTitular.select
					exit function
				end if
				
   			If trim(frmAMEX.txtFecha.value ) = "" then
					msgbox "Debe ingresar la fecha.", , "AFEX"
					frmAMEX.txtFecha.focus
					'txtAMEX.txtFecha.select
					exit function
				end if
				frmAMEX.txtMontoDolares.disabled = false
				frmAMEX.txtMontoPesos.disabled = false
				frmAMEX.txtTipoCambio.disabled = false
				
				ValidarCampos = True
			end function
			
			sub CalcularDolares()
				frmAMEX.txtMontoDolares.value = formatnumber(ccur("0" & frmAMEX.txtMontoDolares.value),2)			
				frmAMEX.txtMontoPesos.value = formatnumber(ccur("0" & frmAMEX.txtMontoPesos.value),0)
				frmAMEX.txtTipoCambio.value = formatnumber(ccur("0" & frmAMEX.txtTipoCambio.value),4)
				
				
				if frmAMEX.rbtPeso.checked then
					if ccur(frmAMEX.txtMontoPesos.value) <> 0 and ccur(frmAMEX.txtTipoCambio.value) <> 0 then
						frmAMEX.txtMontoDolares.value = formatnumber(round(frmAMEX.txtMontoPesos.value / frmAMEX.txtTipoCambio.value,2),2)
					end if
					
				else
					if ccur(frmAMEX.txtTipoCambio.value) <> 0  and ccur(frmAMEX.txtMontoDolares.value) <> 0 then
						frmAMEX.txtMontoPesos.value = formatnumber(round(frmAMEX.txtMontoDolares.value * frmAMEX.txtTipoCambio.value,0),0)
					end if
				end if
			end sub
			
			' CÓDIGO ENVIADO POR AMEX
			Public Function Mod10(vsAccount)
			' RETURNS:
			'   TRUE    -> If card is VALID.
			'   FALSE   -> If card is INVALID.
			    
			    '
			    ' This function returns true if the input
			    ' account number passes the Mod10 check.
			    '
			    Dim mnLoopCnt, mnFinalNum, mnTot
			    '
			    ' Make sure we have a full account number
			    '
			    If Len(Trim(vsAccount)) <> 15 Then
			       Mod10 = False
			       Exit Function
			    End If
			    
			    If vsAccount = "000000000000000" Then
			       Mod10 = False
			       Exit Function
			    End If
			    
			    '
			    ' Loop thru all 15 digits in account number
			    '
			    For mnLoopCnt = 1 To Len(vsAccount)
			       '
			       ' If mnLoopCnt is an even number double the number
			       '
			       If mnLoopCnt Mod 2 = 0 Then
			          mnFinalNum = CInt(Mid(vsAccount, mnLoopCnt, 1)) * 2
			          '
			          ' If 10 or greater, mod 10 plus 1
			          '
			          If mnFinalNum >= 10 Then
			             mnFinalNum = (mnFinalNum Mod 10) + 1
			          End If
			       Else     ' Otherwise just pass the exsisting number
			          mnFinalNum = CInt(Mid(vsAccount, mnLoopCnt, 1))
			        End If
			        mnTot = mnTot + mnFinalNum     ' Running total of result
			    Next
			    '
			    ' Final mod 10 check of totaled result
			    '
			    If mnTot Mod 10 <> 0 Then
			       Mod10 = False
			    Else
			       Mod10 = True
			    End If
			    
			End Function

			Public Function cardValidation(cardNumber)
			' RETURNS:
			'   TRUE    -> If card is CMRS VALID.
			'   FALSE   -> If card is CMRS INVALID.

			   
			    ' Check for Mod10
			    If Not Mod10(cardNumber) Then
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Not Mid(cardNumber, 1, 2) = "37" Then
			        ' This is not a valid American Express Cardnumber
			        cardValidation = False
			        Exit Function
			     End If
			    
			    If Mid(cardNumber, 1, 4) = "3732" Then
			        If Mid(cardNumber, 5, 1) = "6" Then
			           ' This is a Canadian Card, please use Foreign Form
			           cardValidation = False
			           Exit Function
			        End If
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3707" Then
			        ' This is a Mexico Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3733" Then
			        ' This is a Canadian Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3735" Then
			        ' This is a Canadian Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3740" Then
			        ' This is an Austria Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3741" Then
			        ' This is a Belgium Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3742" Then
			        ' This is a U.K. Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3743" Then
			        ' This is an Ireland Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3744" Then
			        ' This is a Bahrain Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3745" Then
			        ' This is a U.K. Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3746" Then
			        ' This is a U.K./France Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3747" Then
			        ' This is a Denmark Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			     
			    If Mid(cardNumber, 1, 4) = "3748" Then
			        ' This is a Finland Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3749" Then
			        ' This is a France Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3750" Then
			        ' This is a Germany Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3752" Then
			        ' This is an Italy Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3753" Then
			        ' This is a Netherlands Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3754" Then
			        ' This is a Norway Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3755" Then
			        ' This is an Israel Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3756" Then
			        ' This is a Spain Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3757" Then
			        ' This is a Sweden Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3758" Then
			        ' This is a Switzerland Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3760" Then
			        ' This is an Australia Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			     
			    If Mid(cardNumber, 1, 4) = "3761" Then
			        ' This is a Japan Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3762" Then
			        ' This is an Asia Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3763" Then
			        ' This is an Asia Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3764" Then
			        ' This is an Argentina/Brazil Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3766" Then
			        ' This is a Brazil/Mexico Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3768" Then
			        ' This is a South Africa Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3769" Then
			        ' This is an Asia Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3770" Then
			        ' This is a Venezuela Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3774" Then
			        ' This is a New Zealand Card, please use Foreign Form
			        cardValidation = False
			        Exit Function
			    End If
			    
			    If Mid(cardNumber, 1, 4) = "3775" Then
			        If Mid(cardNumber, 5, 1) = "0" Then
			           ' This is a Croatia Card, please use Foreign Form
			           cardValidation = False
			           Exit Function
			        Else
			           ' This is a Yugoslavia Card, please use Foreign Form
			           cardValidation = False
			           Exit Function
			        End If
			    End If
			    
			    
			    ' JFMG 05-12-2008 validación agregada por AFEX, no considerarda por AMEX			    
				If instr("37158;37159;37169;37265;37268;37269;37868;37879;37904", left(cardNumber, 5)) = 0 Then
			        ' estos números de tarjeta no pueden ser cancelados por AFEX
			        cardValidation = False
			        Exit Function
			    End If
			    ' ******************************** FIN *********************************
			    
			    ' The card fullfils all requirements.
			    cardValidation = True
			    
			End Function
			' **************************** FIN ************************************
				
		-->
		</script>

<STYLE type=text/css>
			<!--
			.EstiloTexto {
				font-family: Tahoma;
				font-size: 12px;
				border-bottom: 1px solid #FFFFFF;
				border-top: 1px solid #FFFFFF;
				border-left: 1px solid #FFFFFF;
				border-right: 1px solid #FFFFFF;				
			}
			
			.EstiloTabla {
				font-family: Tahoma;			
				border-bottom: 1px solid #CCCCCC;
				border-top: 1px solid #CCCCCC;
				border-left: 1px solid #CCCCCC;
				border-right: 1px solid #CCCCCC;
			}
			-->
		</STYLE>
		
		<form name="frmAMEX" method="post" action="">
			<table border=0 align="center" style="WIDTH: 390px; BACKGROUND-REPEAT: no-repeat; HEIGHT: 653px" background="../Img/ComprobantePagoAMEX.jpg">
				<tr style="height: 170px">
					<td valign="bottom">
						<table border=0 style="WIDTH: 353px; HEIGHT: 100%">
							<tr>
								<td width="35%" ></td>
								<td width="65%" valign="bottom" align=left>
									<INPUT maxlength="13" onkeypress="IngresarTexto(1)" class=EstiloTexto id=text1 name=txtNumeroTarjeta value="<%=Request.Form("txtNumeroTarjeta")%>" style="WIDTH: 120px; HEIGHT: 18px; TEXT-ALIGN: left">
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr style="height: 430px">
					<td valign="top">
						<table border=0 style="WIDTH: 353px;">
							<tr style="height: 20px">
								<td align="right" valign="top">
									<INPUT type="radio" name=rbtEfectivo>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT type=radio name=rbtCheque>&nbsp;&nbsp;&nbsp;&nbsp;
								</td>
							</tr>
							<tr style="height: 25px">
								<td align="right" valign="top">									
									<INPUT type="radio" name=rbtDolar>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT id=radio1 style="WIDTH: 16px;" type=radio size=16 name=rbtPeso>&nbsp;&nbsp;&nbsp;&nbsp;<br>
								</td>
							</tr>
							<tr style="height: 30px">
								<td align="right" valign="center">	
									<INPUT class=EstiloTexto name=txtMontoDolares onkeypress="IngresarTexto(1)" maxlength="8" value="1" style="WIDTH: 180px; HEIGHT: 18px; TEXT-ALIGN: right;"><br>
								</td>
							</tr>
							<tr style="height: 25px">
								<td align="right" valign="top">	
									<INPUT class=EstiloTexto name=txtMontoPesos onkeypress="IngresarTexto(1)" maxlength="10" value="1" style="WIDTH: 180px; HEIGHT: 18px; TEXT-ALIGN: right;"><br>
								</td>
							</tr>
							<tr style="height: 25px">
								<td align="right" valign="bottom">	
									<INPUT class=EstiloTexto name=txtTipoCambio onkeypress="IngresarTexto(1)" maxlength="9" value="1" style="WIDTH: 200px; HEIGHT: 18px; TEXT-ALIGN: right;"><br>
								</td>
							</tr>
							<tr style="height: 30px">
								<td align="right" valign="bottom">	
									<INPUT class=EstiloTexto name=txtNombreSucursal onkeypress="IngresarTexto(2)" maxlength="40" value="1" style="WIDTH: 180px; HEIGHT: 18px; TEXT-ALIGN: left;"><br>
								</td>
							</tr>
							<tr style="height: 25px">
								<td align="right" valign="bottom">&nbsp;</td>
							</tr>
							<tr style="height: 30px">
								<td align="right" valign="bottom">										
									<INPUT class=EstiloTexto name=txtNombreTitular onkeypress="IngresarTexto(2)" maxlength="100" value="1" style="WIDTH: 250px; HEIGHT: 18px; TEXT-ALIGN: left;">
								</td>
							</tr>
							<tr style="height: 40px">
								<td align="right" valign="bottom">	
									<INPUT class=EstiloTexto name=txtTelefonoTitular maxlength="20" value=1 style="WIDTH: 250px; HEIGHT: 18px; TEXT-ALIGN: left"> 
								</td>
							</tr>
							<tr style="height: 40px; font-family: Tahoma; font-size: 12px;">
								<td align="center" valign="bottom">
									<INPUT class = EstiloTexto name =txtFecha value="<%=sFecha%>" maxlength="20" 
                                        value =1 style="WIDTH: 140px; HEIGHT: 22px">
									
								</td>
							</tr>	
							<tr style="height: 40px">
								<td align="right" valign="bottom">&nbsp;</td>
							</tr>		
							<tr style="height: 100px">
								<td align="right" valign="bottom">&nbsp;</td>			
							</tr>
							
						</table>
					</td>
				</tr>
				<tr>
					<%if request("Accion") <> "1" then	' detalle de pago
					%>					
						<td align="right"><input type="button" name="cmdAceptar" value="Aceptar"></td>
					<%end if%>
				</tr>
			</table>
		</form>
		<%if request("Accion") = "1" then %>
			<table align="center" style="WIDTH: 390px;">
				<tr>
					<td align="right"><input type="button" name="cmdImprimir" value="Imprimir"></td>
				</tr>
			</table>
		<%end if%>
	</BODY>
</HTML>