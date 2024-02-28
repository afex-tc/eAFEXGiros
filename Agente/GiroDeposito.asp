<%@  language="VBScript" %>
<!-- #INCLUDE Virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/agente/Constantes.asp" -->
<%
Dim sBanco, sTipo, sCuenta, sPagador, sFormaPago
Dim nAccion
'INTERNO-8479 MS 11-11-2016 
Dim sPais, sMonedaPago
Dim sTelefonoB, sDireccionB

	sBanco = Request.Form("cbxbanco")
	sCuenta = Request.Form("cbxTipoCta")
	sPagador = Request("Pagador")
    sFormaPago = Request("FormaPago")
    sPais = Request("Pais")
    sMonedaPago = Request("MonedaPago")
    sTelefonoB = Request("TelefonoB")
    sDireccionB = Request("DireccionB")


	'INTERNO-3855 MS 26-04-2015
	dim maxLen9,maxLen10,maxLen11,maxLen12,maxLen15, nMaxleng
	dim mensaje1, mensaje2, mensaje3
  
    if sPagador = "FP" then
        nMaxleng = 15
    elseif sPagador = "SW" then
        nMaxleng = 24
    end if
    'FIN INTERNO-3855 MS 26-04-2015
	'FIN INTERNO-8479 MS 11-11-2016 
%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <% If sPagador = "SW" Then  %>
        <title>Smallworld-Choice</title>
    <% else  %>
        <title>Giros Deposito</title>
    <% End If  %>
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css" />
</head>

<script language="VBScript">
<!--

	window.dialogWidth = 40
	window.dialogHeight = 16
	window.dialogLeft = 240
	window.dialogTop = 220
	window.defaultstatus = ""

    'INTERNO-8479 MS 11-11-2016 
	sub window_onload()
        If "<%=sPagador%>"  = "SW"  and "<%=sPais%>"  = "DO" Then 
            if  "<%=sFormaPago%>"  = "0" then
    	        frmGiroDeposito.txtMensaje.value = "Operaciones de USD 1000 o más" & chr(13) & "debe presentar los documentos" & chr(13) & "de identificación "
                frmGiroDeposito.txtMensaje.style.textAlign = "right"
            end if
        end if
    end sub
	'FIN INTERNO-8479 MS 11-11-2016 

	Sub imgAceptar_onClick()

        If "<%=sPagador%>"  <> "SW" Then
		    If frmGiroDeposito.txtNumeroCta.value = "" Then
			    msgbox "Debe Ingresar el número de cuenta",,"AFEX"
			    frmGiroDeposito.txtNumeroCta.focus
			    Exit sub
		    End If
        'miki SMC-9 2016-03-02 FIX2
        else
			'INTERNO-8479 MS 11-11-2016 
            if validarSW() = false then
                exit sub
            end if
			'FIN INTERNO-8479 MS 11-11-2016 
        End if
		If "<%=sPagador%>"  = "BC" Then  
			If frmGiroDeposito.cbxTipoCta.value ="AH" Then
				If len(frmGiroDeposito.txtNumeroCta.value ) <> 14 Then
					msgbox "EL tamaño de la cuenta no corresponde",,"AFEX" 
					frmGiroDeposito.txtNumeroCta.focus
					frmGiroDeposito.txtNumeroCta.select
					Exit sub
				End If
			End IF
			If frmGiroDeposito.cbxTipoCta.value ="CM" or frmGiroDeposito.cbxTipoCta.value = "CC" Then
					If len(frmGiroDeposito.txtNumeroCta.value) <> 13 Then
						msgbox "EL tamaño de la cuenta no corresponde",,"AFEX" 
						frmGiroDeposito.txtNumeroCta.focus
						frmGiroDeposito.txtNumeroCta.select
						Exit Sub
					End If
				End If
		End If
		If "<%=sPagador%>"  = "IK" Then  
			If frmGiroDeposito.cbxTipoCta.value ="T" Then
				frmGiroDeposito.txtValorcuenta.value = replace(frmGiroDeposito.txtNumeroCta.value ,"-","")
				If trim(len(frmGiroDeposito.txtvalorcuenta.value)) <> 16 Then	
					msgbox "El largo del número de tarjeta debe ser 16 caracteres, sin considerar puntos ni guiones",,"AFEX"
					frmGiroDeposito.txtNumeroCta.select 
					frmGiroDeposito.txtNumeroCta.focus 
					Exit Sub
				End if
			End if
			If frmGiroDeposito.cbxTipoCta.value = "A" Then
				frmGiroDeposito.txtValorcuenta.value = replace(frmGiroDeposito.txtNumeroCta.value ,"-","")
				If trim(len(frmGiroDeposito.txtvalorcuenta.value)) <> 13 Then	
					msgbox "El largo del número de cuenta debe ser 13 caracteres, incluyendo guiones",,"AFEX"
					frmGiroDeposito.txtNumeroCta.select 
					frmGiroDeposito.txtNumeroCta.focus 
					Exit Sub
				End if
			End if  
		End if
		'INTERNO-3855 MS 26-04-2015
		If "<%=sPagador%>"  = "FP" Then
            maxLen9 = ";2;4;5;7;10;20;28;"
            maxLen10= ";8;12;19;"
            maxLen11= ";1;14;"
            maxLen12= ";3;6;9;11;13;15;23;"
            maxLen15= ";16;17;18;21;22;24;25;26;27;29;"

            frmGiroDeposito.txtValorcuenta.value = replace(frmGiroDeposito.txtNumeroCta.value ,"-","")
            frmGiroDeposito.txtValorcuenta.value = replace(frmGiroDeposito.txtNumeroCta.value ,".","")

            If  instr(1, maxLen9, ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
                If len(trim(frmGiroDeposito.txtvalorcuenta.value)) <> 12 Then	
                    msgbox "El largo del número de cuenta debe ser de 12 caracteres, sin considerar puntos ni guiones",,"AFEX"
                    frmGiroDeposito.txtNumeroCta.select 
                    frmGiroDeposito.txtNumeroCta.focus 
                    Exit Sub
                End if
            End if

            If  instr(1, maxLen10,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
                If len(trim(frmGiroDeposito.txtvalorcuenta.value)) <> 10 Then	
                    msgbox "El largo del número de cuenta debe ser de 10 caracteres, sin considerar puntos ni guiones",,"AFEX"
                    frmGiroDeposito.txtNumeroCta.select 
                    frmGiroDeposito.txtNumeroCta.focus 
                    Exit Sub
                End if
            End if

            If  instr(1, maxLen11,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
                If len(trim(frmGiroDeposito.txtvalorcuenta.value)) <> 11 Then	
                    msgbox "El largo del número de cuenta debe ser de 11 caracteres, sin considerar puntos ni guiones",,"AFEX"
                    frmGiroDeposito.txtNumeroCta.select 
                    frmGiroDeposito.txtNumeroCta.focus 
                    Exit Sub
                else
                    if frmGiroDeposito.cbxbanco.value = 14 and Mid(frmGiroDeposito.txtvalorcuenta.value,1,2) <> "00" then
                        msgbox "Número de cuenta errónea. La cuenta debe comenzar con los dígitos 00",,"AFEX"
                        frmGiroDeposito.txtNumeroCta.select 
                        frmGiroDeposito.txtNumeroCta.focus 
                        Exit Sub
                    end if
                End if
            End if

            If  instr(1, maxLen12, ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
              if frmGiroDeposito.cbxbanco.value = 6 then 'BANCO AGRARIO
                 if len(trim(frmGiroDeposito.txtvalorcuenta.value)) < 11 then
                   msgbox "El largo del número de cuenta debe tener 11 caracteres mínimo, sin considerar puntos ni guiones.",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub
                 elseif len(trim(frmGiroDeposito.txtvalorcuenta.value)) > 12 then
                   msgbox "El largo del número de cuenta debe tener 12 caracteres máximo, sin considerar puntos ni guiones.",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub
                 end if
              elseif frmGiroDeposito.cbxbanco.value = 9 then 'BANCO SUDAMERIS
                 if len(trim(frmGiroDeposito.txtvalorcuenta.value))= 10 then
                   msgbox "El tamaño de la cuenta no corresponde a la del banco.",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub
                 elseif len(trim(frmGiroDeposito.txtvalorcuenta.value)) < 8 then
                   msgbox "El tamaño de la cuenta debe tener 8 caracteres mínimo, sin considerar puntos ni guiones.",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub
                 elseif trim(len(frmGiroDeposito.txtvalorcuenta.value)) > 12 then
                   msgbox "El tamaño de la cuenta debe tener 12 caracteres máximo, sin considerar puntos ni guiones.",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub                                
                 end if
              elseif frmGiroDeposito.cbxbanco.value = 11 then 'BANCO CAJA SOCIAL
                 if len(trim(frmGiroDeposito.txtvalorcuenta.value))< 11 then 
                   msgbox "El tamaño de la cuenta debe tener 11 caracteres mínimo, sin considerar puntos ni guiones.",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub  
                 elseif len(trim(frmGiroDeposito.txtvalorcuenta.value)) > 12 then                                  
                   msgbox "El tamaño de la cuenta debe tener 12 caracteres máximo, sin considerar puntos ni guiones.",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub  
                 end if
              elseif frmGiroDeposito.cbxbanco.value = 3 then 'DAVIVIENDA
                 if len(trim(frmGiroDeposito.txtvalorcuenta.value)) = 11  then
                   msgbox "El tamaño de la cuenta no corresponde a la del banco",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub
                 elseif len(trim(frmGiroDeposito.txtvalorcuenta.value)) < 9 then
                   msgbox "El largo del número de cuenta debe tener 9 caracteres mínimo, sin considerar puntos ni guiones",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub
                 elseif len(trim(frmGiroDeposito.txtvalorcuenta.value)) > 12 then
                   msgbox "El largo del número de cuenta debe tener 12 caracteres máximo, sin considerar puntos ni guiones",,"AFEX"
                   frmGiroDeposito.txtNumeroCta.select 
                   frmGiroDeposito.txtNumeroCta.focus 
                   Exit Sub
                 end if
              elseif len(trim(frmGiroDeposito.txtvalorcuenta.value)) <> 12  then 'cualquier otro pagador
                  msgbox "El largo del número de cuenta debe ser de 12 caracteres, sin considerar puntos ni guiones. otros agentes",,"AFEX"
                  frmGiroDeposito.txtNumeroCta.select 
                  frmGiroDeposito.txtNumeroCta.focus 
                  Exit Sub
              end if
              'APPL-14657 MS 15-05-2015
            End if 

            If  instr(1, maxLen15,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
                frmGiroDeposito.txtValorcuenta.value = replace(frmGiroDeposito.txtNumeroCta.value ,"-","")
                If len(trim(frmGiroDeposito.txtvalorcuenta.value)) <> 15 Then	
                    'HELM BANK
                    if frmGiroDeposito.cbxbanco.value = 10 then
                        if len(trim(frmGiroDeposito.txtvalorcuenta.value)) <> 9 then
                            msgbox "EL tamaño de la cuenta no corresponde.",,"AFEX"
                            frmGiroDeposito.txtNumeroCta.select 
                            frmGiroDeposito.txtNumeroCta.focus 
                            Exit Sub
                        end if
                    else
                        msgbox "El largo del número de cuenta debe ser de  máximo 15 caracteres, sin considerar puntos ni guiones",,"AFEX"
                        frmGiroDeposito.txtNumeroCta.select 
                        frmGiroDeposito.txtNumeroCta.focus 
                        Exit Sub
                    end if
                End if
            End if
            'FIN APPL-14657 MS 15-05-2015
            
            If frmGiroDeposito.cbxTipoCta.value = "" and instr(1, ";13;14;15;",  ";" & frmGiroDeposito.cbxbanco.value & ";") <= 0 Then
                msgbox "Debe seleccionar el tipo de cuenta",,"AFEX" 
                frmGiroDeposito.cbxTipoCta.focus
                Exit sub
            End IF

            frmGiroDeposito.cbxMonedaDeposito.value = "MNL"	
		End if
        'miki SMC-9 MM 2015-11-30
		'INTERNO-8479 MS 11-11-2016 
        If "<%=sPagador%>"  <> "SW" Then
            'FIN INTERNO-3855 MS 26-04-2015
		    window.returnvalue = frmGiroDeposito.cbxBanco.value & ";" & frmGiroDeposito.cbxTipoCta.value & ";" & _
							     frmGiroDeposito.txtNumeroCta.value & ";" & frmGiroDeposito.cbxMonedaDeposito.value & ";;"
							 
        Else
            If "<%=sPais %>" = "HT" or "<%=sPais %>" = "DO" then
                if "<%=sFormaPago %>" = "2" then
                    
                    window.returnvalue = frmGiroDeposito.cbxBanco.value & ";" & frmGiroDeposito.cbxTipoCta.value & ";" & _
							 frmGiroDeposito.txtNumeroCta.value & ";<%=sMonedaPago%>;" & frmGiroDeposito.txtDireccionB.value & ";"  & _
                             frmGiroDeposito.txtTelefonoB.value
                else 
                    if "<%=sFormaPago %>" = "0" then
                        window.returnvalue = frmGiroDeposito.cbxBanco.value & ";" & frmGiroDeposito.cbxTipoCta.value & ";" & _
							     frmGiroDeposito.txtNumeroCta.value & ";<%=sMonedaPago%>" & ";;"
                    else
                        window.returnvalue = frmGiroDeposito.cbxAgenciaBanco.value & ";" & frmGiroDeposito.cbxTipoCta.value & ";" & _
							     frmGiroDeposito.txtNumeroCta.value & ";<%=sMonedaPago%>" & ";;" 
                    End if
                end if
            else
                window.returnvalue = frmGiroDeposito.cbxBanco.value & ";" & frmGiroDeposito.cbxTipoCta.value & ";" & _
							 frmGiroDeposito.txtNumeroCta.value & ";<%=sMonedaPago%>" & ";;" 
            End if
			'FIN INTERNO-8479 MS 11-11-2016 
        End if
		window.close		
		
	End SUb
	
	Sub txtNumeroCta_onBLur()
		dim i
	
			 If "<%=sPagador%>"  = "BC" Then 
				If frmGiroDeposito.cbxTipoCta.value ="AH" Then
					If len(frmGiroDeposito.txtNumeroCta.value ) = 14 Then
						frmGiroDeposito.txtValorcuenta.value  = mid(frmGiroDeposito.txtNumeroCta.value ,12,1)
					End If
				End IF
				If frmGiroDeposito.cbxTipoCta.value ="CM" or frmGiroDeposito.cbxTipoCta.value = "CC" Then
					If len(frmGiroDeposito.txtNumeroCta.value) =13 Then
						frmGiroDeposito.txtValorcuenta.value  = mid(frmGiroDeposito.txtNumeroCta.value ,11,1)
					End If
				End If
				'If frmGiroDeposito.txtValorcuenta.value  = "1" Then
				'	frmGiroDeposito.cbxMonedaDeposito.value = "USD"
				'ElseIf frmGiroDeposito.txtValorcuenta.value = "0" Then
				'	frmGiroDeposito.cbxMonedaDeposito.value = "MNL"
				'Else
				'	msgbox "El número de la cuenta es erroneo",,"AFEX"
				'	frmGiroDeposito.txtNumeroCta.select
				'End If
			End If
						 
	End Sub
	
	'INTERNO-3855 MS 26-04-2015
	Sub cbxbanco_onChange()	    
		 If "<%=sPagador%>"  = "FP" Then 
             mensaje1 = "* INFORMAR AL CLIENTE: Se cobrará un 2% al momento del desembolso pues este banco es del Gobierno y no pertenece a la red ACH."	
             mensaje2 = "* INFORMAR AL CLIENTE: Se depositan inmediatamente de lunes a sábado en el horario que laboran las oficinas de pago. Lunes a viernes de 8.00 a 17.00 hrs y Sábado de 9.00 a 14.00 hrs."
             mensaje3 = "* INFORMAR AL CLIENTE: El giro se verá reflejado 2 horas posterior al corte de la nómina que es: 8.00 - 10.00 - 15.00. Si el giro es recaudado después de las 15.00 se reflejará al día siguiente"
             mensaje	= mensaje3
	
		     If  instr(1, maxLen9,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
		        nMaxleng = 9
		     end if
		 
		     If  instr(1, maxLen10,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
		        nMaxleng = 10
		     end if
    		 
		     If  instr(1, maxLen11,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
		        nMaxleng = 11
		     end if
    		 
		     If  instr(1, maxLen12,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
		        nMaxleng = 12
		     end if
    		 
		     If  instr(1, maxLen15,  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
		        nMaxleng = 15
		     end if
		     
	         If  instr(1, ";13;14;15;",  ";" & frmGiroDeposito.cbxbanco.value & ";") > 0  Then
	            frmGiroDeposito.cbxTipoCta.value = ""
		        frmGiroDeposito.cbxTipoCta.disabled = true
		     else
		        frmGiroDeposito.cbxTipoCta.disabled = false
		     end if
		  
		    If frmGiroDeposito.cbxbanco.value = "1" Then
			    frmGiroDeposito.txtMensaje.value = mensaje2
		    else	
		        If frmGiroDeposito.cbxbanco.value = "6" Then
			        frmGiroDeposito.txtMensaje.value = mensaje1
		        else
		            frmGiroDeposito.txtMensaje.value = mensaje3
		        End if
		    End if
		End if	
        'miki SMC-9 MM 2015-11-30 FIN
        If "<%=sPagador%>"  = "SW" Then
			'INTERNO-8479 MS 11-11-2016  
            mensaje1 = ""
            If "<%=sPais%>"  = "CO" Then 
                if frmGiroDeposito.cbxbanco.value = "1" then
    	            mensaje1 = "* INFORMAR AL CLIENTE:" & chr(13) & "Bancolombia puede recibir entre 25 y 5000 USD"	
                    frmGiroDeposito.txtNumeroCta.maxLength = 11
                else
                    mensaje1 = "* INFORMAR AL CLIENTE: " & chr(13) & " Davivienda puede recibir entre 25 y 2500 USD"	
                    frmGiroDeposito.txtNumeroCta.maxLength = 12
                end if
            end if

            If "<%=sPagador%>"  = "SW"  and "<%=sPais%>"  = "DO" Then 
                if  "<%=sFormaPago%>"  = "0" then
    	            mensaje1 = "Operaciones de USD 1000 o más" & chr(13) & "debe presentar los documentos" & chr(13) & "de identificación "                                  
                end if
            end if
			'FIN INTERNO-8479 MS 11-11-2016 
            frmGiroDeposito.txtMensaje.value = mensaje1
        end if
        frmGiroDeposito.txtNumeroCta.value = "" 'miki SMC-9 2016-03-02 FIX1
	End Sub
    'INTERNO-3855 MS 26-04-2015
	
	'INTERNO-8479 MS 11-11-2016 
    function ValidarSW()
        ValidarSW = true
        If Trim(frmGiroDeposito.cbxbanco.value) = "" Then
			msgbox "Debe seleccionar el banco destino",,"AFEX"
			frmGiroDeposito.cbxbanco.focus
			ValidarSW = false
            exit function
		End If
        
        if "<%=sFormaPago %>" = "1" then
            If Trim(frmGiroDeposito.cbxTipoCta.value) = "" Then
			    msgbox "Debe seleccionar el tipo de cuenta bancaria",,"AFEX"
			    frmGiroDeposito.cbxTipoCta.focus
			    ValidarSW = false
                exit function
		    End If
            If Trim(frmGiroDeposito.txtNumeroCta.value) = "" Then
			    msgbox "Debe Ingresar el número de cuenta bancaria",,"AFEX"
			    frmGiroDeposito.txtNumeroCta.focus
			    ValidarSW = false
                exit function
		    End If
            If ("<%=sPais%>" = "DO" or "<%=sPais%>" = "HT") Then
                If Trim(frmGiroDeposito.cbxAgenciaBanco.value) = "" Then
			        msgbox "Debe seleccionar la agencia pagadora",,"AFEX"
			        frmGiroDeposito.cbxAgenciaBanco.focus
			        ValidarSW = false
                    exit function
		        End If
            end if
        End if

        'miki SMC-9 2016-03-02 FIX2
        'miki SMC-30 2016-03-08
        'bancolombia
        If "<%=sPais%>"  = "CO" Then
            If frmGiroDeposito.cbxbanco.selectedIndex = 1 and Len(Trim(frmGiroDeposito.txtNumeroCta.value)) <> "11" and "<%=sFormaPago%>" = "1" Then
			    msgbox "La longitud del número de cuenta para Bancolombia debe ser de 11 dígitos",,"AFEX"
			    frmGiroDeposito.txtNumeroCta.focus
			    ValidarSW = false
                exit function
		    End If
            'davivienda
            If frmGiroDeposito.cbxbanco.selectedIndex = 2 and Len(Trim(frmGiroDeposito.txtNumeroCta.value)) <> "12" and "<%=sFormaPago%>" = "1" Then
			    msgbox "La longitud del número de cuenta para Davivienda debe ser de 12 dígitos",,"AFEX"
			    frmGiroDeposito.txtNumeroCta.focus
			    ValidarSW = false
                exit function
		    End If
            'FIN miki SMC-30 2016-03-08
        elseif "<%=sPais%>"  = "BO" and "<%=sFormaPago %>" = "1" then
            If Len(Trim(frmGiroDeposito.txtNumeroCta.value)) <> 10 Then
			    msgbox "La longitud del número de cuenta debe ser de 10 dígitos",,"AFEX"
			    frmGiroDeposito.txtNumeroCta.focus
			    ValidarSW = false
                exit function
		    End If
        else
            If "<%=sFormaPago %>" = "1" and Len(Trim(frmGiroDeposito.txtNumeroCta.value)) < 5 or Len(Trim(frmGiroDeposito.txtNumeroCta.value)) > 24 Then
			    msgbox "La longitud del número de cuenta debe ser entre 5 y 24 dígitos",,"AFEX"
			    frmGiroDeposito.txtNumeroCta.focus
			    ValidarSW = false
                exit function
		    End If
        end if
       
    end function
	'FIN INTERNO-8479 MS 11-11-2016 
//-->
</script>

<body>
    <form name="frmGiroDeposito" method="POST">
    <center>
        <table border="0" cellpadding="1">
            <tr height="10">
				<!--INTERNO-8479 MS 11-11-2016 -->
                <td colspan="3" style="display: <%=sDisplay%>">
                    <% if sPagador = "SW" then %>
                        Banco Pagador<br />
                    <% else %>
                        Banco<br />
                    <% end if %>
                    <span style="display: <%=sDisplay%>">
                        <select name="cbxbanco" style="width: 280px; color: Black; font-weight: bold">
                            <%	If sPagador = "IK" Then
						        CargarBancoIK  sBanco
					        End If %>
                            <% If sPagador = "BC" Then
						        CargarBancoBCP sBanco
					        End If %>
                            <% If sPagador = "FP" Then
						        CargarBancoFPI sBanco
					        End If %>
                            <% If sPagador = "SW" Then
                                'miki SMC-9 MM 2015-11-30
						        CargarBancoSW sPagador, sFormaPago, sPais
					        End If %>
                        </select>
                    </span>
                </td>
				<!-- FIN INTERNO-8479 MS 11-11-2016 -->
            </tr>
            <!--miki SMC-9 MM 2015-11-30-->
			<!-- INTERNO-8479 MS 11-11-2016 -->
            <% If (sPagador <>"SW") or (sPagador = "SW" and sFormaPago = "1") Then
                
                If sPagador = "SW" and (sPais = "HT" or sPais = "DO") Then%>
                  <tr height="10" style="display:block;">
                  <% else %>
                  <tr height="10"  style="display:none;">
                 <%  End if %>
                    <td colspan="3">
                            Agencia Pagadora<br />
                        <span>
                            <select name="cbxAgenciaBanco" style="width: 380px; color: Black; font-weight: bold">
						    <% 
                                CargarAgenciaBancoSW sPagador, sFormaPago, sPais
                            %>
                            </select>
                        </span>
                    </td>
                  </tr>
             <!-- INTERNO-8479 MS 11-11-2016 -->
            <tr>
            <% Else%>
            <tr style="display: none;">
            <% End If%>

                <td style="display: <%=sDisplay%>">
                    Tipo Cuenta<br />
                    <select name="cbxTipoCta" style="width: 280px; color: Black; font-weight: bold">
                        <% If sPagador = "IK" Then
						CargaTipoCuentaInterbank sCuenta
					End If %>
                        <% If sPagador = "BC" Then
						CargaTipoCuentaBCP sCuenta
					End If %>
                        <% If sPagador = "FP" Then
						CargaTipoCuentaFPI sCuenta
					End If %>
                        <% If sPagador = "SW" Then
						CargaTipoCuentaSW
					End If %>
                    </select>
                </td>
                <td colspan="2" style="display: <%=sDisplay%>">
                    Numero Cuenta<br />
                    <input style="height: 22px; text-align: right; width: 150px; color: Black; font-weight: bold"
                        name="txtNumeroCta" size="15" onkeypress="IngresarTexto(1)" maxlength="<%=nMaxleng%>" />
                </td>
            </tr>

             <!-- INTERNO-8479 MS 11-11-2016 -->
			 <% If sPagador = "SW" and sFormaPago = "2" Then %>
                <tr>
                    <td align="left">   Teléfono<br />
                        <input style="height: 22px; text-align: right; width: 280px; color: Black; font-weight: bold"
                            name="txtTelefonoB" size="15" onkeypress="IngresarTexto(1)" maxlength="10" value="<%=sTelefonoB %>"/> 
                        
                    </td>
                 </tr>
                <tr>
                    <td align="left">   Dirección<br />
                        <input style="height: 22px; text-align: right; width: 280px; color: Black; font-weight: bold"
                            name="txtDireccionB" size="30" onkeypress="IngresarTexto(3)" maxlength="50" value="<%=sDireccionB %>" />
                        <br /><b>Modificar o confirmar dirección</b>
                    </td>
                </tr>
                 
             <% End if %>
			 <!-- INTERNO-8479 MS 11-11-2016 -->
            <!-- JFMG 25-05-2012 -->
            <% If sPagador <> "IK" and sPagador <>"FP" and sPagador <>"SW" Then%>
            <!-- FIN JFMG 25-05-2012 -->
            <tr>
                <!-- JFMG 25-05-2012 -->
            <% Else%>
            <tr style="display: none;">
            <% End If%>
                <!-- FIN JFMG 25-05-2012 -->
                <td style="display: <%=sDisplay%>">
                    Moneda Depósito<br />
                    <select name="cbxMonedaDeposito" style="width: 230px; color: Black; font-weight: bold">
                        <% If sPagador <> "" Then
			                CargarMonedaDeposito sMonedaDeposito
			            End If %>
                    </select>
                </td>
            </tr>
        </table>
        <table>
            <tr>
                <td align="right">
                    <img border="0" id="imgAceptar" onclick src="../images/BotonAceptar.jpg" width="70"
                        height="20">
                </td>
            </tr>
            <tr>
                <td>
                    <input type="hidden" name="txtvalorCuenta" value="" />
                </td>
            </tr>
        </table>
        <% If sPagador = "FP" or sPagador = "SW" Then%>
        <table>
            <tr>
                <td align="center">
                    <div style="width: 80%">
                        <textarea id="txtMensaje" style="width: 500px; border-style: none; font-weight: bold;
                            background-color: transparent; overflow: hidden; color: #000000;" rows="5" visible="true" disabled></textarea> <!-- INTERNO-8479 MS 11-11-2016 -->
                    </div>
                </td>
            </tr>
        </table>
        <% Else%>
        <% End If%>

    </center>
    </form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</html>
<%
	Response.Expires =0
%>