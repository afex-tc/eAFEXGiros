<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%

	' verifica si el cliente ya fue consultado
	Dim rs 
	Dim sSQL
	Dim sMensaje 
	Dim bClienteConsultado
				
	bClienteConsultado = False
				
	On Error Resume Next
		
	sSQL = " exec Mensajeria.VerificarClienteConsultado " & EvaluarStr(Session("ATCAFEXpress")) 	
	set rs = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
	if err.number <> 0 then
		sMensaje = "Error al verificar Mensajeria. " & err.Description 
	elseif not rs.eof then
		bClienteConsultado = rs("ClienteConsultadoMensajeria")
	end if
	set rs = nothing
	
%>

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<body LANGUAGE="VBScript">
<script LANGUAGE="VBScript">
<!--

	If <%=bClienteConsultado%> Then
		window.dialogWidth = 0
		window.dialogHeight = 0
		window.dialogLeft = 0
		window.dialogTop = 0
	Else
		window.dialogWidth = 25
		window.dialogHeight = 15
		window.dialogLeft = 300
		window.dialogTop = 200
	End If
		
	Sub Window_onLoad()
		If <%=bClienteConsultado%> Then			
			window.returnvalue = ""
			window.close					
		End If
		frmMensajeria.txtPresionoAceptar.value = ""
		
	End Sub
	
	Sub window_onUnload()		
		If frmMensajeria.txtPresionoAceptar.value <> "1" and Not <%=bClienteConsultado%> Then
			frmMensajeria.txtPresionoAceptar.value = "2"
			imgAceptar_onclick
		End If
	End Sub
	
	Sub imgAceptar_onclick()
		Dim bDeseaMensajeria
		bDeseaMensajeria = True 
		
		if frmMensajeria.chkCelular.checked then
			if frmMensajeria.txtCelular.value = "" then
				IF frmMensajeria.txtPresionoAceptar.value = "2" Then
					botonPresionado = msgbox("Se indicará al sistema que el Cliente no desea recibir Mensajería.",,"AFEX")
					bDeseaMensajeria = False
				Else 
					msgbox "Debe ingresar el Celular.",,"AFEX"
					exit sub
				End If
			end if
		end if
		
		if frmMensajeria.chkEmail.checked then
			if frmMensajeria.txtEmail.value = "" then
				IF frmMensajeria.txtPresionoAceptar.value = "2" Then
					botonPresionado = msgbox("Se indicará al sistema que el Cliente no desea recibir Mensajería.",,"AFEX")
					bDeseaMensajeria = False
				Else 
					msgbox "Debe ingresar el Email.",,"AFEX"
					exit sub
				End If
			end if
		end if
		
		if not frmMensajeria.chkCelular.checked and not frmMensajeria.chkEmail.checked then
			dim botonPresionado
			IF frmMensajeria.txtPresionoAceptar.value = "2" Then
				botonPresionado = msgbox("Se indicará al sistema que el Cliente no desea recibir Mensajería.",,"AFEX")
				bDeseaMensajeria = False
			Else 
				botonPresionado = msgbox("Se indicará al sistema que el Cliente no desea recibir Mensajería.",4,"AFEX")
				if botonpresionado <> 6 then exit sub
				bDeseaMensajeria = False
			End If			
		end if
				
		frmMensajeria.txtPresionoAceptar.value = "1"
				
		Dim sString 
		
		If bDeseaMensajeria Then
			sString = "1"			
		Else
			sString = "0"
		End if
		sString = sString & ", 1" ' pregunta
		If frmMensajeria.txtCelular.value <> "" Then
			sString = sString & _
					", 1" & _
					", " & frmMensajeria.txtCelular.value
		Else 
			sString = sString & _
					", 0" & _
					", "
		End If
		If frmMensajeria.txtEmail.value <> "" Then
			sString = sString & _
					", 1" & _
					", " & frmMensajeria.txtEmail.value
		Else
			sString = sString & _
					", 0" & _
					", "
		End If
		
		window.returnvalue = sString
		window.close
		
		'frmMensajeria.txtCelular.disabled = False
		'frmMensajeria.txtEmail.disabled = false 
		'frmMensajeria.action = "SolicitarMensajeriaCliente.asp?Accion=1&DeseaMensajeria=" & bDeseaMensajeria
		'frmMensajeria.submit()
		'frmMensajeria.action = ""
	End Sub
		
	sub ValidarTecla(Tipo)
		dim iTecla
		
		iTecla = window.event.keyCode 
		
		If Tipo = 1 then					
			if iTecla < 48 or iTecla > 57 then		
				window.event.keyCode = 0		
			end if
		else
			if iTecla < 48 or iTecla > 57 then
				if (iTecla < 65 or iTecla > 90) and _
				   (iTecla <> 64 and iTecla <> 46 and iTecla <> 95 and iTecla <> 95) then
					if iTecla < 97 or iTecla > 122 then
						window.event.keyCode = 0
					end if
				end if
			end if
		End if 
	end sub
	
	Sub chkCelular_onClick()
		If frmMensajeria.chkCelular.checked Then
			frmMensajeria.txtCelular.disabled = false 
			frmMensajeria.txtCelular.focus()
		else
			frmMensajeria.txtCelular.value = ""
			frmMensajeria.txtCelular.disabled = true
		End if
		
	End Sub
	
	Sub chkEmail_onClick()
		If frmMensajeria.chkEmail.checked Then
			frmMensajeria.txtEmail.disabled = false 
			frmMensajeria.txtEmail.focus()
		else
			frmMensajeria.txtEmail.value = ""
			frmMensajeria.txtEmail.disabled = true
		End if
		
	End Sub
	
	Sub txtCelular_onBlur()
		If frmMensajeria.txtCelular.value = "" Then Exit Sub 
	
		If Len(frmMensajeria.txtCelular.value) <> 8 Then
			msgbox "El Celular es incorrecto, favor verificar la Cantidad de dígitos.",,"AFEX"
			frmMensajeria.txtCelular.value = ""
			frmMensajeria.txtCelular.focus() 
		End if
	End Sub
-->
</script>

<script language="javascript">

	function valEmail(valor){
		re=/^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/ // /^[_a-z0-9-]+(.[_a-z0-9-]+)*@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,3})$/
		var valorEmail = document.getElementById(valor).value
		if(!re.exec(valorEmail))    {
			alert('Email incorrecto');
			document.getElementById(valor).value = '';
		    return false;
		}
		else{
			return true;
		}
	}

</script>

<form name="frmMensajeria" method="post" action="">
<center>
<table class="Borde" width="400px" cellSpacing="1" cellPadding="0" border="0">  
	<tr>
		<td colspan="2" class="titulo">Mensajeria</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>	
	<tr>
		<td colspan="2"><b>Sr(a). Cajero(a) Favor realizar la siguiente pregunta al Cliente:</b></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>	
	<tr>
		<td colspan="2" align="center">
			¿Le gustaria recibir una confirmación del pago del giro <br />
			en su celular, correo electrónico o ambos?
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="100px">
			<input type="checkbox" name="chkCelular">Celular</input>
		</td>
		<td>
			<input type="text" onkeypress="ValidarTecla(1)" maxlength="8" disabled name="txtCelular" size="10">&nbsp;
			<font size="1px" color="Gray">sin "0" a la izquierda</font>
		</td>
	</tr>
	<tr>
		<td >
			<input type="checkbox" name="chkEmail">Email</input>
		</td>
		<td>
			<input type="text" onblur="valEmail('txtEmail')" onkeypress="ValidarTecla(2)" maxlength="50" disabled name="txtEmail" size="30" >
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<IMG id=imgAceptar style="CURSOR: hand" onclick ="" height=20 src="../images/BotonAceptar.jpg" width="70" border="0" >
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>  
</table>
	<input type="hidden" name="txtPresionoAceptar" value="" />
</center>
</form>

</body>
</html>
