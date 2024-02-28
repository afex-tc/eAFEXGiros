<%@ Language=VBScript%>
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	dim sSQL, rs
	dim sEstadoGiro, swEditado
	
	on error resume next
	
	sSQl = "exec mostrardatosgiro " & request("CodigoGiro")
	set rs = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
	if err.number <> 0 then
		MostrarErrorMS "Error al consultar el Giro."
	end if
	if not rs.eof then	
		sEstadoGiro = cint("0" & rs("estado_giro"))
		swEditado = cint("0" & rs("sw_editado"))	
		
		' verifica el estado del giro
		select case sEstadoGiro
			case 1	' Captacion
				if trim(rs("pais_beneficiario")) = "CL" then '	giro a pagar en chile
					select case swEditado
						case 0
							sMensaje = "Su Giro está en estado Recibido. "
						case 1
							sMensaje = "Su Giro está en estado Disponible. "							
					end select
					
				else	' giro a pagar en el extranjero
					sMensaje = "Su Giro está Pendiente de Envío. "
				end if
				
			case 2	' Envio
				sMensaje = "Su Giro está Disponible de Pago. "
			case 3	' Confirmacion de pago desde el pagador
				sMensaje = "Su Giro ha sido pagado al beneficiario. "
			case 4  ' Aviso
				sMensaje = "Su Giro está en estado de Aviso. " & trim(rs("HistoriaActual"))
			case 5	' Pago
				sMensaje = "Su Giro ha sido pagado el día " & trim(rs("fecha_pago")) & " " & trim(rs("hora_pago")) & ". "
			case 6	' Confirmacion de pago al agente captador
				sMensaje = "Su Giro ha sido pagado al beneficiario."
			case 7	' Reclamo
				sMensaje = "Su Giro se encuentra RECLAMADO temporalmente."
			case 8	' Solucionado
				sMensaje = "Su Giro procesado nuevamente ya que se encontraba en proceso de RECLAMO."
			case 9	' Anulado
				sMensaje = "Su Giro ha sido ANULADO."
		end select
		
		rs.close
	else
		sMensaje = "No existe el Giro en nuestros registros. Verifique si el código ingresado está correcto."
	end if

	set rs = nothing
%>
<HTML>
	<BODY>		
		<script language="VBScript">
		<!--
			window.dialogleft = 0
			window.dialogtop = 0
			window.dialogheight = 0
			window.dialogwidth = 0
			
			sub window_onload()
				window.returnvalue = "<%=sMensaje%>"
				window.close 
			end sub			
			
		-->
		</script>
		
	</BODY>
</HTML>