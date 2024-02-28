<%@  language="VBScript" %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%
dim Cod_Giro,Monto,Detalle
Dim sSQL4,rs4
dim invoice,codi,titulo

Cod_Giro=request.QueryString("Giro")
Monto=request.QueryString("Monto")
Detalle=request.QueryString ("Detalle")


	    On Error Resume Next
	
	    sSQL4="SELECT * FROM GIRO with(nolock) WHERE codigo_giro= '" & Cod_giro & "' "
		SET rs4 = ejecutarsqlcliente(session("afxcnxAfexpress"),sSQL4)
			
		invoice=(rs4.Fields("invoice").Value )
			
		if (trim(invoice)<>empty) then
		    codi=invoice
			titulo="Invoice"
		else
		    codi=cod_giro
			titulo="Codigo Giro"
		end if	
	
	    
        ' Tecnova rperez 30-08-2016 - Sprint 2: Se comenta bloque de código por requerimiento INTERNO 5810
	    '' JFMG 13-06-2012 temporalmente lo deja en reclamo para que no sea enviado
        'Dim sSQL, rs
        'If rs4("estado_giro") = 1 Then
        '    sSQL = "execute ReclamarGiroAnular '" & Cod_Giro & "', '" & Session("CodigoAgente") & "', '" & Session("NombreUsuarioOperador") & "'"		
        '    SET rs = ejecutarsqlcliente(session("afxcnxAfexpress"),sSQL)
        '    If err.number <> 0 then
        '        MostrarErrorMS err.Description                 
        '    end if
        '    'response.Write ssql
        '    'response.end
        'End If
        '' FIN JFMG 13-06-2012		

        ' Tecnova rperez 30-08-2016 - Sprint 2
        Dim sSQL, rs
        

        If rs4("estado_giro") = 1 Then
            sSQL = "EXEC AnularGiro " & EvaluarStr(cod_giro) & ", " & EvaluarStr(Session("NombreUsuarioOperador")) & ", 1" ' 1=Con Comisión

            SET rs = EjecutarSqlCliente(Session("afxcnxAfexpress"),sSQL)
            If Err.number <> 0 Then
                MostrarErrorMS "Enviar Anulación "
            End If
        End If

        ' Verifica si el giro se actualizó
		sSQL = "SELECT estado_giro FROM Giro WITH(NOLOCK) WHERE codigo_giro=" & EvaluarStr(cod_giro)	
    		
		SET rs = EjecutarSqlCliente(session("afxCnxAFEXpress"), sSQL)
		If Err.number <> 0 Then
			Set rs = Nothing
			MostrarErrorMS "Enviar Anulación - Consulta "
		End If
		If rs.EOF Then
			Set rs = Nothing
			MostrarErrorMS "Enviar Anulación - Validación "
		End If
		If Not rs.EOF Then
			If rs("estado_giro") <> 7 and rs("estado_giro") <> 9 Then
				MostrarErrorMS "El giro no se actualizó. Comuníquese con Informática para informarles del problema."
			End If
		End if

        ' Agrega la historia
        Dim d, fecha, hora, motivo

        Function Lpad (sValue, sPadchar, iLength)
          Lpad = string(iLength - Len(sValue),sPadchar) & sValue
        End Function
        
	    d = Now
	    fecha = Day(d) & "/" & Month(d) & "/" & Year(d) & " " & Hour(d) & ":" & Minute(d) & ":" & Second(d)
        hora = Lpad(Hour(d), "0", 2) & Lpad(Minute(d), "0", 2) & Lpad(Second(d), "0", 2)
        motivo = "Giro Anulado por Sucursal"

        sSQL = "INSERT INTO Historia (codigo_giro, fecha, hora, descripcion, codigo_usuario) " & _
                "VALUES (" & EvaluarStr(cod_giro) & ", " & EvaluarStr(fecha) & ", " & hora & _
                         ", " & EvaluarStr(motivo) & ", " & EvaluarStr(Session("NombreUsuarioOperador")) & ")"

        SET rs = EjecutarSqlCliente(session("afxCnxAFEXpress"), sSQL)
        If Err.number <> 0 Then
            MostrarErrorMS "Agregar Historia "
        End If
		
	    SET rs = nothing
		Set rs4 = Nothing
%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <title>Anular Giro</title>
    <link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
</head>
<body>
    <script language="VBScript">
<!--
	Const sEncabezadoFondo = "Anular Giro"
	Const sEncabezadoTitulo = "Anular Giro"
	Const sClass = "TituloPrincipal"
	
-->
    </script>
    <!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
    <!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->


    <br>
    <form name="Anula" method="post">
        <input type="hidden" name="Cod_Giro" value="<%=Cod_Giro%>">
        <input type="hidden" name="Monto" value="<%=Monto%>">
        <input type="hidden" name="Detalle" value="<%=Detalle%>">
        <table border="0">
            <tr>
                <td width="10%"></td>
                <td width="90%">
                    <table border="0">
                        <tr>
                            <td>

                                <%=titulo%> :<b> <%=codi%></b>
                            </td>
                        </tr>
                        <tr>
                            <td>Monto		: <%=Monto%> Dolares
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%=Detalle%>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                        </tr>
                        <tr>
                            <td></td>
                        </tr>
                        <tr>
                            <td>
                                <font face="verdana" size="2">Nombre Sucursal o Agente</font>
                                <br>
                                <input type="text" name="txtNombre" maxlength="50" id="txtNombre" onblur="Anula.txtNombre.value=MayusculaMinuscula(Anula.txtNombre.value)"></input><br>
                                <font face="verdana" size="2">Email de Contacto</font>
                                <br>
                                <input type="text" name="txtEmail" maxlength="50" id="txtEmail"></input>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                        </tr>
                        <tr>
                            <td>
                                <font face="verdana" size="2">Motivo de Anulación<br>
					<textarea name="txtAnular" rows="10" cols="50" ></textarea>				
					</font>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table align="center" border="0">
                                    <tr>
                                        <td>
                                            <input type="button" id="cmdEnviar" value=" Enviar "></input></td>
                                        <td>
                                            <input type="button" id="cmdLimpiar" value=" Limpiar "></input></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
<script language="vbScript">
	
		Sub cmdEnviar_onClick()
            
            ' Tecnova rperez 29-08-2016 - Sprint 2 Validación de campos vacíos
            Dim nombre, email, motivo

            nombre = Trim(Anula.txtNombre.value)
            email = Trim(Anula.txtEmail.value)
            motivo = Trim(Anula.txtAnular.value)

            If nombre = "" Then
                MsgBox ("Ingrese el Nombre de Sucursal o Agente")
                Anula.txtNombre.focus()
			ElseIf email = ""  Then
				MsgBox ("Ingrese el Email de Contacto")
                Anula.txtEmail.focus()					
            ElseIf motivo = "" Then
                MsgBox ("Ingrese el Motivo de Anulación")
                Anula.txtAnular.focus()				
			Else	
			 	Anula.action = "EnviarAnulacion.asp"
				Anula.submit
				Anula.action = ""
			End If	
		
		End Sub
		
		Sub LimpiarControles
			
			Anula.txtNombre.value = ""
			Anula.txtEmail.value = ""
			Anula.txtAnular.value = ""
		
		End Sub
		
		Sub cmdLimpiar_onclick()
			LimpiarControles
		End Sub
</script>
</html>
