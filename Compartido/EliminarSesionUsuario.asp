<%@ LANGUAGE = VBSCRIPT %>
<!--# INCLUDE VIRTUAL="/Compartido/Errores.asp" -->
<%

	' elimina la sesion del usuario en la bd
	dim sSQL, rs
	
	on error resume next
	
	sSQL = " exec auditoria.eliminarconexion " & cint(Session("CodigoConexionUsuario"))
	set rs = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
	
	'response.write Session("afxCnxAFEXpress") & " - " & ssql
	'response.end

	if err.number <> 0 then
		set rs = nothing
		mostrarerrorms "Eliminar registro. Comuniquece con Informática. "

	end if
	
	set rs = nothing

	Private Function EjecutarSQLCliente(ByVal Conexion, ByVal SQL)
		Dim rsESQL
		Dim Cnn
		Const adUseClient = 2
		Const adOpenStatic = 3
		Const adLockBatchOptimistic = 4
		Dim sError
		
		On Error Resume Next
	

		Set EjecutarSQLCliente = Nothing		
   
		Set Cnn = server.CreateObject("ADODB.Connection")
		Cnn.CommandTimeout = 600
		Cnn.Open Conexion
	   
		If Err.number <> 0 Then
			Cnn.Close
			Set Cnn = Nothing

			mostrarerrorms sql & "//" & conexion

			MostrarErrorMS "Ejecutar SQL 1"
		End If

		Set rsESQL = server.CreateObject("ADODB.Recordset")
		rsESQL.CursorLocation = 3
		rsESQL.Open SQL, Cnn, 3, 4
	   
		If Err.number <> 0 Then		
			Set rsESQL = Nothing			
			MostrarErrorMS "Ejecutar SQL 1"
			sError = err.description			
			err.clear			
			err.Raise 50000, "EjecutarSQLCliente", sError		
		End If		

		if sError = empty then		
			'If rsESQL Is Nothing Then Exit Function
			Set rsESQL.ActiveConnection = Nothing
			'MostrarErrorMS "Despues"
		end if	
		
		Set EjecutarSQLCliente = rsESQL
		'rsESQL.Close
		Set rsESQL = Nothing
		
		Cnn.Close
		Set Cnn = Nothing
	End Function	

%>
<html>
<body>
	<script language="vbscript">
	<!--
		window.dialogleft = 0
		window.dialogtop = 0
		window.dialogheight = 0
		window.dialogwidth = 0

		sub window_onload()
			window.close
		end sub
	-->
	</script>
</body>
</html>