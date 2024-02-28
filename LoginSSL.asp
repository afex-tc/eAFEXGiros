
<%
'IMPORTANTE ****************************************************************
'Para incluir esta página en otra se deben agregar las siguientes
'lineas al inicion en la página fuente
'***************************************************************************
'
'<!--INCLUDE virtual="/compartido/Constantes.asp" -->
'<!--INCLUDE virtual="/compartido/Rutinas.asp" -->
'<!--INCLUDE virtual="/compartido/Errores.asp" -->
'<!--INCLUDE virtual="/compartido/LoginSSL.asp" -->
'
'***************************************************************************
Sub ValidarCliente()
	ValidarClienteSSL sURL, sErrorUsuario, nTipo, "", ""		
End Sub

Function ValidarClienteSSL(ByRef sURL, ByRef sErrorUsuario, ByRef nTipo, ByVal Usuario, ByVal Password)
	Dim afxCliente
	Dim rs, sString
	Dim sUsuario, sPassword
	Dim sCodigo, sCodigoSistemas, nCampo
		
	On Error Resume Next
	ValidarClienteSSL = False
	sUsuario = Usuario
	sPassWord = Password
	If sUsuario = "" Then 
		sUsuario = Request.Form("txtUsuario")
		sPassword = Request.Form("txtPassword")	
	End If
	If sUsuario = "" Then 
		sUsuario = Request("sslUsuario")
		sPassword = Request("sslPassword")
	End If
	If sUsuario = "" Then  Exit Function
		
	If sUsuario = "moneda" And sPassWord = "1234" Then
		Session("ModoPrueba") = True
		sUsuario = "afexmoneda"
		sPassword = "1235"
		CambiarODBC
	End If
	If sUsuario = "ibero" And sPassWord = "4567" Then
		Session("ModoPrueba") = True
		sUsuario = "ibero"
		sPassword = "123456789"
		CambiarODBC
	End If
	If sUsuario = "mx" And sPassWord = "1234" Then
		Session("ModoPrueba") = True
		sUsuario = "multiexpress"
		sPassword = "123456789"
		CambiarODBC
	End If
		
	Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
	Set rs = Server.CreateObject("ADODB.recordset")
	Set rs = afxCliente.ObtenerCliente(Session("afxCnxCorporativa"), sUsuario, sPassword)
	If Err.Number <> 0 Then			
		sErrorUsuario = Err.description 
		Set rs = Nothing
		Set afxCliente = Nothing
		Exit Function
	End If
		
	If afxCliente.ErrNumber <> 0 Then			
		sErrorUsuario = replace(afxCliente.ErrDescription, vbCrLf , "")
		Set rs = Nothing
		Exit Function
	End If
		
	If rs("estado") = 0 Then			
		sErrorUsuario = "Usted aún no está habilitado para operar con nuestros servicios AFEXweb. Comuníquese con nuestro departamento de Atención al Cliente al teléfono (562) 6369052 o en nuestra zona de «Contáctenos» de esta página o al correo «atencion.clientes@afex.cl»"
		Set rs = Nothing
		Exit Function
	End If
		
	sCodigo = rs("codigo")
	nTipo = rs("tipo")		
		
	Session("CodigoCliente") = rs("codigo")
	Session("NombreCliente") = rs("nombre")
	If Not IsNull(rs("codigo_agente")) Then
		Session("CodigoAgente") = uCase(Trim(rs("codigo_agente")))
		Session("Categoria") = rs("categoria_agente")
		Session("CiudadCliente") = Trim(UCase(rs("ciudad_agente")))
		Session("PaisCliente") = Trim(UCase(rs("pais_agente")))
	Else
		Session("Categoria") = 0
		Session("CiudadCliente") = Trim(UCase(rs("codigo_ciudad")))
		Session("PaisCliente") = Trim(UCase(rs("codigo_pais")))
	End If		
	Session("AFEXpress") = EvaluarVar(rs("codigo_afexpress"), "")
	Session("AFEXchange") = EvaluarVar(rs("codigo_afexchange"), "")
	Session("CodigoCaja") = EvaluarVar(rs("codigo_caja"), "")
	Session("FechaApertura") = "01-01-2003"
	Session("NombreUsuario") = EvaluarVar(rs("nombre_usuario"), "")

	Set afxCliente = Nothing
	Set rs = Nothing
		
	'response.Redirect "compartido/error.asp?description=" & Session("CodigoCliente") & ", " & Session("NombreCiente")		
	Select Case nTipo
	Case 4 
		sURL = "Agente/Default.asp?Codigo=" & sCodigo
			
	Case 5
		sURL = "Agente/Default.asp?Codigo=" & sCodigo
			
	Case Else			
		Session("CodigoAgente") = "WB"
		sURL = "Cliente/Default.asp?Codigo=" & sCodigo

	End Select
	ValidarClienteSSL = True
End Function

Sub CambiarODBC()
	Session("afxCnxCorporativa") = "DSN=wAfexCorporativa;UID=corporativa;PWD=afxsqlcor;"
	Session("afxCnxAFEXchange") = "DSN=wAFEXchange;UID=cambios;PWD=cambios;"
	Session("afxCnxAFEXpress") = "DSN=wAFEX_giros;UID=giros;PWD=giros;"
	Session("afxCnxAFEXweb") = "DSN=wAFEXweb;UID=cambios;PWD=cambios;"
End Sub	

Function EvaluarVar(ByVal Valor, _
						  ByVal Devuelve)
	If Devuelve = Empty Then Devuelve = ""		
   If IsNull(Valor) Then
      EvaluarVar = Devuelve
   Else
      EvaluarVar = Valor
   End If

End Function

%>
