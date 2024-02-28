<%@ LANGUAGE = VBScript %>

<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<%	
	
Dim CodigoClienteXc  
Dim CodigoClienteXp  
Dim sCodigo, sNombre
Dim nTipoCliente
Dim nClienteAgencia
Dim sPaisPas

nTipoCliente = 0
nClienteAgencia = 0

sRut = Request("RutCautela")
sNserie = Request("NserieCautela")
sPas = Request("PasCautela")
sNombre = Request("NomCautela")
sPaterno = Request("PatCautela")
sMaterno = Request("MatCautela")
sDescripcion = Request("DescCautela")
sPaisPas = Request("PaisPas")

sHistoria = "Agrega Cliente CAUTELA: " & sDescripcion

nTipoCliente = Request("TC")

AgregarClienteCautela

BuscaCLienteCautela 
AgregaHistoria
    
Response.Redirect "http:DetalleClienteCautela.asp?cc=" & sCodigo

Private Sub BuscaCLienteCautela

    On Error Resume Next
    
    sSQL = "EXECUTE [ObtenerCLienteCautela] "

    If Not IsNull(sRut) And sRut <> "" And sRut <> "null" then
        sSQL = sSQL & "@Rut = '" & valorrut(sRut) & "'"
    End If
    
    If Not IsNull(sPas) And sPas <> ""  And sPas <> "null" then
        sSQL = sSQL & "@Pasaporte = '" & sPas & "'"
    End If
    
    Dim rs		
	Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	If rs.Fields.Count > 0 Then
	    sCodigo = rs("codigo")
    Else
        sCodigo = 0
    End If
	
	If Err.Number <> 0 Then 
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br> BuscaCLienteCautela " & Err.Description & sSQL
		Exit Sub	
    End If

End Sub

Private Sub AgregaHistoria
	AgregarHistoria sCodigo, sHistoria, 1, 0
End Sub

Private Sub AgregarClienteCautela
    On Error Resume Next
	If Err.number <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en HágaseCliente 0&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
	End If
	sCodigo = 0
	sCodigo = AgregarCautela(Session("afxCnxCorporativa"), "WB", sRut, sNserie, sPas, nTipoCliente, sPaterno, sMaterno, sNombre)

End Sub

Function AgregarCautela(ByVal Conexion, ByVal Sucursal, ByVal Rut, ByVal NumSerie, ByVal Pasaporte, ByVal Tipo, ByVal ApellidoPaterno,  ByVal ApellidoMaterno, ByVal Nombres)
	Dim sSQL
	Dim BD
	Dim rsCodigoCliente
	Dim Mensaje
	Dim NombreCompleto 
      
	AgregarCautela = 0
   
	' controla los errores
	On Error Resume Next
	If Rut <> Empty Then
		Rut = ValorRut(Rut)
	End If
	
	CodigoClienteXc = Empty
	CodigoClienteXp = Empty
	
	NombreCompleto = Trim(Nombres) & " " & Trim(ApellidoPaterno) & " " & Trim(ApellidoMaterno)
	
    Dim  ipUsuario, nombrePC
	ipusuario = request.servervariables("REMOTE_ADDR")
    nombrePC = request.servervariables("REMOTE_HOST")

    
    
    If Pasaporte = "NULL" Or IsNull(Pasaporte) then
       sSQL = " exec  InsertarClienteCorporativoCautela " & EvaluarStr(Rut) & ", " & EvaluarStr(NumSerie) & ", " & EvaluarStr(NombreCompleto) & ", " & EvaluarStr(Nombres) & ", " & EvaluarStr(ApellidoPaterno) & ", " & EvaluarStr(ApellidoMaterno) & ", " & Pasaporte  & ", " & EvaluarStr(sPaisPas) & "," & Session("CodigoCliente") & ", 1"
    Else
	    sSQL = " exec InsertarClienteCorporativoCautela " & EvaluarStr(Rut) & ", " & EvaluarStr(NumSerie) & ", " & EvaluarStr(NombreCompleto) & ", " & EvaluarStr(Nombres) & ", " & EvaluarStr(ApellidoPaterno) & ", " & EvaluarStr(ApellidoMaterno) & ", " & EvaluarStr(Pasaporte)  & ", " & EvaluarStr(sPaisPas) & "," & Session("CodigoCliente") & ", 1"
    End If

	Dim rs1		
	Set rs1 = EjecutarSQLCliente(Conexion, sSQL)

    If Err.Number <> 0 Then 
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Function	
    End If

    Set rs = Nothing
End Function

Function EvaluarStr(ByVal Valor)
	Dim Devuelve
	
	If Valor="" Then 
		EvaluarStr = "Null"	
	Else
		EvaluarStr = "'" & Valor & "'"
	End If

End Function

	

%>
