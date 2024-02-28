<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
tnumero = request.QueryString("tnumero") 
nmn= request.QueryString ("nmn")
mn = request.QueryString ("mn")
mto= request.QueryString ("mto")
bs = request.QueryString ("bs")

Dim sSQL3,sSQL4,rs4,sw,cnn,rs3

		'On Error Resume Next		
			Set Cnn = server.CreateObject("ADODB.Connection")		
			Cnn.CommandTimeout = 60
			Cnn.Open Session("afxCnxAFEXpress")	   
			Set rs3 = server.CreateObject("ADODB.Recordset") 
			Set rs4 = server.CreateObject("ADODB.Recordset") 
			'response.Write bs
			sSQL3 = "SELECT *  FROM  tarjeta where numero_boleta= "& bs &"  order by numero_boleta "		
			sSQL4 = "select * from tarjeta where  monto= "& mto &" and codigo_moneda= '"& mn &"' and numero_boleta is null and codigo_giro is null and tipo_pin=2 and codigo_usuario is null"
			'response.Write ssql3
	    	rs3.open sSQL3,cnn, 3, 1
	    	rs4.Open sSQL4,cnn, 3, 1 

		'If Err.number <> 0 Then
			'Set rs3 = Nothing
			'MostrarErrorMS "BuscarBoleta", Err.Description			
		'End If
'	response.Write sw
		if rs3.EOF then	  
			if rs4.EOF then
				sw=11
			else
				sw=0
			end if	
		elseif clng(rs3.Fields("numero_boleta").value) = clng(bs)then
			sw=1			
		end if	
		
		'response.Write sw
		
		Set rs3 = Nothing	
		Set rs4 = Nothing
		Cnn.Close
		Set Cnn = Nothing
		'response.Write sw
		response.redirect "VenderTarjeta.asp?sw="& sw & "&nmn=" & nmn & "&mn=" & mn & "&mto=" & mto & "&bs= " & bs 		
	

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

 


</BODY>
</HTML>
