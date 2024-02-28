<%
   ' Cierra la conexion si se encuentra abierta 
    if (conexion) then  
	    oConn.close
    Set oConn = nothing		
	end if
%>