<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%

    Dim rs
    Dim sSQL
    Dim sMensaje
    
    On Error Resume Next
    sSQL = " exec EliminarActividadEconomicaCliente " & request("Actividad") & ", " & _
                                                        request("cc") & ", " & _
                                                        evaluarstr(session("NombreUsuarioOperador"))
    Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
    If err.number <> 0 Then
        sMensaje = "0;" & err.Description      
    Else
        sMensaje = "1;"
    End If
    Set rs = Nothing
    
    Response.Expires = 0
%>
<script type="text/vbscript" language="vbscript">
<!--
    window.dialogwidth = 0
    window.dialogleft = 0
    window.dialogtop = 0
    window.dialogheight = 0
    
    window.returnvalue = "<%=sMensaje %>"
    window.close()
    
-->
</script>