<%
 ConnString = "DSN=AFEXcorporativa;UID=corporativa;PWD=afxsqlcor;"

 ' Condición para abrir una conexion nueva
 if (conexion) then  
    set oConn = Server.CreateObject("ADODB.Connection")
        oConn.Open ConnString
 end if
%>
<style type="text/css">
select
{
    FONT-SIZE: 8pt;
    COLOR: #344249
}
input
{
    BORDER-RIGHT: #40515b 1px solid;
    BORDER-TOP: #40515b 1px solid;
    FONT-SIZE: 8pt;
    BORDER-LEFT: #40515b 1px solid;
    COLOR: #344249;
    BORDER-BOTTOM: #40515b 1px solid
}
td
{
    FONT-WEIGHT: normal;
    FONT-SIZE: 10px;
    COLOR: #313e53;
    FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif
}
</style>
<!--#include file="includes/formularios.asp"-->