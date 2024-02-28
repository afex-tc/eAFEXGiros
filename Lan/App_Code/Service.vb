Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols

<WebService(Namespace:="http://www.afex.cl")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function GrabarGiro(ByVal Captador, ByVal Pagador, ByVal Monto, ByVal NombresB, ByVal ApellidosB, _
                                ByVal DireccionB, ByVal CiudadB, ByVal PaisB, ByVal FonoB, ByVal NombresR, ByVal ApellidosR, ByVal DireccionR, _
                                ByVal CiudadR, ByVal PaisR, ByVal FonoR, ByVal Usuario) As String
        Dim smensaje As String
        Dim sSQL As String
        Dim Conexion As String
        Dim rs As ADODB.Recordset

        Conexion = "Provider=SQLOLEDB.1;Password=giros;User ID=giros;Initial Catalog=giros;Data Source=cipres;"

        sSQL = " execute enviargiro " & EvaluarSTR(Captador) & ", " & EvaluarSTR(Pagador) & ", " & EvaluarNUM(Monto) & ", " & _
                                    " 2, 0, 1, 0, 'USD', 'USD', null, null, NULL, NULL, NULL, " & _
                                    EvaluarSTR(NombresB) & ", " & EvaluarSTR(ApellidosB) & ", " & _
                                    EvaluarSTR(DireccionB) & ", " & EvaluarSTR(CiudadB) & ", NULL, " & EvaluarSTR(PaisB) & ", " & _
                                    " 1, 512, " & FonoB & ", null, null, null, " & EvaluarSTR(NombresR) & ", " & _
                                    EvaluarSTR(ApellidosR) & ", " & EvaluarSTR(DireccionR) & ", " & EvaluarSTR(CiudadR) & ", null, " & _
                                    EvaluarSTR(PaisR) & ", 56, 2, " & FonoR & ", " & EvaluarSTR(Usuario) & ", 'AF465148', 'AF465148', " & _
                                    "null, null, 2, 4, 1, 2, 1, 4, null, null, null, null, null"
        rs = EjecutarSQLCliente(Conexion, sSQL)
        If Err.Number <> 0 Then
            Err.Raise(50000, "Error al grabar el giro." & Err.Description)
        Else
            GrabarGiro = rs.Fields(0).Value
            rs.Close()
        End If
        rs = Nothing

        '    sSQL = " execute enviargiro " & _
        '  EvaluarSTR(Session("CodigoAgente")) & ", " & EvaluarSTR(sPagador) & ", " & FormatoNumeroSQL(cCur(CDbl(Request.Form("txtMonto")))) & ", " & _
        ' FormatoNumeroSQL(cCur(CDbl(Request.Form("txtTarifaCobrada")))) & ", " & nPrioridad & ", 1, 0," & EvaluarSTR(Request.Form("cbxMonedaGiro")) & ", " & _
        'EvaluarSTR(Request.Form("cbxMonedaPago")) & ", " & EvaluarSTR(Request.Form("txtMensajeB")) & ", " & EvaluarSTR(Request.Form("txtMsjPagador")) & ", " & _
        '" NULL, NULL, NULL, " & EvaluarSTR(sNombreB) & ", " & EvaluarSTR(sApellidoB) & ", " & EvaluarSTR(sDireccion) & ", " & EvaluarSTR(Request.Form("cbxCiudadB")) & ", " & _
        '" NULL, " & EvaluarSTR(Request.Form("cbxPaisB")) & ", " & CInt(0 & Request.Form("txtPaisFonoB")) & ", " & CInt(0 & Request.Form("txtAreaFonoB")) & ", " & _
        'cCur(0 & Request.Form("txtFonoB")) & ", " & EvaluarSTR(Request.Form("txtRut")) & ", " & EvaluarSTR(Request.Form("txtPasaporte")) & ", " & _
        'EvaluarSTR(Request.Form("cbxPaisPasaporte")) & ", " & EvaluarSTR(Trim(Request.Form("txtNombres")) & Trim(Request.Form("txtRazonSocial"))) & ", " & _
        'EvaluarSTR(Request.Form("txtApellidos")) & ", " & EvaluarSTR(Request.Form("txtDireccion")) & ", " & EvaluarSTR(Request.Form("cbxCiudad")) & ", " & _
        'EvaluarSTR(Request.Form("cbxComuna")) & ", " & EvaluarSTR(Request.Form("cbxPais")) & ", " & _
        'CInt(0 & Request.Form("txtPaisFono")) & ", " & CInt(0 & Request.Form("txtAreaFono")) & ", " & cCur(0 & Request.Form("txtFono")) & ", " & _
        'EvaluarSTR(Session("NombreUsuarioOperador")) & ", " & EvaluarSTR(sCodigoBeneficiario) & ", " & _
        'EvaluarSTR(sAFEXpress) & ", " & EvaluarSTR(Request.Form("txtInvoiceMG")) & ", " & cCur(0 & Request.Form("txtBoleta")) & ", " & FormatoNumeroSQL(cCur(CDbl(Request.Form("txtTarifaSugerida")))) & ", " & _
        'FormatoNumeroSQL(cCur(CDbl(Request.Form("txtGasto")))) & ", " & FormatoNumeroSQL(cCur(CDbl(Request.Form("txtComisionCaptador")))) & ", " & _
        'FormatoNumeroSQL(cCur(CDbl(Request.Form("txtComisionPagador")))) & ", " & FormatoNumeroSQL(cCur(CDbl(Request.Form("txtComisionMatriz")))) & ", " & _
        'FormatoNumeroSQL(cCur(CDbl(Request.Form("txtAfectoIva")))) & ", " & _
        'EvaluarStr(sBancoBR) & ", " & EvaluarStr(sAgenciaBR) & ", " & EvaluarStr(sCtaCteBR) & ", " & EvaluarStr(sCpfBR) & ", " & cRealesBR


    End Function

    Private Function EjecutarSQLCliente(ByVal Conexion, ByVal SQL) As ADODB.Recordset
        Dim rsESQL As ADODB.Recordset
        Dim smensaje As String
        Dim Cnn As ADODB.Connection

        On Error Resume Next

        EjecutarSQLCliente = Nothing

        Cnn = New ADODB.Connection
        Cnn.CommandTimeout = 600
        Cnn.Open(Conexion)
        If Err.Number <> 0 Then
            Err.Raise(50000, "EjecutarSQLCliente", "Error al conectar a la BD. " & Err.Description)
            Cnn.Close()
            Cnn = Nothing
            Exit Function
        End If

        rsESQL = New ADODB.Recordset
        rsESQL.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsESQL.Open(SQL, Cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
        If Err.Number <> 0 Then
            Err.Raise(50000, "EjecutarSQLCliente", "Error al ejecutar la consulta. " & Err.Description)
        End If

        rsESQL.ActiveConnection = Nothing

        EjecutarSQLCliente = rsESQL

        Cnn.Close()
        Cnn = Nothing
    End Function

    <WebMethod()> _
    Public Function EstadoGiro(ByVal Giro) As String
        Dim sSQL As String
        Dim Conexion As String
        Dim rs As ADODB.Recordset

        Conexion = "Provider=SQLOLEDB.1;Password=giros;User ID=giros;Initial Catalog=giros;Data Source=cipres;"

        sSQL = " select e.descripcion as estadogiro " & _
               " from estados e, giro g " & _
               " where g.codigo_giro = '" & giro & "' and " & _
                    " e.campo = 'ESTADO_GIRO' and e.codigo = g.estado_giro "
        rs = EjecutarSQLCliente(Conexion, sSQL)
        If Err.Number <> 0 Then
            Err.Raise(50000, "EstadoGiro", "Error al consultar el estado del giro. " & Err.Description)
        Else
            EstadoGiro = rs.Fields(0).Value
            rs.Close()
        End If

        rs = Nothing
    End Function

    Private Function EvaluarSTR(ByVal STR) As String
        If STR = "" Then
            EvaluarSTR = "null"
        Else
            EvaluarSTR = "'" & STR & "'"
        End If
    End Function

    Private Function EvaluarNUM(ByVal NUM) As String
        If NUM = "" Then
            EvaluarNUM = "null"
        Else
            EvaluarNUM = Replace(CDbl(NUM), ",", ".")
        End If
    End Function
End Class