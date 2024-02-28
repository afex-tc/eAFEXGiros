<!-- Constantes.asp -->
<%
	
	Const afxAccionNada = 0
	Const afxAccionBuscar = 1
	Const afxAccionNuevo = 2
	Const afxAccionActualizar = 3
	Const afxAccionPais = 4
	Const afxAccionCiudad = 5
	Const afxAccionComuna = 6
	Const afxAccionPagador = 7
	Const afxAccionMonedaPago = 8
	Const afxAccionMonto = 9
	Const afxAccionIngresarMG = 10
	Const afxAccionClienteActual = 11
	Const afxAccionPagar = 12
	Const afxAccionPagarTercero = 13
	Const afxAccionTarifa = 14
	Const afxAccionIngresarTX = 20
	
	Const afxMenuNuevo = 1
	Const afxMenuActualizar = 2
	Const afxMenuNormal = 0
	
   Const afxCampoRut = 1
   Const afxCampoPasaporte = 2
   Const afxCampoCodigo = 3
   Const afxCampoNombres = 4
   Const afxCampoCodigoExchange = 5
   Const afxCampoCodigoExpress = 6

   'Const afxPagoEfectivo = 0
   'Const afxPagoDeposito = 1
	
   'Const afxPagoDomicilio = 0
   'Const afxPagoAgencia = 1

   Const afxPrioridadNormal = 0
   Const afxPrioridadUrgente = 1

   Const afxSentidoRecibido = 0
   Const afxSentidoEnviado = 1
		
	Const afxGiros = 1
	Const afxCambios = 2

	Const afxTCCompra = 1
	Const afxTCVenta = 2		
	Const afxTCTransferencia = 3
	Const afxTCParidad = 4
	
	Const afxEfectivoUSD = 1
	Const afxEfectivoCLP = 2
	Const afxDepositoUSD = 3
	Const afxDepositoCLP = 4
	Const afxCustodiaUSD = 5

   'Const afxOperacionPosicion = 0
   'Const afxOperacionCompra = 1
   'Const afxOperacionVenta = 2
   'Const afxOperacionCanje = 3
   'Const afxOperacionServicios = 4
   'Const afxOperacionEntrada = 5
   'Const afxOperacionSalida = 6
   'Const afxOperacionCanjeEntrada = 7
   'Const afxOperacionCanjeSalida = 8
   'Const afxOperacionComision = 9
   'Const afxOperacionTraspaso = 10
   'Const afxOperacionPagoCheque = 20
   'Const afxOperacionPagoDeposito = 21

   'Const afxProductoIndefinido = 0
   'Const afxProductoEfectivo = 1
   'Const afxProductoCheque = 2
   'Const afxProductoTransferencia = 3
   'Const afxProductoTraveler = 4
   'Const afxProductoGiro = 5
   'Const afxProductoTurismo = 6
   'Const afxProductoCuentaCorriente = 7
   'Const afxProductoBanco = 8
   'Const afxProductoDeposito = 9
   'Const afxProductoMCBOPGirada = 10
   'Const afxProductoMCBChequeDevuelto = 11
   'Const afxProductoMCBCargo = 12
   'Const afxProductoMCBOPDepositada = 13
   'Const afxProductoMCBChqPorPagar = 14
   'Const afxProductoComisiones = 15
   'Const afxProductoPagoSPCaja = 16
   'Const afxProductoPagoSPDeposito = 17
   'Const afxProductoPagoSPCheque = 18
   
   Const afxAgente = 0
   Const afxCliente = 1
   Const afxPrincipal = 2

	Const afxTrfEnviadas = 9
	Const afxCompraVenta = 10
	Const afxListaGirosPendiantes = 1
	Const afxListaGirosRecibidos = 2
	Const afxListaGirosEnviados = 3
	Const afxListaGirosCartola = 5
	Const afxListaGirosCodigo = 7
	
	Const afxHTTP = "http://192.168.111.13/AfexMoneyWeb/"
	
	Const afxEstadoGiroCaptacion = 1
	Const afxEstadoGiroEnvio = 2
	Const afxEstadoGiroConfPagoAGP = 3
	Const afxEstadoGiroAviso = 4
	Const afxEstadoGiroPago = 5
	Const afxEstadoGiroConfPagoAGC = 6
	Const afxEstadoGiroReclamo = 7
	Const afxEstadoGiroSolucion = 8
	Const afxEstadoGiroNulo = 9
	'Pruebas
	'ok
	
	Const afxTrfNulo = 0
    Const afxTrfCorrelativo = 1
    Const afxTrfNumero = 2
    Const afxTrfCliente = 3
    Const afxTrfSP = 4
    Const afxTrfVoucher = 5
    Const Where = 6
%>