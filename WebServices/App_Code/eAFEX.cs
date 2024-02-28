using System;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Collections;
using System.Web.Services;
using System.Web.Services.Protocols;
using WebServidorC;

[WebService(Namespace = "http://www.afex.cl/", Description = "Web Services AFEX")]
public class eAFEX : System.Web.Services.WebService {
    clsconexion clsconn = new clsconexion();
    SqlConnection oconn;
    [WebMethod(Description = "Valores Para el Tipo de Cambio")]
    public Valores[] Obtenervalores()
    {
        DataSet dsvalores = new DataSet();
        // Jonathan Miranda G. 29-01-2007
        //String queryval = " select * from ( " +
        //                  " select a.codigo_moneda codigo_moneda,alias_moneda," +
        //                   " convert(varchar,convert(int," +
        //                             " case a.codigo_moneda " +
        //                               " when 'USD' then tipo_cambio_compra - 2 " +
        //                               " when 'EUR' then tipo_cambio_compra - 2 " +
        //                               " else tipo_cambio_compra " +
        //                              " end )) tipo_cambio_compra," +
        //                    " convert(varchar,convert(int," +
        //                             " case a.codigo_moneda " +
        //                               " when 'USD' then tipo_cambio_venta + 2 " +
        //                               " when 'EUR' then tipo_cambio_venta + 2 " +
        //                               " else tipo_cambio_venta " +
        //                             " end )) tipo_cambio_venta," +
        //                         " case a.codigo_moneda " +
        //                            " when 'USD' then 1 /* Dolar Americano */ " +
        //                            " when 'EUR' then 2 /* Euro */ " +
        //                            " when 'BRL' then 3 /* Real Brasil */ " +
        //                            " when 'ARP' then 4 /* Peso Argentino */ " +
        //                            " when 'AUD' then 5 /* Dolar Australiano */ " +
        //                            " when 'CAD' then 6 /* Dolar Canadiense */ " +
        //                            " when 'CHF' then 7 /* Franco Suizo */ " +
        //                            " when 'SEK' then 8 /* Corona Sueca */ " +
        //                            " when 'BOB' then 9 /* Peso Boliviano */ " +
        //                            " when 'PEN' then 10 /* Sol Peru */ " +
        //                            " when 'MXP' then 11 /* Peso Mexicano */ " +
        //                            " when 'ARP' then 12 /* Peso Argentino Se Repite */ " +
        //                            " when 'NOK' then 13 /* Corona Noruega */ " +
        //                            " when 'DKK' then 14 /* Corona Danesa */ " +
        //                            " when 'UKT' then 15 /* Libra Esterlina */" +
        //                            " when 'BEF' then 16 /* Franco Belga */ " +
        //                            " else 0 " +
        //                    " end orden " +
        //                    " from plan_moneda a,moneda b " +
        //                    " where a.codigo_moneda = b.codigo_moneda and " +
        //                    " a.codigo_producto = 1 and a.estado_plan_moneda = 1 and " +
        //                    " (b.shower_moneda > 0 or b.shower_grilla > 0)) " +
        //                    " valores order by orden /* shower_moneda desc ,shower_grilla */ ";
        String queryval = " execute obtenervalores ";
        //---------------------------- Fin --------------------------------

        int i = 0;

        oconn = clsconn.Abrir_conexion;
        SqlDataAdapter adpvalores = new SqlDataAdapter(queryval, oconn);
        adpvalores.Fill(dsvalores, "valores");

        DataTable dtabla = dsvalores.Tables["valores"];

        Valores[] datos = new Valores[dtabla.Rows.Count];

        foreach (DataRow fila in dtabla.Rows)
        {
            datos[i] = new Valores();
            datos[i].TipoCambio = (String)fila["codigo_moneda"];
            datos[i].Moneda = (String)fila["alias_moneda"];
            datos[i].Vcompra = (String)fila["tipo_cambio_compra"];
            datos[i].Vventa = (String)fila["tipo_cambio_venta"];
            i++;
        }
        oconn.Close();
        return datos;
    }
}

