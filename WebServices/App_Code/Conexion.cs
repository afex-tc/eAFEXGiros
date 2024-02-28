using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;

public class clsconexion
{
    // Jonathan Miranda G. 29-01-2007
    //private String SConexion = "Server=canelo,1433;uid=cambios;pwd=cambios;Database=cambios_moneda;";
    private String SConexion = "Server=canelo,1433;uid=cambios;pwd=cambios;Database=cambios;";
    //---------------------------------- Fin -----------------------------------------    
    private SqlConnection oconn;
	public String String_Conexion{
        get{
            return SConexion;
        }
        set{
            SConexion = value;
        }
	}
    public SqlConnection Abrir_conexion{
        get{
            try{
                oconn = new SqlConnection(SConexion);
                oconn.Open();
            }
            catch{
                Console.WriteLine("Error en la conexión");
            }     
        return oconn;     
        }
    }
}
