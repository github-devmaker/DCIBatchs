using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.OracleClient;

public class OraConnectDB
{
    private string connStr = "";
    //static string constr = @"Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.214.123)(PORT = 1531)) (CONNECT_DATA = (SID = DCIOS01)));User Id=master;Password=master";

    public OraConnectDB(string DBSource)
    {
        //
        // TODO: Add constructor logic here
        //
        if (DBSource == "ALPHA01")
        {
            connStr = @"Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.214.123)(PORT = 1531)) (CONNECT_DATA = (SID = DCIOS01)));User Id=plan;Password=plan;";

        }
        else if (DBSource == "ALPHA02")
        {
            connStr = @"Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.214.124)(PORT = 1532)) (CONNECT_DATA = (SID = DCIOS02)));User Id=mc;Password=mc;";

        }
        else if (DBSource == "ALPHAPD")
        {
            connStr = @"Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.215.193)(PORT = 1521)) (CONNECT_DATA = (SID = DCIPD)));User Id=se;Password=isse;";

        }
    }

    private bool useDB = true;
    //private string connStr = System.Configuration.ConfigurationSettings.AppSettings["DesignMachineMonitoring.Properties.Settings.ETD_YCConnectionString"];
    //private string connStr = "Data Source=costy;Initial Catalog=dbDCI; Persist Security Info=True; User ID=sa;Password=decjapan;";


    //Property ObjectManages As Object

    /// <summary>
    /// Query table by string and return table 
    /// </summary>
    /// <param name="queryStr">String of sql query</param>
    /// <returns>DataTable</returns>
    /// <remarks></remarks>
    public DataTable Query(string queryStr)
    {
        if (useDB)
        {
            OracleConnection conn = new OracleConnection(connStr);
            OracleDataAdapter adapter = new OracleDataAdapter(queryStr, conn);
            DataTable dTable = new DataTable();
            DataSet dSet = new DataSet();

            try
            {
                adapter.Fill(dSet, "dataTable");
                return dSet.Tables["dataTable"];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return dTable;
            }
            finally
            {
                conn.Close();
            }

        }
        else
        {
            return new DataTable();
        }

    } 
    public DataTable Query(OracleCommand commandDb)
    {
        if (useDB)
        {
            OracleConnection conn = new OracleConnection(connStr);
            commandDb.Connection = conn;
            OracleDataAdapter adapter = new OracleDataAdapter(commandDb);
            DataTable dTable = new DataTable();
            DataSet dSet = new DataSet();

            try
            {
                adapter.Fill(dSet, "dataTable");
                return dSet.Tables["dataTable"];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return dTable;
            }
            finally
            {
                conn.Close();
            }

        }
        else
        {
            return new DataTable();
        } 
    } 

    public void ExecuteCommand(string exeStr)
    {
        if (useDB)
        {
            OracleConnection conn = new OracleConnection(connStr);
            try
            {
                OracleCommand commandDb = new OracleCommand(exeStr, conn);
                ExecuteCommand(commandDb);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

    } 
    public void ExecuteCommand(OracleCommand commandDb)
    {
        if (useDB)
        {
            OracleConnection conn = new OracleConnection(connStr);
            try
            {
                commandDb.Connection = conn;
                conn.Open();
                commandDb.ExecuteNonQuery();
                conn.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }
    }


}