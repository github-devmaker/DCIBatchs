using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
public class SqlConnectDB
{

    private bool useDB = true;
    private string connStr = "";


    public SqlConnectDB(string _DB)
    {
        if (_DB == "dbSCM")
        {
            connStr = "Data Source=192.168.226.86;Initial Catalog=dbSCM; Persist Security Info=True; User ID=sa;Password=decjapan; TrustServerCertificate=True;";
        }
        else if (_DB == "dbDCI")
        {
            connStr = "Data Source=COSTY;Initial Catalog=dbDCI; Persist Security Info=True; User ID=sa;Password=decjapan; TrustServerCertificate=True;";
        }
        else if (_DB == "dbBCS")
        {
            connStr = "Data Source=192.168.226.86;Initial Catalog=dbBCS; Persist Security Info=True; User ID=sa;Password=decjapan; TrustServerCertificate=True;";
        }
        else if (_DB == "dbIoT")
        {
            connStr = "Data Source=COSTY;Initial Catalog=dbIoT; Persist Security Info=True; User ID=sa;Password=decjapan; TrustServerCertificate=True;";
        }
        else if (_DB == "dbIoTFac2")
        {
            connStr = "Data Source=COSTY;Initial Catalog=dbIoTFac2; Persist Security Info=True; User ID=sa;Password=decjapan; TrustServerCertificate=True;";
        }
        else if (_DB == "dbIoTFac3")
        {
            connStr = "Data Source=10.194.40.103;Initial Catalog=dbIoT; Persist Security Info=True; User ID=sa;Password=decjapan;TrustServerCertificate=True;";
        }
        else if (_DB == "dbHRM")
        {
            connStr = "Data Source=192.168.226.86;Initial Catalog=dbHRM; Persist Security Info=True; User ID=sa;Password=decjapan; TrustServerCertificate=True;";
        }
        else if (_DB == "dbIoTL8")
        {
            connStr = "Data Source=10.194.122.103;Initial Catalog=dbIoT; Persist Security Info=True; User ID=sa;Password=decjapan;TrustServerCertificate=True;";
        }

        //connStr = _configuration.GetConnectionString("SCMConnStr").ToString();

    }

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
            SqlConnection conn = new SqlConnection(connStr);
            SqlDataAdapter adapter = new SqlDataAdapter(queryStr, conn);
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

    /// <summary>
    /// Query table by string and return table 
    /// </summary>
    /// <param name="commandDb">CommandDB for query</param>
    /// <returns>DataTable</returns>
    /// <remarks></remarks>
    public DataTable Query(SqlCommand commandDb)
    {
        if (useDB)
        {
            SqlConnection conn = new SqlConnection(connStr);
            commandDb.Connection = conn;
            SqlDataAdapter adapter = new SqlDataAdapter(commandDb);
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

        //=================================================================================
    }

    /// <summary>
    /// Execute คำสั่ง sql
    /// </summary>
    /// <param name="exeStr">String of sql</param>
    /// <remarks></remarks>
    public int ExecuteNonCommand(SqlCommand commandDb)
    {
        int result = 0;
        if (useDB)
        {
            SqlConnection conn = new SqlConnection(connStr);
            try
            {
                commandDb.Connection = conn;
                conn.Open();
                result = commandDb.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                conn.Close();
            }
        }
        return result;
    }
    public void ExecuteCommand(string exeStr)
    {
        if (useDB)
        {
            SqlConnection conn = new SqlConnection(connStr);
            try
            {
                SqlCommand commandDb = new SqlCommand(exeStr, conn);
                ExecuteCommand(commandDb);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                conn.Close();
            }

        }

    }

    /// <summary>
    /// ExecuteCommand
    /// </summary>
    /// <param name="commandDb">commanddb for execute</param>
    /// <remarks></remarks>
    public void ExecuteCommand(SqlCommand commandDb)
    {
        if (useDB)
        {
            SqlConnection conn = new SqlConnection(connStr);
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
            finally
            {
                conn.Close();
            }

        }
    }
    /// <summary>
    /// Execute หลายๆคำสั่ง พร้อมการ Rollback เมื่อคำสั่งไม่สำเร็จทั้งชุด
    /// </summary>
    /// <param name="exeStr"></param>
    /// <remarks></remarks>
    public void ExecuteCommand(List<string> exeStr)
    {
        if (useDB)
        {
            SqlConnection conn = new SqlConnection(connStr);
            SqlCommand commandDb = new SqlCommand();
            conn.Open();

            SqlTransaction transaction = null;
            // Start a local transaction
            transaction = conn.BeginTransaction("SampleTransaction");

            commandDb.Connection = conn;
            commandDb.Transaction = transaction;

            try
            {
                for (int index = 0; index <= exeStr.Count - 1; index++)
                {
                    commandDb.CommandText = exeStr[index];
                    commandDb.ExecuteNonQuery();

                }

                //Commit
                transaction.Commit();

            }
            catch (Exception ex)
            {
                //Rollback
                transaction.Rollback();
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                conn.Close();
            }
        }
    }

    /// <summary>
    /// Execute หลายๆคำสั่ง พร้อมการ Rollback เมื่อคำสั่งไม่สำเร็จทั้งชุด
    /// </summary>
    /// <param name="commandDb"></param>
    /// <remarks></remarks>
    public void ExecuteCommand(List<SqlCommand> commandDb)
    {
        if (useDB)
        {
            SqlConnection conn = new SqlConnection(connStr);
            conn.Open();

            SqlTransaction transaction = null;
            // Start a local transaction
            transaction = conn.BeginTransaction("SampleTransaction");

            try
            {

                for (int index = 0; index <= commandDb.Count - 1; index++)
                {
                    commandDb[index].Connection = conn;
                    commandDb[index].Transaction = transaction;
                    commandDb[index].ExecuteNonQuery();

                }

                //Commit
                transaction.Commit();
            }
            catch (Exception ex)
            {
                //Rollback
                transaction.Rollback();
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                conn.Close();
            }
        }
    }

}