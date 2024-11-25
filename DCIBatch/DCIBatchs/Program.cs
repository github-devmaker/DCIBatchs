using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Data.OracleClient;
using System.Globalization;
using DCIBatchs.Services;

namespace DCIBatchs
{
    internal class Program
    {
        public static IniFile ini = new IniFile("Config.ini");
        public static DateTime dtNow = DateTime.Now;
        public static string ym = dtNow.ToString("yyyyMM");
        public static string ymd = dtNow.ToString("yyyyMMdd");
        public static string configPath = "config.txt";
        public static Dictionary<string, string> config = new Dictionary<string, string>();
        public static Helper oHelper = new Helper();
        public static SqlConnectDB MSSCM = new SqlConnectDB("dbSCM");
        public static OraConnectDB OraPD2 = new OraConnectDB("ALPHA02");
        static void Main(string[] args)
        {
            string PULL_STOCK_PART_SUPPLY_REALTIME = ini.GetString("METHOD", "PULL_STOCK_PART_SUPPLY_REALTIME", "FALSE");
            string PULL_STOCK_PART_SUPPLY_MORNING = ini.GetString("METHOD", "PULL_STOCK_PART_SUPPLY_MORNING", "FALSE");
            string PULL_STOCK_UKEHARAI = ini.GetString("METHOD", "PULL_STOCK_UKEHARAI", "FALSE");

            if (PULL_STOCK_PART_SUPPLY_REALTIME == "TRUE")
            {
                Console.WriteLine($"METHOD RUN : [PULL_STOCK_PART_SUPPLY_REALTIME] Date : {dtNow.ToString("dd/MM/yyyy HH:mm:ss")}");
                PullStockPSRealtime();
                // ================= FOR PROGRAM ===================== //
                // ======= 1. APS ==================================== // 
                // ==================== END ========================== //
            }
            else if (PULL_STOCK_PART_SUPPLY_MORNING == "TRUE")
            {
                Console.WriteLine($"METHOD RUN : PULL_STOCK_PART_SUPPLY_MORNING Date : {dtNow.ToString("dd/MM/yyyy HH:mm:ss")}");
                PullStockPSMorning();
                // ================= FOR PROGRAM ===================== //
                // ======= 1. Delivery Order  ======================== //
                // ======= 2. Ukeharai =============================== //
                // ==================== END ========================== //
            }
            else if (PULL_STOCK_UKEHARAI == "TRUE")
            {
                Console.WriteLine($"METHOD RUN : PULL_STOCK_UKEHARAI Date : {dtNow.ToString("dd/MM/yyyy HH:mm:ss")}");
                PullStockUkeharai();
                // ================= FOR PROGRAM ===================== //
                // ======= 1. Warning ================================ //
                // ======= 2. Ukeharai =============================== //
                // ==================== END ========================== //
            }
            else
            {
                Console.WriteLine("BATCH NO KEY !");
                Console.ReadKey();
            }
        }

        private static void PullStockUkeharai()
        {
            UkeharaiService UkeService = new UkeharaiService();

            DataTable dtWmsInventory = UkeService.WmsGetInventoryIVW01(ym);
            DataTable dtWmsAssortInOut = UkeService.WmsGetAssortInOut(ym);
            //DataTable dtWmsTransferInOut = UkeService.WmsGetTransferInOut(ym);
            DataTable dtWmsSaleExport = UkeService.WmsGetSaleExport(ym);
            DataTable dtWmsSales = UkeService.WmsGetSales(ym);
            DataTable dtWmsSaleDomestic = UkeService.WmsGetSaleDomestic(ym);
            DataTable dtModels = UkeService.ScmGetModels();

            string strDel = $"DELETE FROM UKE_INITIAL_STOCK_DCI_OF_DAY WHERE YMD = '{ymd}' ";
            SqlCommand cmdDel = new SqlCommand();
            cmdDel.CommandText = strDel;
            MSSCM.ExecuteCommand(cmdDel);


            int err = 0, rec = 0;
            foreach (DataRow dr in dtModels.Rows)
            {
                string model = dr["MODEL"].ToString().Trim();
                string sebango = dr["SEBANGO"].ToString();
                decimal Inv = 0;
                try
                {
                    Inv = dtWmsInventory.AsEnumerable().Where(x => x.Field<string>("MODEL").Trim() == model).Sum(x => x.Field<decimal>("LBALSTK"));
                    decimal CntAssetIn = dtWmsAssortInOut.AsEnumerable().Where(x => DateTime.ParseExact(x.Field<string>("ASTDATE"), "yyyy-MM-dd", CultureInfo.InvariantCulture).Date <= dtNow.Date && x.Field<string>("ASTTYPE") == "IN" && x.Field<string>("MODEL").Trim() == model).Sum(x => x.Field<decimal>("ASTQTY"));
                    decimal CntSaleDomestic = dtWmsSaleDomestic.AsEnumerable().Where(x => DateTime.ParseExact(x.Field<string>("DELDATE"), "yyyyMMdd", CultureInfo.InvariantCulture).Date <= dtNow.Date && x.Field<string>("MODEL") == model).Sum(x => x.Field<decimal>("PICQTY"));
                    decimal CntSales = dtWmsSales.AsEnumerable().Where(x => Convert.ToInt32(x.Field<string>("LOADDATE")) <= Convert.ToInt32(dtNow.ToString("yyyyMMdd")) && x.Field<string>("MODEL") == model).Sum(x => x.Field<decimal>("PICQTY"));
                    decimal CntAssetOut = dtWmsAssortInOut.AsEnumerable().Where(x => DateTime.ParseExact(x.Field<string>("ASTDATE"), "yyyy-MM-dd", CultureInfo.InvariantCulture).Date <= dtNow.Date && x.Field<string>("ASTTYPE") == "OUT" && x.Field<string>("MODEL").Trim() == model).Sum(x => x.Field<decimal>("ASTQTY"));

                    decimal CntCurrentInv = (Inv + CntAssetIn) - (CntSales + CntAssetOut);

                    string strInstr = $"INSERT INTO  UKE_INITIAL_STOCK_DCI_OF_DAY  ([YMD],[MODEL],[INVENTORY],[CREATE_DATE],[CREATE_BY]) VALUES ('{ymd}', '{model}','{CntCurrentInv}',CURRENT_TIMESTAMP, 'BATCH')  ";
                    SqlCommand cmdInstr = new SqlCommand();
                    cmdInstr.CommandText = strInstr;
                    int res = MSSCM.ExecuteNonCommand(cmdInstr);
                    if (res <= 0)
                    {
                        err++;
                    }
                }
                catch (Exception ex)
                {
                    err++;
                    Console.WriteLine($"Error Model : {model} message : {ex.Message}");
                }
                rec++;
            }
            if (err > 0)
            {
                Console.WriteLine($"Error Insert or Update [PULL_STOCK_PART_SUPPLY_MORNING]");
                Console.ReadKey();
            }
        }
        private static void PullStockPSMorning()
        {
            bool Result = true;
            string _Year = dtNow.Year.ToString();
            string _Month = dtNow.Month.ToString("00");
            string date = _Year + "" + _Month;
            OracleCommand cmd = new OracleCommand();
            //PARTS = PARTS != "" ? PARTS : _PART_JOIN_STRING;
            cmd.CommandText = @"SELECT '" + dtNow.ToString("yyyyMMdd") + "',MC1.PARTNO, MC1.CM, DECODE(SB1.DSBIT,'1','OBSOLETE','2','DEAD STOCK','3',CASE WHEN TRIM(SB1.STOPDATE) IS NOT NULL AND SB1.STOPDATE <= TO_CHAR(SYSDATE,'YYYYMMDD') THEN 'NOT USE ' || SB1.STOPDATE ELSE ' ' END, ' ') PART_STATUS, MC1.LWBAL, NVL(AC1.ACQTY,0) AS ACQTY, NVL(PID.ISQTY,0) AS ISQTY, MC1.LWBAL + NVL(AC1.ACQTY,0) - NVL(PID.ISQTY,0) AS WBAL,NVL(RT3.LREJ,0) + NVL(PID.REJIN,0) - NVL(AC1.REJOUT,0) AS REJQTY, MC2.QC, MC2.WH1, MC2.WH2, MC2.WH3, MC2.WHA, MC2.WHB, MC2.WHC, MC2.WHD, MC2.WHE,ZUB.HATANI AS UNIT, EPN.KATAKAN AS DESCR, F_GET_HTCODE_RATIO(MC1.JIBU,MC1.PARTNO, '" + dtNow.ToString("yyyyMMdd") + "') AS HTCODE, F_GET_MSTVEN_VDABBR(MC1.JIBU,F_GET_HTCODE_RATIO(MC1.JIBU,MC1.PARTNO,'" + dtNow.ToString("yyyyMMdd") + "')) SUPPLIER, SB1.LOCA1, SB1.LOCA2, SB1.LOCA3, SB1.LOCA4, SB1.LOCA5, SB1.LOCA6, SB1.LOCA7, SB1.LOCA8 FROM	(SELECT	* FROM	DST_DATMC1 WHERE TRIM(YM) = :YM  AND CM LIKE '%'";
            cmd.CommandText = cmd.CommandText + @") MC1, 
        		(SELECT	PARTNO, CM, SUM(WQTY) AS ACQTY, SUM(CASE WHEN WQTY < 0 THEN -1 * WQTY ELSE 0 END) AS REJOUT 
        		 FROM	DST_DATAC1 
        		 WHERE	ACDATE >= :DATE_START 
        			AND	ACDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%'  AND CM LIKE '%'
        		 GROUP BY PARTNO, CM 
        		) AC1, 
        		(SELECT	PARTNO, BRUSN AS CM, SUM(FQTY) AS ISQTY, SUM(DECODE(REJBIT,'R',-1*FQTY,0)) AS REJIN 
        		 FROM	MASTER.GST_DATPID@ALPHA01 
        		 WHERE	IDATE >= :DATE_START 
        			AND	IDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%'  AND BRUSN LIKE '%'
        		 GROUP BY PARTNO, BRUSN 
        		) PID, 
        		(SELECT    PARTNO, CM, SUM(DECODE(WHNO,'QC',BALQTY)) AS QC,SUM(DECODE(WHNO,'W1',BALQTY)) AS WH1,SUM(DECODE(WHNO,'W2',BALQTY)) AS WH2,SUM(DECODE(WHNO,'W3',BALQTY)) AS WH3, 
                           SUM(DECODE(WHNO,'WA',BALQTY)) AS WHA,SUM(DECODE(WHNO,'WB',BALQTY)) AS WHB,SUM(DECODE(WHNO,'WC',BALQTY)) AS WHC,SUM(DECODE(WHNO,'WD',BALQTY)) AS WHD,SUM(DECODE(WHNO,'WE',BALQTY)) AS WHE 
                    FROM    (SELECT    MC2.PARTNO, MC2.CM, MC2.WHNO, MC2.LWBAL, NVL(AC1.ACQTY,0) AS ACQTY, NVL(PID.ISQTY,0) AS ISQTY, MC2.LWBAL + NVL(AC1.ACQTY,0) - NVL(PID.ISQTY,0) AS BALQTY 
                            FROM    (SELECT    * 
                                    FROM    DST_DATMC2 
                                    WHERE    YM = :YM  AND TRIM(PARTNO) LIKE '%' AND CM LIKE '%'
                                   ) MC2, 
                                   (SELECT    PARTNO, CM, WHNO, SUM(WQTY) AS ACQTY 
                                    FROM    DST_DATAC1 
                                    WHERE    ACDATE >= :DATE_START 
                                       AND    ACDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%' AND CM LIKE '%'
                                    GROUP BY PARTNO, CM, WHNO 
                                   ) AC1, 
                                   (SELECT    PARTNO, BRUSN AS CM, WHNO, SUM(FQTY) AS ISQTY 
                                    FROM    (SELECT    * 
                                             FROM    MASTER.GST_DATPID@ALPHA01 
                                             WHERE    IDATE >= :DATE_START 
                                               AND    IDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%' AND BRUSN LIKE '%'
                                            UNION ALL 
                                             SELECT    * 
                                             FROM    DST_DATPID3 
                                             WHERE    IDATE >= :DATE_START 
                                               AND    IDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%' AND BRUSN LIKE '%'
                                           ) 
                                    GROUP BY PARTNO, BRUSN, WHNO 
                                   ) PID 
                            WHERE    MC2.PARTNO    = AC1.PARTNO(+) 
                               AND    MC2.CM        = AC1.CM(+) 
                               AND    MC2.WHNO    = AC1.WHNO(+) 
                               AND    MC2.PARTNO    = PID.PARTNO(+) 
                               AND    MC2.CM        = PID.CM(+) 
                               AND    MC2.WHNO    = PID.WHNO(+) 
                           ) 
                    GROUP BY PARTNO, CM 
                   ) MC2, 
                   MASTER.ND_EPN_TBL_V1@ALPHA01 EPN, DST_MSTSB1 SB1, MASTER.ND_ZUB_TBL@ALPHA01 ZUB, DST_DATRT3 RT3 
           WHERE    MC1.PARTNO    = AC1.PARTNO(+) 
               AND    MC1.CM        = AC1.CM(+) 
               AND    MC1.PARTNO    = PID.PARTNO(+) 
               AND    MC1.CM        = PID.CM(+) 
               AND    MC1.YM        = RT3.YM(+) 
               AND    MC1.PARTNO    = RT3.PARTNO(+) 
               AND    MC1.CM        = RT3.CM(+) 
               AND    MC1.PARTNO    = EPN.PARTNO(+) 
               AND    MC1.PARTNO    = SB1.PARTNO(+) 
               AND    MC1.CM        = SB1.CM(+) 
               AND    MC1.PARTNO    = MC2.PARTNO(+) 
               AND    MC1.CM        = MC2.CM(+) 
               AND    MC1.PARTNO    = ZUB.PARTNO(+) 
               AND    ZUB.STRYMN(+) <= :DATE_START 
               AND    ZUB.ENDYMN(+) >  :DATE_RUN 
               AND    ZUB.KSNBIT(+) <> '2'";
            cmd.Parameters.Add(new OracleParameter(":YM", date));
            cmd.Parameters.Add(new OracleParameter(":DATE_START", dtNow.ToString("yyyyMM01")));
            cmd.Parameters.Add(new OracleParameter(":DATE_RUN", dtNow.ToString("yyyyMMdd")));
            DataTable dt = OraPD2.Query(cmd);
            if (dt.Rows.Count > 0)
            {
                SqlCommand SqlDelDOStockAlphaOfDay = new SqlCommand();
                SqlDelDOStockAlphaOfDay.CommandText = $@"DELETE FROM [dbSCM].[dbo].[DO_STOCK_ALPHA] WHERE DATE_PD = '{dtNow.ToString("yyyy-MM-dd")}'";
                MSSCM.Query(SqlDelDOStockAlphaOfDay);
            }
            foreach (DataRow dr in dt.Rows)
            {
                string PartNo = dr["PARTNO"].ToString().Trim();
                string Cm = dr["CM"].ToString().Trim();
                double Bal = oHelper.ConvStrToDB(dr["WBAL"].ToString());
                string VdCode = dr["HTCODE"].ToString().Trim();
                SqlCommand sqlInsert = new SqlCommand();
                sqlInsert.CommandText = $@"INSERT INTO [dbSCM].[dbo].[DO_STOCK_ALPHA] ([DATE_PD] ,[PARTNO] ,[CM] ,[VDCODE] ,[STOCK] ,[REV] ,[INSERT_DT]) VALUES ('{dtNow.ToString("yyyy-MM-dd")}','{PartNo}','{Cm}','{Bal}','{VdCode}','999',GETDATE())";
                int Insert = MSSCM.ExecuteNonCommand(sqlInsert);
                if (Insert <= 0)
                {
                    Console.WriteLine($"Insert Error  Vender : {VdCode}, PartNo : {PartNo}, CM : {Cm}, Stock : {Bal}");
                    Result = false;
                }
            }
            if (Result == false)
            {
                Console.WriteLine($"บันทึกข้อมูลไม่สำเร็จ Method : [PULL_STOCK_PART_SUPPLY_MORNING]");
                Console.ReadKey();
            }
        }

        private static void PullStockPSRealtime()
        {
            string _Year = dtNow.Year.ToString();
            string _Month = dtNow.Month.ToString("00");
            string date = _Year + "" + _Month;
            OracleCommand cmd = new OracleCommand();
            cmd.CommandText = @"SELECT '" + dtNow.ToString("yyyyMMdd") + "',MC1.PARTNO, MC1.CM, DECODE(SB1.DSBIT,'1','OBSOLETE','2','DEAD STOCK','3',CASE WHEN TRIM(SB1.STOPDATE) IS NOT NULL AND SB1.STOPDATE <= TO_CHAR(SYSDATE,'YYYYMMDD') THEN 'NOT USE ' || SB1.STOPDATE ELSE ' ' END, ' ') PART_STATUS, MC1.LWBAL, NVL(AC1.ACQTY,0) AS ACQTY, NVL(PID.ISQTY,0) AS ISQTY, MC1.LWBAL + NVL(AC1.ACQTY,0) - NVL(PID.ISQTY,0) AS WBAL,NVL(RT3.LREJ,0) + NVL(PID.REJIN,0) - NVL(AC1.REJOUT,0) AS REJQTY, MC2.QC, MC2.WH1, MC2.WH2, MC2.WH3, MC2.WHA, MC2.WHB, MC2.WHC, MC2.WHD, MC2.WHE,ZUB.HATANI AS UNIT, EPN.KATAKAN AS DESCR, F_GET_HTCODE_RATIO(MC1.JIBU,MC1.PARTNO, '" + dtNow.ToString("yyyyMMdd") + "') AS HTCODE, F_GET_MSTVEN_VDABBR(MC1.JIBU,F_GET_HTCODE_RATIO(MC1.JIBU,MC1.PARTNO,'" + dtNow.ToString("yyyyMMdd") + "')) SUPPLIER, SB1.LOCA1, SB1.LOCA2, SB1.LOCA3, SB1.LOCA4, SB1.LOCA5, SB1.LOCA6, SB1.LOCA7, SB1.LOCA8 FROM	(SELECT	* FROM	DST_DATMC1 WHERE	TRIM(YM) = :YM  AND CM LIKE '%'";
            cmd.CommandText = cmd.CommandText + @") MC1, 
        		(SELECT	PARTNO, CM, SUM(WQTY) AS ACQTY, SUM(CASE WHEN WQTY < 0 THEN -1 * WQTY ELSE 0 END) AS REJOUT 
        		 FROM	DST_DATAC1 
        		 WHERE	ACDATE >= :DATE_START 
        			AND	ACDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%'  AND CM LIKE '%'
        		 GROUP BY PARTNO, CM 
        		) AC1, 
        		(SELECT	PARTNO, BRUSN AS CM, SUM(FQTY) AS ISQTY, SUM(DECODE(REJBIT,'R',-1*FQTY,0)) AS REJIN 
        		 FROM	MASTER.GST_DATPID@ALPHA01 
        		 WHERE	IDATE >= :DATE_START 
        			AND	IDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%'  AND BRUSN LIKE '%'
        		 GROUP BY PARTNO, BRUSN 
        		) PID, 
        		(SELECT    PARTNO, CM, SUM(DECODE(WHNO,'QC',BALQTY)) AS QC,SUM(DECODE(WHNO,'W1',BALQTY)) AS WH1,SUM(DECODE(WHNO,'W2',BALQTY)) AS WH2,SUM(DECODE(WHNO,'W3',BALQTY)) AS WH3, 
                           SUM(DECODE(WHNO,'WA',BALQTY)) AS WHA,SUM(DECODE(WHNO,'WB',BALQTY)) AS WHB,SUM(DECODE(WHNO,'WC',BALQTY)) AS WHC,SUM(DECODE(WHNO,'WD',BALQTY)) AS WHD,SUM(DECODE(WHNO,'WE',BALQTY)) AS WHE 
                    FROM    (SELECT    MC2.PARTNO, MC2.CM, MC2.WHNO, MC2.LWBAL, NVL(AC1.ACQTY,0) AS ACQTY, NVL(PID.ISQTY,0) AS ISQTY, MC2.LWBAL + NVL(AC1.ACQTY,0) - NVL(PID.ISQTY,0) AS BALQTY 
                            FROM    (SELECT    * 
                                    FROM    DST_DATMC2 
                                    WHERE    YM = :YM  AND TRIM(PARTNO) LIKE '%' AND CM LIKE '%'
                                   ) MC2, 
                                   (SELECT    PARTNO, CM, WHNO, SUM(WQTY) AS ACQTY 
                                    FROM    DST_DATAC1 
                                    WHERE    ACDATE >= :DATE_START 
                                       AND    ACDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%' AND CM LIKE '%'
                                    GROUP BY PARTNO, CM, WHNO 
                                   ) AC1, 
                                   (SELECT    PARTNO, BRUSN AS CM, WHNO, SUM(FQTY) AS ISQTY 
                                    FROM    (SELECT    * 
                                             FROM    MASTER.GST_DATPID@ALPHA01 
                                             WHERE    IDATE >= :DATE_START 
                                               AND    IDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%' AND BRUSN LIKE '%'
                                            UNION ALL 
                                             SELECT    * 
                                             FROM    DST_DATPID3 
                                             WHERE    IDATE >= :DATE_START 
                                               AND    IDATE <= :DATE_RUN  AND TRIM(PARTNO) LIKE '%' AND BRUSN LIKE '%'
                                           ) 
                                    GROUP BY PARTNO, BRUSN, WHNO 
                                   ) PID 
                            WHERE    MC2.PARTNO    = AC1.PARTNO(+) 
                               AND    MC2.CM        = AC1.CM(+) 
                               AND    MC2.WHNO    = AC1.WHNO(+) 
                               AND    MC2.PARTNO    = PID.PARTNO(+) 
                               AND    MC2.CM        = PID.CM(+) 
                               AND    MC2.WHNO    = PID.WHNO(+) 
                           ) 
                    GROUP BY PARTNO, CM 
                   ) MC2, 
                   MASTER.ND_EPN_TBL_V1@ALPHA01 EPN, DST_MSTSB1 SB1, MASTER.ND_ZUB_TBL@ALPHA01 ZUB, DST_DATRT3 RT3 
           WHERE    MC1.PARTNO    = AC1.PARTNO(+) 
               AND    MC1.CM        = AC1.CM(+) 
               AND    MC1.PARTNO    = PID.PARTNO(+) 
               AND    MC1.CM        = PID.CM(+) 
               AND    MC1.YM        = RT3.YM(+) 
               AND    MC1.PARTNO    = RT3.PARTNO(+) 
               AND    MC1.CM        = RT3.CM(+) 
               AND    MC1.PARTNO    = EPN.PARTNO(+) 
               AND    MC1.PARTNO    = SB1.PARTNO(+) 
               AND    MC1.CM        = SB1.CM(+) 
               AND    MC1.PARTNO    = MC2.PARTNO(+) 
               AND    MC1.CM        = MC2.CM(+) 
               AND    MC1.PARTNO    = ZUB.PARTNO(+) 
               AND    ZUB.STRYMN(+) <= :DATE_START 
               AND    ZUB.ENDYMN(+) >  :DATE_RUN 
               AND    ZUB.KSNBIT(+) <> '2'";
            cmd.Parameters.Add(new OracleParameter(":YM", date));
            cmd.Parameters.Add(new OracleParameter(":DATE_START", dtNow.ToString("yyyyMM01")));
            cmd.Parameters.Add(new OracleParameter(":DATE_RUN", dtNow.ToString("yyyyMMdd")));
            DataTable dt = OraPD2.Query(cmd);
            DataRow[] drs = dt.Select("WBAL <> '0'  ");
            dt = drs.CopyToDataTable();
            DataTable dtAlphaTI2PartSupply = GetAlphaTI2PartSupply();
            foreach (DataRow dr in dt.Rows)
            {
                string Drawing = dr["PARTNO"].ToString().Trim();
                string cm = dr["CM"].ToString().Trim();
                int Inventory = oHelper.ConvStrToInt(dr["WBAL"].ToString());
                try
                {
                    var hasRow = dtAlphaTI2PartSupply.AsEnumerable().FirstOrDefault(x => x.Field<string>("Drawing") == Drawing && x.Field<string>("Cm") == cm);
                    if (hasRow != null)
                    {
                        SqlCommand sqlUpdate = new SqlCommand();
                        sqlUpdate.CommandText = $@"UPDATE  [dbo].[ALPHA_TI2_STOCK_PART_SUPPLY] SET Inventory = '{Inventory}', CM = '{cm}',UpdateDate = GETDATE() WHERE Drawing = '{Drawing}' AND CM = '{cm}'";
                        int update = MSSCM.ExecuteNonCommand(sqlUpdate);
                    }
                    else
                    {
                        SqlCommand sqlInsert = new SqlCommand();
                        sqlInsert.CommandText = $@"INSERT INTO [dbo].[ALPHA_TI2_STOCK_PART_SUPPLY] (Drawing,CM,Inventory,Area) VALUES ('{Drawing}','{cm}','{Inventory}','PART SUPPLY')";
                        int Insert = MSSCM.ExecuteNonCommand(sqlInsert);
                        if (Insert <= 0)
                        {
                            Console.WriteLine($"Insert Error : Drawing = {Drawing} Cm = {cm} Inventory = {Inventory}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error : Drawing = {Drawing} Cm = {cm} Inventory = {Inventory}    ===>  Catch : {ex.Message}  ");
                }
            }
        }

        public static Dictionary<string, string> LoadConfig(string path)
        {
            var config = new Dictionary<string, string>();

            if (File.Exists(path))
            {
                var lines = File.ReadAllLines(path);
                foreach (var line in lines)
                {
                    if (line.Contains("="))
                    {
                        var keyValue = line.Split(new char[] { '=' }, 2);
                        if (keyValue.Length == 2)
                        {
                            config[keyValue[0].Trim()] = keyValue[1].Trim();
                        }
                    }
                }
            }

            return config;
        }
        private static DataTable GetAlphaTI2PartSupply()
        {
            SqlCommand sql = new SqlCommand();
            sql.CommandText = $@"SELECT * FROM [dbSCM].[dbo].[ALPHA_TI2_STOCK_PART_SUPPLY]";
            return MSSCM.Query(sql);
        }
    }
}
