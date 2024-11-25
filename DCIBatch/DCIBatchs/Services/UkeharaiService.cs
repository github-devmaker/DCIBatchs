using System;
using System.Data;
using System.Data.OracleClient;
using System.Data.SqlClient;
using System.Globalization;

namespace DCIBatchs.Services
{
    internal class UkeharaiService
    {
        SqlConnectDB MSSCM = new SqlConnectDB("dbSCM");
        private OraConnectDB _ALPHAPD = new OraConnectDB("ALPHAPD");

        public DataTable ScmGetModels()
        {
            SqlCommand sql = new SqlCommand();
            sql.CommandText = @"SELECT  ModelCode SEBANGO ,Model AS MODEL,SUBSTRING(ModelType,1,3) as FAC FROM [dbSCM].[dbo].[PN_Compressor]
              WHERE Status = 'ACTIVE' AND ModelType NOT IN ('PACKING','SPECIAL') AND LEN(ModelType) >= 3
               -- AND Trim(Model) = '1YC36DXD#A' 
              GROUP BY ModelCode,Model,ModelType ORDER BY Model ";
            DataTable dt = MSSCM.Query(sql);
            return dt;
        }
        internal DataTable WmsGetTransferInOut(string ym)
        {
            OracleCommand str = new OracleCommand();
            str.CommandText = $@"select to_char(h.recdate,'YYYYMMDD') RECYMD, d.trnno, d.model, d.pltype, d.recqty, h.fromwh, h.towh, case when fromwh='DCI' then 'OUT' else 'IN' end trntype 
                               from wms_trnctl h
                               left join wms_trndtl d on d.trnno = h.trnno  
                               where (h.fromwh = 'DCI' or h.towh = 'DCI') and h.trnsts= 'Transfered' and h.recsts = 'Received'
                               and to_char(h.recdate,'YYYYMMDD') LIKE '{ym}%'
                               order by h.recdate asc  ";
            DataTable dt = _ALPHAPD.Query(str);
            return dt;
        }

        internal DataTable WmsGetSaleExport(string ym)
        {
            OracleCommand str = new OracleCommand();
            str.CommandText = $@"SELECT H.IVNO,TO_CHAR(H.UDATE, 'yyyyMMdd') DELDATE,  W.MODEL,SUM(W.PICQTY) PICQTY  FROM SE.WMS_DELCTN W
                                            LEFT JOIN SE.WMS_DELCTL H ON H.COMID='DCI' AND H.IVNO = W.IVNO AND H.DONO = W.DONO 
                                            WHERE W.CFBIT = 'F' AND W.IFBIT IN ('F','U') AND TO_CHAR(H.DELDATE, 'yyyyMMdd') LIKE '{ym}%'
                                            AND TO_CHAR(W.UDATE, 'yyyyMMdd') LIKE '{ym}%'
                                            AND SUBSTR(H.IVNO,1,1) = 'E'
                                            GROUP BY H.IVNO,TO_CHAR(H.UDATE, 'yyyyMMdd') , W.MODEL  ";
            DataTable dt = _ALPHAPD.Query(str);
            return dt;
        }

        internal DataTable WmsGetSales(string ym)
        {
            OracleCommand str = new OracleCommand();
            str.CommandText = $@"WITH H AS (
                                    SELECT CTD.IVNO, CTD.DONO, TO_CHAR(MIN(CTD.LOADDATE), 'yyyyMMdd') LOADDATE
                                    FROM SE.WMS_DELCTD CTD
                                    WHERE TO_CHAR(CTD.LOADDATE, 'yyyyMMdd') LIKE '{ym}%'  
                                    GROUP BY CTD.IVNO, CTD.DONO
                                )
                                SELECT H.IVNO, H.DONO, H.LOADDATE, D.MODEL, SUM(D.PICQTY) PICQTY 
                                FROM H
                                LEFT JOIN SE.WMS_DELDTL D ON H.IVNO = D.IVNO AND H.DONO = D.DONO 
                                GROUP BY H.IVNO, H.DONO, H.LOADDATE, D.MODEL  ";
            DataTable dt = _ALPHAPD.Query(str);
            return dt;
        }

        internal DataTable WmsGetSaleDomestic(string ym)
        {
            OracleCommand str = new OracleCommand();
            str.CommandText = $@"SELECT TO_CHAR(H.DELDATE, 'yyyyMMdd') DELDATE,  
                                               W.MODEL, W.PLTYPE,   
                                               SUM(W.QTY) QTY, 
                                               SUM(W.ALQTY) ALQTY,   
                                               SUM(W.PICQTY) PICQTY   
                                            FROM SE.WMS_DELCTN W
                                            LEFT JOIN SE.WMS_DELCTL H ON H.COMID='DCI' AND H.IVNO = W.IVNO AND H.DONO = W.DONO 
                                            WHERE W.CFBIT = 'F' AND W.IFBIT = 'F' AND TO_CHAR(H.DELDATE, 'yyyyMMdd') LIKE '{ym}%'
                                            GROUP BY TO_CHAR(H.DELDATE, 'yyyyMMdd') , W.MODEL, W.PLTYPE ";
            DataTable dt = _ALPHAPD.Query(str);
            return dt;
        }

        internal DataTable WmsGetInventoryIVW01(string ym)
        {
            OracleCommand str = new OracleCommand();
            str.CommandText = $@" SELECT W.YM, 
                               W.MODEL, SUM(W.LBALSTK) LBALSTK, SUM(W.INSTK) INSTK, 
                               SUM(W.OUTSTK) OUTSTK, SUM(W.BALSTK) BALSTK
                               FROM SE.WMS_STKBAL W
                               WHERE comid = 'DCI' and ym = '{ym}' and wc in ('DCI', 'SKO')
                               GROUP BY W.YM, W.MODEL";
            DataTable dt = _ALPHAPD.Query(str);
            return dt;
        }

        internal DataTable WmsGetAssortInOut(string ym)
        {
            DateTime ymdStart = DateTime.ParseExact(ym + "01", "yyyyMMdd", CultureInfo.InvariantCulture);
            int year = ymdStart.Year;
            int month = ymdStart.Month;
            int dayInMonth = DateTime.DaysInMonth(year, month);
            OracleCommand strInWH = new OracleCommand();
            strInWH.CommandText = $@"SELECT TO_CHAR(W.ASTDATE,'YYYY-MM-DD') AS ASTDATE, W.ASTTYPE, W.MODEL,  W.PLTYPE, SUM(W.ASTQTY) ASTQTY  FROM SE.WMS_ASSORT W 
                                    WHERE comid = 'DCI'  AND MODEL LIKE '%' AND PLNO LIKE '%' AND TO_CHAR(astdate,'YYYY-MM-DD') BETWEEN '" + ymdStart.ToString("yyyy-MM-dd") + "' AND '" + ($"{year}-{month}-{dayInMonth}") + "' GROUP BY W.ASTDATE, W.ASTTYPE, W.MODEL,  W.PLTYPE";
            DataTable dt = _ALPHAPD.Query(strInWH);
            return dt;
        }
    }
}
