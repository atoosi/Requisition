using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SAP_Q100.Class
{
    class SqlData
    {
        private SqlConnection sqlConnection1;
        public SqlData()
        {
           sqlConnection1 = new SqlConnection("Data Source=SAP;Initial Catalog=Quest_Train;Persist Security Info=True;User ID=sa1;Password=s1ungod");

        }
    public DataTable Form1MainQuery(int jobNo)
        {
            try
            {
                #region sqlcommand
                SqlCommand cmd = new SqlCommand();
               // cmd.CommandText = "SELECT dbo.OITM.ItemCode,dbo.ordr.NumAtCard, dbo.OSCL.itemCode AS Expr1, dbo.OMRC.FirmName, dbo.OSCL.callType, dbo.OSCL.problemTyp, dbo.ORDR.DocNum, dbo.ORDR.Handwrtten, dbo.OSCL.callID, dbo.OSCT.callTypeID, dbo.ORDR.DocStatus,dbo.ORDR.DocTotal, dbo.OSCL.customer, dbo.OSCL.status, dbo.OSCS.Name, dbo.OSCL.custmrName, dbo.OSCL.U_BPCatNo, dbo.OSCL.internalSN, dbo.OITM.FrgnName, dbo.OSCL.U_WOSalPer, dbo.OSCL.U_WORcvdBy,dbo.OSCL.U_WORcvRmk, dbo.OSCL.U_WORecvDt, dbo.OSCL.U_WOPONo, dbo.OSCL.U_WOInbCar, dbo.OSCL.U_WOArrDt, dbo.OSCL.U_WODOM, dbo.OSCL.closeDate, dbo.OSCL.U_WOSONo, dbo.OSCL.U_WONaked,dbo.OSCL.U_WOCabnt, dbo.OSCL.U_WOSwivel, dbo.OSCL.U_WOCord, dbo.OSCL.U_WOVidCrd, dbo.OSCL.U_WOBook, dbo.OSCL.U_WORckCld, dbo.OSCL.U_WORack, dbo.OSCL.U_WOPCBMod, dbo.OSCL.subject,dbo.OSCL.U_WOPatt, dbo.OSCL.U_WOBroken, dbo.OSCL.U_WOKeys, dbo.OSCL.U_WOKeybrd, dbo.OSCL.U_WOMouse, dbo.OSCL.U_WOCmpSys, dbo.OITM.ItemName, dbo.OSCP.Name AS Expr2, dbo.ORDR.DocEntry FROM dbo.OSCP INNER JOIN dbo.OITM INNER JOIN dbo.OMRC ON dbo.OITM.FirmCode = dbo.OMRC.FirmCode INNER JOIN dbo.OSCL ON dbo.OITM.ItemCode = dbo.OSCL.itemCode INNER JOIN dbo.ORDR ON dbo.OSCL.callID = dbo.ORDR.DocNum ON dbo.OSCP.prblmTypID = dbo.OSCL.problemTyp INNER JOIN dbo.OSCS ON dbo.OSCL.status = dbo.OSCS.statusID LEFT OUTER JOIN dbo.OSCT ON dbo.OSCL.callType = dbo.OSCT.callTypeID WHERE(dbo.OSCL.callID = @jobNo)";
                cmd.CommandText = "SELECT dbo.OITM.ItemCode, dbo.OSCL.itemCode AS Expr1, dbo.OMRC.FirmName, dbo.OSCL.callType, dbo.OSCL.problemTyp, dbo.OSCL.callID, dbo.OSCT.callTypeID, dbo.OSCL.customer, dbo.OSCL.status, dbo.OSCS.Name,dbo.OSCL.custmrName, dbo.OSCL.U_BPCatNo, dbo.OSCL.internalSN, dbo.OITM.FrgnName, dbo.OSCL.U_WOSalPer, dbo.OSCL.U_WORcvdBy, dbo.OSCL.U_WORcvRmk, dbo.OSCL.U_WORecvDt, dbo.OSCL.U_WOPONo,dbo.OSCL.U_WOInbCar, dbo.OSCL.U_WOArrDt, dbo.OSCL.U_WODOM, dbo.OSCL.closeDate, dbo.OSCL.U_WOSONo, dbo.OSCL.U_WONaked, dbo.OSCL.U_WOCabnt, dbo.OSCL.U_WOSwivel, dbo.OSCL.U_WOCord,dbo.OSCL.U_WOVidCrd, dbo.OSCL.U_WOBook, dbo.OSCL.U_WORckCld, dbo.OSCL.U_WORack, dbo.OSCL.U_WOPCBMod, dbo.OSCL.subject, dbo.OSCL.U_WOPatt, dbo.OSCL.U_WOBroken, dbo.OSCL.U_WOKeys,dbo.OSCL.U_WOKeybrd, dbo.OSCL.U_WOMouse, dbo.OSCL.U_WOCmpSys, dbo.OITM.ItemName, dbo.OSCP.Name AS Expr2 FROM dbo.OSCP INNER JOIN dbo.OITM INNER JOIN dbo.OMRC ON dbo.OITM.FirmCode = dbo.OMRC.FirmCode INNER JOIN dbo.OSCL ON dbo.OITM.ItemCode = dbo.OSCL.itemCode ON dbo.OSCP.prblmTypID = dbo.OSCL.problemTyp INNER JOIN dbo.OSCS ON dbo.OSCL.status = dbo.OSCS.statusID LEFT OUTER JOIN dbo.OSCT ON dbo.OSCL.callType = dbo.OSCT.callTypeID WHERE(dbo.OSCL.callID = @jobNo)";
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@jobNo", SqlDbType.Int);
                cmd.Parameters["@jobNo"].Value = jobNo;
                cmd.CommandType = CommandType.Text;
                cmd.Connection = sqlConnection1;
                var query1 = new DataTable();

                //ds3 = new DataSet();
                SqlDataAdapter sda1 = new SqlDataAdapter();
                sqlConnection1.Open();
                sda1.SelectCommand = cmd;
                sda1.Fill(query1);

                sqlConnection1.Close();
                if (query1.Rows.Count > 0)
                    return query1;
                else return null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }

                #endregion
            }


        public string BpCodeList(string itemCode,string vendorCode)
        {
            try
            {
                #region sqlcommand
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "SELECT * FROM dbo.OSCN WHERE ItemCode = @itemcode and showSCN='Y' and CardCode=@vendorcode";
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@itemcode", SqlDbType.Int);
                cmd.Parameters["@itemcode"].Value = itemCode;
                cmd.Parameters.AddWithValue("@vendorcode", SqlDbType.Int);
                cmd.Parameters["@vendorcode"].Value = vendorCode;
                cmd.CommandType = CommandType.Text;
                cmd.Connection = sqlConnection1;
                var query1 = new DataTable();

                SqlDataAdapter sda1 = new SqlDataAdapter();
                sqlConnection1.Open();
                sda1.SelectCommand = cmd;
                sda1.Fill(query1);

                sqlConnection1.Close();
                if (query1.Rows.Count > 0)
                    return query1.Rows[0]["Substitute"].ToString();
                else return null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }

            #endregion
        }

        public string LastPo()
        {
            try
            {
                #region sqlcommand
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "  select top 1 docEntry from opor order by docnum desc";
                cmd.Parameters.Clear();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = sqlConnection1;
                var query1 = new DataTable();

                SqlDataAdapter sda1 = new SqlDataAdapter();
                sqlConnection1.Open();
                sda1.SelectCommand = cmd;
                sda1.Fill(query1);

                sqlConnection1.Close();
                if (query1.Rows.Count > 0)
                    return query1.Rows[0]["docEntry"].ToString();
                else return "";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }

            #endregion
                      
        }

        public DataTable BpCodeList2( string vendorCode)
        {
            try
            {
                #region sqlcommand
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "SELECT * FROM dbo.OSCN WHERE showSCN='Y' and CardCode=@vendorcode";
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@vendorcode", SqlDbType.Int);
                cmd.Parameters["@vendorcode"].Value = vendorCode;
                cmd.CommandType = CommandType.Text;
                cmd.Connection = sqlConnection1;
                var query1 = new DataTable();

                SqlDataAdapter sda1 = new SqlDataAdapter();
                sqlConnection1.Open();
                sda1.SelectCommand = cmd;
                sda1.Fill(query1);

                sqlConnection1.Close();
                if (query1.Rows.Count > 0)
                    return query1;
                else return null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }

            #endregion
        }

        public DataTable Form1WorkOrder1Query(int jobNo)
        {
            try
            {
                #region sqlcommand
                SqlCommand cmd = new SqlCommand();
                // cmd.CommandText = "SELECT dbo.OITM.ItemCode,dbo.ordr.NumAtCard, dbo.OSCL.itemCode AS Expr1, dbo.OMRC.FirmName, dbo.OSCL.callType, dbo.OSCL.problemTyp, dbo.ORDR.DocNum, dbo.ORDR.Handwrtten, dbo.OSCL.callID, dbo.OSCT.callTypeID, dbo.ORDR.DocStatus,dbo.ORDR.DocTotal, dbo.OSCL.customer, dbo.OSCL.status, dbo.OSCS.Name, dbo.OSCL.custmrName, dbo.OSCL.U_BPCatNo, dbo.OSCL.internalSN, dbo.OITM.FrgnName, dbo.OSCL.U_WOSalPer, dbo.OSCL.U_WORcvdBy,dbo.OSCL.U_WORcvRmk, dbo.OSCL.U_WORecvDt, dbo.OSCL.U_WOPONo, dbo.OSCL.U_WOInbCar, dbo.OSCL.U_WOArrDt, dbo.OSCL.U_WODOM, dbo.OSCL.closeDate, dbo.OSCL.U_WOSONo, dbo.OSCL.U_WONaked,dbo.OSCL.U_WOCabnt, dbo.OSCL.U_WOSwivel, dbo.OSCL.U_WOCord, dbo.OSCL.U_WOVidCrd, dbo.OSCL.U_WOBook, dbo.OSCL.U_WORckCld, dbo.OSCL.U_WORack, dbo.OSCL.U_WOPCBMod, dbo.OSCL.subject,dbo.OSCL.U_WOPatt, dbo.OSCL.U_WOBroken, dbo.OSCL.U_WOKeys, dbo.OSCL.U_WOKeybrd, dbo.OSCL.U_WOMouse, dbo.OSCL.U_WOCmpSys, dbo.OITM.ItemName, dbo.OSCP.Name AS Expr2, dbo.ORDR.DocEntry FROM dbo.OSCP INNER JOIN dbo.OITM INNER JOIN dbo.OMRC ON dbo.OITM.FirmCode = dbo.OMRC.FirmCode INNER JOIN dbo.OSCL ON dbo.OITM.ItemCode = dbo.OSCL.itemCode INNER JOIN dbo.ORDR ON dbo.OSCL.callID = dbo.ORDR.DocNum ON dbo.OSCP.prblmTypID = dbo.OSCL.problemTyp INNER JOIN dbo.OSCS ON dbo.OSCL.status = dbo.OSCS.statusID LEFT OUTER JOIN dbo.OSCT ON dbo.OSCL.callType = dbo.OSCT.callTypeID WHERE(dbo.OSCL.callID = @jobNo)";
                cmd.CommandText = "SELECT dbo.ORDR.NumAtCard, dbo.ORDR.DocNum, dbo.ORDR.Handwrtten, dbo.ORDR.DocStatus, dbo.ORDR.DocTotal, dbo.ORDR.DocEntry FROM dbo.OSCL INNER JOIN dbo.ORDR ON dbo.OSCL.callID = dbo.ORDR.DocNum WHERE(dbo.OSCL.callID = @jobNo)";
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@jobNo", SqlDbType.Int);
                cmd.Parameters["@jobNo"].Value = jobNo;
                cmd.CommandType = CommandType.Text;
                cmd.Connection = sqlConnection1;
                var query1 = new DataTable();

                //ds3 = new DataSet();
                SqlDataAdapter sda1 = new SqlDataAdapter();
                sqlConnection1.Open();
                sda1.SelectCommand = cmd;
                sda1.Fill(query1);

                sqlConnection1.Close();
                if (query1.Rows.Count > 0)
                    return query1;
                else return null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }

            #endregion
        }


        public string UserName(int userId)
        {
            try
            {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "SELECT  U_NAME from OUSR WHERE USERID = @userID";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@userID", SqlDbType.Int);
            cmd.Parameters["@userID"].Value = userId;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query2 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query2);
            sqlConnection1.Close();
            if (query2.Rows.Count > 0)

            return query2.Rows[0]["U_NAME"].ToString();
            else return null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }

        }


        public int UserId(string userName)
        {
            try
            {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "SELECT  USERID from OUSR WHERE U_NAME = @userName";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@userName", SqlDbType.Int);
            cmd.Parameters["@userName"].Value = userName;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query2 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query2);
            sqlConnection1.Close();
            if (query2.Rows.Count > 0)

                return System.Convert.ToInt32(query2.Rows[0]["USERID"]);
            else return 0;
        }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return 0;
            }
}


        public int InventoryDraftKey(int jobNo)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select DocEntry from ODRF where U_ITWONo = @jobNo and U_PartReqStatus = '1' and ObjType = '67'";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@jobNo", SqlDbType.Int);
            cmd.Parameters["@jobNo"].Value = jobNo;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query3 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query3);
            sqlConnection1.Close();
            if (query3.Rows.Count > 0)
                return System.Convert.ToInt32(query3.Rows[0]["DocEntry"]);
            else return 0;
        }

        public List<string> Form1Grid1Query1()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select Name from dbo.[@V33SYMPTOMCODES]";
            cmd.Parameters.Clear();
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query4 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query4);
            sqlConnection1.Close();
            List<string> items=new List<string>();
            foreach(DataRow item in query4.Rows)
            {
                items.Add(item["Name"].ToString());
            }
            return items.ToList();
        }




   
        public string GoodReceive(int reqNumber)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "SELECT dbo.POR1.LineStatus FROM dbo.ODRF AS a INNER JOIN dbo.DRF1 AS b ON a.DocEntry = b.DocEntry INNER JOIN dbo.POR1 ON b.DocEntry = dbo.POR1.U_RequisitionNo AND b.LineNum = dbo.POR1.U_RequisitionLineNo WHERE(a.ObjType = '67') AND(a.U_ITWONo IS NOT NULL) AND(a.DocEntry = @reqNo) GROUP BY dbo.POR1.LineStatus HAVING(dbo.POR1.LineStatus = 'C')";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@reqNo", SqlDbType.Int);
            cmd.Parameters["@reqNo"].Value = reqNumber;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query4 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query4);
            sqlConnection1.Close();
            if (query4.Rows.Count > 0)
                return query4.Rows[0]["LineStatus"].ToString();
            return "";
            
        }



        public List<string> Form1Grid1Query2()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select Name from dbo.[@V33TESTCODES]";
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query5 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query5);
            sqlConnection1.Close();
            List<string> items = new List<string>();
            foreach (DataRow item in query5.Rows)
            {
                items.Add(item["Name"].ToString());
            }
            return items;
        }

        public List<string> Form1Grid2Query1()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select Name from dbo.[@V33REPAIRCODES]";
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            List<string> items = new List<string>();
            foreach (DataRow item in query6.Rows)
            {
                items.Add(item["Name"].ToString());
            }
            return items.ToList();
        }

        public List<string> Form1Grid2Query2()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select Name from dbo.[@V33CAUSECODES]";
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            List<string> items = new List<string>();
            foreach (DataRow item in query6.Rows)
            {
                items.Add(item["Name"].ToString());
            }
            return items.ToList();
        }

        public string[] UserCodeList()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select ItemCode, ItemName, FrgnName, ItemType, AvgPrice, ItmsGrpCod from OITM where (ItemType = 'L') AND (ItmsGrpCod = 100)";

            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            List<string> items = new List<string>();
            items.Add(" ");
            foreach (DataRow item in query6.Rows)
            {
                items.Add(item["ItemCode"].ToString());
            }
            return items.ToArray();
        }


        public string[] ItemCodesList()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select ItemCode from OITM where InvntItem ='Y' and SellItem ='Y' and frozenFor='N'";

            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            List<string> items = new List<string>();
            foreach (DataRow item in query6.Rows)
            {
                items.Add(item["ItemCode"].ToString());
            }
            return items.ToArray();
        }

        public string[] CardCodeList()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select CardCode,CardName from OCRD where CardType = 'S'";

            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            List<string> items = new List<string>();
            foreach (DataRow item in query6.Rows)
            {
                items.Add(item["CardName"].ToString());
            }
            return items.ToArray();
        }


        public string CardCodeByName(string name)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select CardCode,CardName from OCRD where CardType = 'S' and CardName=@cardName";

            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@cardName", SqlDbType.Text);
            cmd.Parameters["@cardName"].Value = name;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            List<string> items = new List<string>();
            foreach (DataRow item in query6.Rows)
            {
                items.Add(item["CardCode"].ToString());
            }
            return items[0].ToString() ;
        }

        public DataTable ItemsByItemCode(string itemcode)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select ItemCode,OnHand,FrgnName from OITM where ItemCode=@itemcode";

            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@itemcode", SqlDbType.Text);
            cmd.Parameters["@itemcode"].Value = itemcode;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            return query6;
        }

        public DataTable Form1Grid5Query1(int jobNo)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select c.VisOrder, a.DocNum,a.CreateDate from OPOR as a join ODRF as b on a.U_RequisitionNumber = b.DocEntry join DRF1 as c on b.DocEntry = c.DocEntry join POR1 as d on c.VisOrder = d.U_RequisitionLineNo where a.DocEntry = d.DocEntry and b.ObjType = '67' and b.U_ITWONo = @jobNo";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@jobNo", SqlDbType.Int);
            cmd.Parameters["@jobNo"].Value = jobNo;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            return query6 ;
        }

        public DataTable Form2Grid2Query1(int draftKey)
        {
            SqlCommand cmd = new SqlCommand();
            //cmd.CommandText = "select  c.VisOrder, a.DocNum,a.CreateDate from OPOR as a join ODRF as b on a.U_RequisitionNumber = b.DocEntry join DRF1 as c on b.DocEntry = c.DocEntry join POR1 as d on c.VisOrder = d.U_RequisitionLineNo where a.DocEntry = d.DocEntry and b.ObjType = '67' and a.DocEntry = d.DocEntry and  b.DocEntry= @draftKey";
            cmd.CommandText = "SELECT  a.DocDueDate,c.VisOrder, a.DocNum, a.CreateDate FROM dbo.DRF1 AS c INNER JOIN dbo.ODRF AS b ON c.DocEntry = b.DocEntry INNER JOIN dbo.POR1 AS d ON c.VisOrder = d.U_RequisitionLineNo AND b.DocEntry = d.U_RequisitionNo INNER JOIN dbo.OPOR AS a ON d.DocEntry = a.DocEntry AND d.DocEntry = a.DocEntry WHERE(b.ObjType = '67')  and  b.DocEntry= @draftKey";
            cmd.Parameters.Clear();
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            cmd.Parameters.AddWithValue("@draftKey", SqlDbType.Int);
            cmd.Parameters["@draftKey"].Value = draftKey;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            return query6;
        }

        public List<Requisition> Requisitions(int jobNo)
        {

            SqlCommand cmd = new SqlCommand();
            //cmd.CommandText = "SELECT a.DocEntry, a.U_ITWONo, a.CreateDate, SUM(b.Quantity) AS Qu, a.U_PartReqStatus, a.U_LocationStatus, dbo.OUSR.U_NAME AS Owner, c.Descr  FROM dbo.ODRF AS a INNER JOIN dbo.DRF1 AS b ON a.DocEntry = b.DocEntry LEFT OUTER JOIN dbo.OUSR ON a.OwnerCode = dbo.OUSR.USERID LEFT OUTER JOIN (SELECT FldValue, Descr FROM dbo.UFD1 WHERE TableID = N'ODRF' AND FieldID = 222) AS c ON a.U_PartReqStatus = c.FldValue WHERE a.ObjType = '67' and a.U_ITWONo = @jobNo GROUP BY a.DocEntry, a.U_ITWONo, a.U_PartReqStatus, a.CreateDate, a.U_LocationStatus, dbo.OUSR.U_NAME, c.Descr HAVING a.U_ITWONo IS NOT NULL ORDER BY a.U_ITWONo DESC";
            cmd.CommandText = "select a.DocEntry,a.U_ITWONo, a.CreateDate, sum(b.Quantity) as Qu, a.U_PartReqStatus from  ODRF as a  join DRF1 as b on a.DocEntry= b.DocEntry where  a.U_ITWONo = @jobNo and a.ObjType = '67' group by a.DocEntry,a.U_ITWONo, a.U_PartReqStatus, a.CreateDate ";
            cmd.Parameters.Clear();
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            cmd.Parameters.AddWithValue("@jobNo", SqlDbType.Int);
            cmd.Parameters["@jobNo"].Value = jobNo;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query7 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query7);
            sqlConnection1.Close();

            List<Requisition> items = new List<Requisition>();
            foreach (DataRow item in query7.Rows)
            {
                items.Add(new Requisition {CreationDate=System.Convert.ToDateTime(item["CreateDate"]),ItemsQuantity=System.Convert.ToDecimal(item["Qu"]),ItemsType=0,SeqNumber=System.Convert.ToInt32(item["DocEntry"]),WorkOrderNumber=System.Convert.ToInt32(item["U_ITWONo"]),Status=item["U_PartReqStatus"].ToString(), ReqisitionStatus = 0});
            }
            return items.ToList();
        }

        public DataTable Form2Query1(int userId,string pass)
        {
     
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select a.empID, a.firstName, a.passportNo, b.teamID from ohem as a join HTM1 as b on a.empID = b.empID where  a.passportNo = @pass and a.userId = @userId ";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@userId", SqlDbType.Int);
            cmd.Parameters["@userId"].Value = userId;
            cmd.Parameters.AddWithValue("@pass", SqlDbType.Text);
            cmd.Parameters["@pass"].Value = pass;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query8 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query8);
            sqlConnection1.Close();
            return query8;
        }


        public string UserCheck(string userName)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select ItemCode from OITM where ItemCode = @userName";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@userName", SqlDbType.Int);
            cmd.CommandType = CommandType.Text;
            cmd.Parameters["@userName"].Value = userName;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();
            if (query6.Rows.Count > 0)
                return query6.Rows[0]["ItemCode"].ToString();
            else return null;
        }


        public int CheckPoNumber(string keyPo)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select DocEntry from OPOR where DocNum = @docNum and DocStatus='O'";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@docNum", SqlDbType.Int);
            cmd.CommandType = CommandType.Text;
            cmd.Parameters["@docNum"].Value = keyPo;
            cmd.Connection = sqlConnection1;
            SqlDataAdapter sda1 = new SqlDataAdapter();
            var query6 = new DataTable();
            sda1.SelectCommand = cmd;
            sqlConnection1.Open();
            sda1.Fill(query6);
            sqlConnection1.Close();

            if (query6.Rows.Count > 0)
                return System.Convert.ToInt32(query6.Rows[0]["DocEntry"]);
            else return 0;
        }


        //public bool ItemReceive(int reqNo)
        //{
        //    SqlCommand cmd = new SqlCommand();
        //    cmd.CommandText = "SELECT ODRF.DocEntry, SUM(POR1.OpenQty) AS IsOpen FROM dbo.ODRF INNER JOIN DRF1 ON ODRF.DocEntry = dbo.DRF1.DocEntry INNER JOIN POR1 ON DRF1.DocEntry = POR1.U_RequisitionNo AND DRF1.VisOrder = POR1.U_RequisitionLineNo INNER JOIN PDN1 ON POR1.DocEntry = PDN1.BaseEntry AND POR1.LineNum = PDN1.BaseLine AND POR1.ObjType = PDN1.BaseType WHERE POR1.TargetType IS NOT NULL GROUP BY ODRF.DocEntry HAVING ODRF.DocEntry = @docEntry AND SUM(POR1.OpenQty) = 0";
        //    cmd.Parameters.Clear();
        //    cmd.Parameters.AddWithValue("@docEntry", SqlDbType.Int);
        //    cmd.CommandType = CommandType.Text;
        //    cmd.Parameters["@docEntry"].Value = reqNo;
        //    cmd.Connection = sqlConnection1;
        //    SqlDataAdapter sda1 = new SqlDataAdapter();
        //    var query6 = new DataTable();
        //    sda1.SelectCommand = cmd;
        //    sqlConnection1.Open();
        //    sda1.Fill(query6);
        //    //sqlConnection1.Close();


        //    cmd.CommandText = "SELECT dbo.ODRF.DocEntry, SUM(dbo.POR1.OpenQty) AS IsOpen FROM dbo.ODRF INNER JOIN dbo.DRF1 ON dbo.ODRF.DocEntry = dbo.DRF1.DocEntry INNER JOIN dbo.POR1 ON dbo.DRF1.DocEntry = dbo.POR1.U_RequisitionNo AND dbo.DRF1.VisOrder = dbo.POR1.U_RequisitionLineNo WHERE dbo.POR1.TargetType IS NOT NULL GROUP BY dbo.ODRF.DocEntry HAVING dbo.ODRF.DocEntry = @docEntry AND SUM(dbo.POR1.OpenQty) = 0";
        //    cmd.Parameters.Clear();
        //    cmd.Parameters.AddWithValue("@docEntry", SqlDbType.Int);
        //    cmd.CommandType = CommandType.Text;
        //    cmd.Parameters["@docEntry"].Value = reqNo;
        //    cmd.Connection = sqlConnection1;
        //    sda1 = new SqlDataAdapter();
        //    var query7 = new DataTable();
        //    sda1.SelectCommand = cmd;
        //    sda1.Fill(query7);
        //    sqlConnection1.Close();

        //    if (query6.Rows.Count == query7.Rows.Count)
        //        return true;
        //    else return false;

        //}
    }
}
