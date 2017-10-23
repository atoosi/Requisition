using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace SAP_Q100.Class
{
    public class Form1GridViewsInitializer
    {

        private static SAPbobsCOM.Company pCompany;
        private SAPbobsCOM.Documents oOrder;
        public SAPbobsCOM.Documents oPoDraft;
        private int jobNo;
        private int orderKey;
        private SqlData data = new SqlData();
        private List<string> counter;
        public Dictionary<int, int> LineSelector;
        public Label label;

        public Form1GridViewsInitializer(int jobNumber, int orderObjectKey, SAPbobsCOM.Company company)
        {
            jobNo = jobNumber;
            pCompany = company;

            orderKey = orderObjectKey;

            oOrder = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            oOrder.GetByKey(orderKey);

        }

        public void DataGridView1Initializer(DataGridView datagrid1)
        {

            if (oOrder != null && oOrder.HandWritten.Equals(SAPbobsCOM.BoYesNoEnum.tYES))//&& oOrder.DocumentStatus == SAPbobsCOM.BoStatus.bost_Open)
            {
                var dt1 = new DataTable();
                dt1.Columns.Add("SYMP", typeof(String));
                dt1.Columns.Add("TEST", typeof(String));
                dt1.Columns.Add("TEST TIME", typeof(String));
                dt1.Columns.Add("TECH", typeof(String));
                dt1.Columns.Add("DATE", typeof(String));
                dt1.Columns.Add("VENDOR NO", typeof(string));
                dt1.Columns.Add("LINE NUMBER", typeof(string));

                for (var i = 0; i < oOrder.Lines.Count; i++)
                {
                    oOrder.Lines.SetCurrentLine(i);
                    if (!string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORSymCd").Value.ToString()) || !string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORTestC").Value.ToString()))
                    {
                        dt1.Rows.Add(new object[]
                    {

                    oOrder.Lines.UserFields.Fields.Item("U_WORSymCd").Value, oOrder.Lines.UserFields.Fields.Item("U_WORTestC").Value,
                                oOrder.Lines.Quantity,oOrder.Lines.ItemCode,oOrder.Lines.UserFields.Fields.Item("U_WORDate").Value.ToString("MM/dd/yyyy"),
                               oOrder.Lines.VisualOrder,oOrder.Lines.LineNum
                    });
                    }
                }


                //  GridMaker(datagrid1, dt1, data.Form1Grid1Query1(), data.Form1Grid1Query2());





                var ds = new BindingSource();
                ds.DataSource = dt1;

                DataGridViewTextBoxCell noCell = new DataGridViewTextBoxCell();

                DataGridViewComboBoxCell firstCell = new DataGridViewComboBoxCell();
                firstCell.Items.AddRange(data.Form1Grid1Query1().ToArray());


                DataGridViewComboBoxCell secondCell = new DataGridViewComboBoxCell();
                secondCell.Items.AddRange(data.Form1Grid1Query2().ToArray());

                DataGridViewColumn noColumn = new DataGridViewColumn(noCell);
                DataGridViewColumn firstColumn = new DataGridViewColumn(firstCell);
                DataGridViewColumn secondColumn = new DataGridViewColumn(secondCell);


                datagrid1.Columns.Add(noColumn);
                datagrid1.Columns.Add(firstColumn);
                datagrid1.Columns.Add(secondColumn);
                datagrid1.DataSource = ds;
                datagrid1.Columns[0].HeaderText = "No.";
                datagrid1.Columns[1].HeaderText = datagrid1.Columns[3].HeaderText;
                datagrid1.Columns[2].HeaderText = datagrid1.Columns[4].HeaderText;
                counter = new List<string>();
                LineSelector = new Dictionary<int, int>();

                foreach (DataGridViewRow item in datagrid1.Rows)
                {
                    if (item.Cells[3].Value != null)
                    {
                        item.Cells[0].Value = item.Index + 1;
                        counter.Add(System.Convert.ToString(item.Index + 1));
                        LineSelector.Add(item.Index + 1, System.Convert.ToInt32(item.Cells[9].Value));
                    }

                    foreach (string itm1 in data.Form1Grid1Query1().ToArray())
                    {
                        if (!(item.Cells[3].Value == null))
                        {
                            if (itm1.Length <= 3)
                            {
                                if (itm1 == item.Cells[3].Value.ToString())
                                    item.Cells[1].Value = item.Cells[3].Value;
                            }
                            else if (itm1.Substring(0, 3).Equals(item.Cells[3].Value.ToString()))
                            {
                                item.Cells[1].Value = itm1;
                                break;
                            }
                        }
                    }

                    foreach (string itm2 in data.Form1Grid1Query2().ToArray())
                    {
                        if (!(item.Cells[4].Value == null))
                        {
                            if (itm2.Length <= 3)
                            {
                                if (itm2 == item.Cells[4].Value.ToString())
                                    item.Cells[2].Value = item.Cells[4].Value;
                            }
                            else if (itm2.Substring(0, 3).Equals(item.Cells[4].Value.ToString()))
                            {
                                item.Cells[2].Value = itm2;
                                break;
                            }
                        }
                    }
                }
                //datagrid.Columns[0].ReadOnly = true;
                //datagrid.Columns[6].ReadOnly = true;
                //datagrid.Columns[7].ReadOnly = true;
                //datagrid.Columns[8].ReadOnly = true;

                DisableColumn(datagrid1, new int[] { 0, 6, 7, 8 });


                datagrid1.Columns[3].Visible = false;
                datagrid1.Columns[4].Visible = false;
                datagrid1.Columns[8].Visible = false;
                 datagrid1.Columns[9].Visible = false;


                datagrid1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                datagrid1.AutoResizeColumns();


            }

        }

        public void DataGridView2Initializer(DataGridView datagrid2)
        {


            if (oOrder != null && oOrder.HandWritten.Equals(SAPbobsCOM.BoYesNoEnum.tYES)) //&& oOrder.DocumentStatus == SAPbobsCOM.BoStatus.bost_Open)
            {
                var dt2 = new DataTable();

                dt2.Columns.Add("REPAIR", typeof(String));
                dt2.Columns.Add("CAUSE", typeof(String));
                dt2.Columns.Add("REPAIR TIME", typeof(String));
                dt2.Columns.Add("TECH", typeof(String));
                dt2.Columns.Add("DATE", typeof(String));
                dt2.Columns.Add("VENDOR NO", typeof(string));
                dt2.Columns.Add("RELATED LINE", typeof(string));

                for (var i = 0; i < oOrder.Lines.Count; i++)
                {
                    oOrder.Lines.SetCurrentLine(i);
                    if (!string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORRepCd").Value.ToString()) || !string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORCause").Value.ToString()))
                    {
                        dt2.Rows.Add(new object[]
                        {
                    oOrder.Lines.UserFields.Fields.Item("U_WORRepCd").Value, oOrder.Lines.UserFields.Fields.Item("U_WORCause").Value,oOrder.Lines.Quantity,oOrder.Lines.ItemCode,oOrder.Lines.UserFields.Fields.Item("U_WORDate").Value.ToString("MM/dd/yyyy"),oOrder.Lines.VisualOrder
                    ,oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value
                        });
                    }

                }
                //   GridMaker(datagrid2, dt2, data.Form1Grid2Query1(),data.Form1Grid2Query2());




                var ds = new BindingSource();
                ds.DataSource = dt2;

                //DataGridViewTextBoxCell noCell = new DataGridViewTextBoxCell();
                DataGridViewComboBoxCell noCell = new DataGridViewComboBoxCell();
                noCell.Items.AddRange(counter.ToArray());

                DataGridViewComboBoxCell firstCell = new DataGridViewComboBoxCell();
                firstCell.Items.AddRange(data.Form1Grid2Query1().ToArray());


                DataGridViewComboBoxCell secondCell = new DataGridViewComboBoxCell();
                secondCell.Items.AddRange(data.Form1Grid2Query2().ToArray());

                DataGridViewColumn noColumn = new DataGridViewColumn(noCell);
                DataGridViewColumn firstColumn = new DataGridViewColumn(firstCell);
                DataGridViewColumn secondColumn = new DataGridViewColumn(secondCell);


                datagrid2.Columns.Add(noColumn);
                datagrid2.Columns.Add(firstColumn);
                datagrid2.Columns.Add(secondColumn);
                datagrid2.DataSource = ds;
                datagrid2.Columns[0].HeaderText = "No.";
                datagrid2.Columns[1].HeaderText = datagrid2.Columns[3].HeaderText;
                datagrid2.Columns[2].HeaderText = datagrid2.Columns[4].HeaderText;



                foreach (DataGridViewRow item in datagrid2.Rows)
                {
                    if (item.Cells[9].Value != null)
                    {
                        if (!string.IsNullOrEmpty(item.Cells[9].Value.ToString()))
                        {
                            foreach (var it in LineSelector)
                            {
                                if (System.Convert.ToInt32(item.Cells[9].Value) == it.Value)
                                    item.Cells[0].Value = it.Key.ToString();
                            }
                        }
                    }


                    foreach (string itm1 in data.Form1Grid2Query1().ToArray())
                    {
                        if (!(item.Cells[3].Value == null))
                        {
                            if (itm1.Length <= 3)
                            {
                                if (itm1 == item.Cells[3].Value.ToString())
                                    item.Cells[1].Value = item.Cells[3].Value;
                            }
                            else if (itm1.Substring(0, 3).Equals(item.Cells[3].Value.ToString()))
                            {
                                item.Cells[1].Value = itm1;
                                break;
                            }
                        }
                    }

                    foreach (string itm2 in data.Form1Grid2Query2().ToArray())
                    {
                        if (!(item.Cells[4].Value == null))
                        {
                            if (itm2.Length <= 3)
                            {
                                if (itm2 == item.Cells[4].Value.ToString())
                                    item.Cells[2].Value = item.Cells[4].Value;
                            }
                            else if (itm2.Substring(0, 3).Equals(item.Cells[4].Value.ToString()))
                            {
                                item.Cells[2].Value = itm2;
                                break;
                            }
                        }
                    }
                }
                //datagrid.Columns[0].ReadOnly = true;
                //datagrid.Columns[6].ReadOnly = true;
                //datagrid.Columns[7].ReadOnly = true;
                //datagrid.Columns[8].ReadOnly = true;

                DisableColumn(datagrid2, new int[] { 6, 7, 8 });


                datagrid2.Columns[3].Visible = false;
                datagrid2.Columns[4].Visible = false;
                datagrid2.Columns[8].Visible = false;
                datagrid2.Columns[9].Visible = false;

                datagrid2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                datagrid2.AutoResizeColumns();

            }
        }

        public void DataGridView3Initializer(DataGridView datagrid3)
        {
            if (oOrder != null && oOrder.HandWritten.Equals(SAPbobsCOM.BoYesNoEnum.tYES))// && oOrder.DocumentStatus==SAPbobsCOM.BoStatus.bost_Open)
            {
                var dt3 = new DataTable();

                dt3.Columns.Add("NAME OR P/N", typeof(String));
                dt3.Columns.Add("DESCRIPTION", typeof(String));
                dt3.Columns.Add("DIS", typeof(String));
                dt3.Columns.Add("LOC", typeof(String));
                dt3.Columns.Add("QTY", typeof(String));
                dt3.Columns.Add("FIX Y/N", typeof(String));
                dt3.Columns.Add("TECH", typeof(String));
                dt3.Columns.Add("DATE", typeof(String));
                dt3.Columns.Add("VENDOR NO", typeof(String));
                dt3.Columns.Add("REQU NO", typeof(String));
                dt3.Columns.Add("REQU QTY", typeof(String));
                dt3.Columns.Add("REQU DEL QTY", typeof(String));
                dt3.Columns.Add("RELATED LINE", typeof(string));

                for (var i = 0; i < oOrder.Lines.Count; i++)
                {
                    oOrder.Lines.SetCurrentLine(i);

                    if (string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORRepCd").Value.ToString()) && string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORCause").Value.ToString()) && string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORSymCd").Value.ToString()) && string.IsNullOrEmpty(oOrder.Lines.UserFields.Fields.Item("U_WORTestC").Value.ToString()))
                    {
                        dt3.Rows.Add(new object[]
                    {

                            oOrder.Lines.ItemCode=="UNKNOWN"? oOrder.Lines.UserFields.Fields.Item("U_POSrvcItm").Value:oOrder.Lines.ItemCode,
                            oOrder.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value,
                            oOrder.Lines.UserFields.Fields.Item("U_WORDis").Value,
                            oOrder.Lines.UserFields.Fields.Item("U_WORLoc").Value,
                            oOrder.Lines.PackageQuantity,
                            oOrder.Lines.UserFields.Fields.Item("U_WORFix").Value,
                            oOrder.Lines.ItemCode,
                            oOrder.Lines.UserFields.Fields.Item("U_WORDate").Value.ToString("MM/dd/yyyy"),
                            oOrder.Lines.VisualOrder,
                            oOrder.Lines.UserFields.Fields.Item("U_WORReqNo").Value,
                            oOrder.Lines.UserFields.Fields.Item("U_WORReqQty").Value,
                            oOrder.Lines.UserFields.Fields.Item("U_WORDelQty").Value,
                            oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value

        });
                    }
                }




                var ds = new BindingSource();
                ds.DataSource = dt3;

                //  DataGridViewTextBoxCell noCell = new DataGridViewTextBoxCell();
                DataGridViewComboBoxCell noCell = new DataGridViewComboBoxCell();
                noCell.Items.AddRange(counter.ToArray());

                DataGridViewComboBoxCell secondCell = new DataGridViewComboBoxCell();
                secondCell.Items.AddRange(new string[] { "N", "R", "U", "-" });

                DataGridViewComboBoxCell thirdCell = new DataGridViewComboBoxCell();
                thirdCell.Items.AddRange(new string[] { "Yes", "No", "-" });

                //DataGridViewComboBoxCell forthCell = new DataGridViewComboBoxCell();
                //forthCell.Items.AddRange(q1);


                DataGridViewColumn noColumn = new DataGridViewColumn(noCell);

                DataGridViewColumn secondColumn = new DataGridViewColumn(secondCell);
                DataGridViewColumn thirdColumn = new DataGridViewColumn(thirdCell);

                datagrid3.Columns.Add(noColumn);
                //  datagrid.Columns.Add(firstColumn);

                datagrid3.DataSource = ds;
                datagrid3.Columns.Add(secondColumn);
                datagrid3.Columns.Add(thirdColumn);

                datagrid3.Columns[0].HeaderText = "No.";
                //datagrid.Columns[1].HeaderText = datagrid.Columns[3].HeaderText;
                datagrid3.Columns[14].HeaderText = datagrid3.Columns[3].HeaderText;
                datagrid3.Columns[15].HeaderText = datagrid3.Columns[6].HeaderText;

                foreach (DataGridViewRow item in datagrid3.Rows)
                {


                    //if (item.Cells[2].Value != null)
                    //    item.Cells[0].Value = item.Index + 1;

                    if (item.Cells[13].Value != null)
                    {
                        foreach (var itm in LineSelector)
                        {
                            if (!string.IsNullOrEmpty(item.Cells[13].Value.ToString()))
                            {
                                if (System.Convert.ToInt32(item.Cells[13].Value) == itm.Value)
                                    item.Cells[0].Value = itm.Key.ToString();
                            }
                        }
                    }
                    if (item.Cells[3].Value != null)
                        item.Cells[14].Value = item.Cells[3].Value;

                    if (item.Cells[6].Value != null)
                    {
                        if (item.Cells[6].Value.ToString().Equals("Y"))
                        {
                            item.Cells[15].Value = "Yes";
                        }
                        else if (item.Cells[6].Value.ToString().Equals("N"))
                        {
                            item.Cells[15].Value = "No";
                        }
                        else
                        {
                            item.Cells[15].Value = "-";
                        }
                    }

                    if (item.Cells[7].Value != null)
                    {
                        if (!item.Cells[7].Value.ToString().Equals("UNKNOWN"))
                        {
                            item.Cells[2].ReadOnly = true;
                        }
                    }

                }

                DisableColumn(datagrid3, new int[] { 3, 6, 7, 8, 9, 10 });

                datagrid3.Columns[3].Visible = false;
                datagrid3.Columns[6].Visible = false;
                datagrid3.Columns[7].Visible = false;
                datagrid3.Columns[9].Visible = false;
                datagrid3.Columns[10].Visible = false;
                datagrid3.Columns[11].Visible = false;
                datagrid3.Columns[12].Visible = false;
                datagrid3.Columns[13].Visible = false;

                //datagrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                //datagrid.AutoResizeColumns();

                datagrid3.Columns[0].Width = 30;
                datagrid3.Columns[1].Width = 160;




                // GridMaker(datagrid3, dt3, query1);
                //  GridMaker2(datagrid3, dt3, data.ItemCodesList());
                //  }
                label.Visible = false;
            }
        }








        public void StatusChangerForm2()
        {
            foreach (var item in data.Requisitions(jobNo))
            {
                if (item.Status == "6")
                {

                    if (data.GoodReceive(item.SeqNumber ?? default(int)).Equals("C"))
                    {
                        var oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                        var isAvailable = oInvTransDraft.GetByKey(item.SeqNumber ?? default(int));
                        if (isAvailable)
                        {
                            if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "6")
                            {
                                //    if (data.ItemReceive(item.SeqNumber ?? default(int)))
                                //{


                                oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "10";

                                oInvTransDraft.Update();
                                //
                                int IretCode = oInvTransDraft.Update();
                                if (IretCode != 0)
                                {
                                    // MessageBox.Show(LastErrorMessage(IretCode));
                                    oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                    oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                    isAvailable = oInvTransDraft.GetByKey(item.SeqNumber ?? default(int));
                                }

                                else
                                {

                                    LocationClass.Changer(pCompany, jobNo, 34);

                                    //form2.SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "34", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());

                                }
                            }
                        }
                    }
                }
            }

        }






        public void DataGridView4Initializer(DataGridView datagrid4)
        {
            // odrfObject = new ODRFConnector(jobNo, false);
            //if (odrfObject.lResult == 0)
            //{
            var dt4 = new DataTable();

            dt4.Columns.Add("NO.", typeof(int));
            dt4.Columns.Add("SEQ Number", typeof(int));
            //dt4.Columns.Add("Work Order Number", typeof(int));
            //dt4.Columns.Add("Items Type", typeof(int));
            //dt4.Columns.Add("Items Quantity", typeof(int));
            dt4.Columns.Add("Status", typeof(string));
            dt4.Columns.Add("Creation Date", typeof(DateTime));
            int counter = 0;
            string status = "";
            foreach (var item in data.Requisitions(jobNo))
            {
                switch (item.Status)
                {
                    case "1":
                        status = "Draft";
                        break;
                    case "2":
                        status = "Rejected/Void";
                        break;
                    case "3":
                        status = "Production Pending";
                        break;
                    case "4":
                        status = "Inventory Pending";
                        break;
                    case "5":
                        status = "Purchase Pending";
                        break;
                    case "6":
                        status = "Good Receipt";
                        break;
                    case "7":
                        status = "Esclated";
                        break;
                    case "8":
                        status = "Sending Items Approval";
                        break;
                    case "9":
                        status = "Completed";
                        break;

                    case "10":
                        status = "Issue Part";
                        break;
                }
                dt4.Rows.Add(new object[] { ++counter,item.SeqNumber, /*item.WorkOrderNumber, item.ItemsType, item.ItemsQuantity,*/
                 status  , item.CreationDate });
            }

            datagrid4.DataSource = dt4;

            datagrid4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            datagrid4.AutoResizeColumns();
            DisableColumn(datagrid4, new int[] { 0, 1, 2, 3 });

            // }
        }

        public void DataGridView5Initializer(DataGridView datagrid5)
        {

            var dt5 = new DataTable();



            dt5.Columns.Add("NO.", typeof(int));
            dt5.Columns.Add("SEQ Number", typeof(int));
            dt5.Columns.Add("Status", typeof(string));
            dt5.Columns.Add("Creation Date", typeof(DateTime));
            int counter = 0;

            foreach (DataRow item in data.Form1Grid5Query1(jobNo).Rows)
            {


                dt5.Rows.Add(new object[] { ++counter,item["DocNum"].ToString(), /*item.WorkOrderNumber, item.ItemsType, item.ItemsQuantity,*/
                    "",item["CreateDate"].ToString() });
            }

            datagrid5.DataSource = dt5;

            datagrid5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            datagrid5.AutoResizeColumns();

            DisableColumn(datagrid5, new int[] { 0, 1, 2, 3 });
            //}
        }




        //private void GridMaker(DataGridView datagrid, DataTable dt, string[] q1)
        //{
        //    var ds = new BindingSource();
        //    ds.DataSource = dt;

        //    DataGridViewTextBoxCell noCell = new DataGridViewTextBoxCell();

        //    DataGridViewComboBoxCell firstCell = new DataGridViewComboBoxCell();
        //    firstCell.Items.AddRange(q1);

        //    DataGridViewComboBoxCell secondCell = new DataGridViewComboBoxCell();
        //    secondCell.Items.AddRange(new string[] { "N", "R", "U", "-" });

        //    DataGridViewComboBoxCell thirdCell = new DataGridViewComboBoxCell();
        //    thirdCell.Items.AddRange(new string[] { "Yes", "No","-" });


        //    DataGridViewColumn noColumn = new DataGridViewColumn(noCell);
        //    DataGridViewColumn firstColumn = new DataGridViewColumn(firstCell);

        //    DataGridViewColumn secondColumn = new DataGridViewColumn(secondCell);
        //    DataGridViewColumn thirdColumn = new DataGridViewColumn(thirdCell);

        //    datagrid.Columns.Add(noColumn);
        //    datagrid.Columns.Add(firstColumn);
        //    datagrid.Columns.Add(secondColumn);
        //    datagrid.DataSource = ds;
        //    datagrid.Columns.Add(thirdColumn);

        //    datagrid.Columns[0].HeaderText = "No.";
        //    datagrid.Columns[1].HeaderText = datagrid.Columns[3].HeaderText;
        //    datagrid.Columns[2].HeaderText = datagrid.Columns[4].HeaderText;
        //    datagrid.Columns[14].HeaderText = datagrid.Columns[7].HeaderText;

        //    foreach (DataGridViewRow item in datagrid.Rows)
        //    {
        //        if (item.Cells[3].Value != null)
        //            item.Cells[0].Value = item.Index + 1;
        //        if (item.Cells[3].Value != null)
        //            item.Cells[1].Value = item.Cells[3].Value;
        //        if (item.Cells[4].Value != null)
        //            item.Cells[2].Value = item.Cells[4].Value;
        //        if (item.Cells[7].Value != null)
        //        {
        //            if (item.Cells[7].Value.ToString().Equals("Y"))
        //            {
        //                item.Cells[14].Value = "Yes";
        //            }
        //            else if (item.Cells[7].Value.ToString().Equals("N"))
        //            {
        //                item.Cells[14].Value = "No";
        //            }
        //            else
        //            {
        //                item.Cells[14].Value = "-";
        //            }
        //        }

        //    }

        //    DisableColumn(datagrid, new int[] { 0, 3, 4, 7, 8, 9, 10, 11 });

        //    //datagrid.Columns[0].ReadOnly = true;
        //    //datagrid.Columns[3].ReadOnly = true;
        //    //datagrid.Columns[4].ReadOnly = true;
        //    //datagrid.Columns[7].ReadOnly = true;
        //    //datagrid.Columns[8].ReadOnly = true;
        //    //datagrid.Columns[9].ReadOnly = true;
        //    //datagrid.Columns[10].ReadOnly = true;
        //    //datagrid.Columns[11].ReadOnly = true;

        //    datagrid.Columns[3].Visible = false;
        //    datagrid.Columns[4].Visible = false;

        //    datagrid.Columns[7].Visible = false;
        //    datagrid.Columns[8].Visible = false;

        //    datagrid.Columns[10].Visible = false;
        //    datagrid.Columns[12].Visible = false;
        //    datagrid.Columns[13].Visible = false;

        //    datagrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        //    datagrid.AutoResizeColumns();
        //}



        //private void GridMaker2(DataGridView datagrid, DataTable dt, string[] q1)
        //{
        //    var ds = new BindingSource();
        //    ds.DataSource = dt;

        //    DataGridViewTextBoxCell noCell = new DataGridViewTextBoxCell();

        //    DataGridViewComboBoxCell secondCell = new DataGridViewComboBoxCell();
        //    secondCell.Items.AddRange(new string[] { "N", "R", "U", "-" });

        //    DataGridViewComboBoxCell thirdCell = new DataGridViewComboBoxCell();
        //    thirdCell.Items.AddRange(new string[] { "Yes", "No", "-" });

        //    //DataGridViewComboBoxCell forthCell = new DataGridViewComboBoxCell();
        //    //forthCell.Items.AddRange(q1);


        //    DataGridViewColumn noColumn = new DataGridViewColumn(noCell);
        // //   DataGridViewColumn firstColumn = new DataGridViewColumn(firstCell);

        //    DataGridViewColumn secondColumn = new DataGridViewColumn(secondCell);
        //    DataGridViewColumn thirdColumn = new DataGridViewColumn(thirdCell);

        //    datagrid.Columns.Add(noColumn);
        //  //  datagrid.Columns.Add(firstColumn);

        //    datagrid.DataSource = ds;
        //    datagrid.Columns.Add(secondColumn);
        //    datagrid.Columns.Add(thirdColumn);

        //    datagrid.Columns[0].HeaderText = "No.";
        //    //datagrid.Columns[1].HeaderText = datagrid.Columns[3].HeaderText;
        //    datagrid.Columns[13].HeaderText = datagrid.Columns[3].HeaderText;
        //    datagrid.Columns[14].HeaderText = datagrid.Columns[6].HeaderText;

        //    foreach (DataGridViewRow item in datagrid.Rows)
        //    {


        //        if (item.Cells[2].Value != null)
        //            item.Cells[0].Value = item.Index + 1;
        //        //if (item.Cells[3].Value != null)
        //        //    item.Cells[1].Value = item.Cells[3].Value;
        //        if (item.Cells[3].Value != null)
        //            item.Cells[13].Value = item.Cells[3].Value;

        //        if (item.Cells[6].Value != null)
        //        {
        //            if (item.Cells[6].Value.ToString().Equals("Y"))
        //            {
        //                item.Cells[14].Value = "Yes";
        //            }
        //            else if (item.Cells[6].Value.ToString().Equals("N"))
        //            {
        //                item.Cells[14].Value = "No";
        //            }
        //            else
        //            {
        //                item.Cells[14].Value = "-";
        //            }
        //        }

        //        if (item.Cells[7].Value != null)
        //        {
        //            if (!item.Cells[7].Value.ToString().Equals("UNKNOWN"))
        //            {
        //                item.Cells[2].ReadOnly = true;
        //            }
        //        }

        //    }

        //    DisableColumn(datagrid, new int[] { 0, 3, 6, 7, 8, 9, 10 });

        //    datagrid.Columns[3].Visible = false;
        //    datagrid.Columns[6].Visible = false;
        //    datagrid.Columns[7].Visible = false;
        //    datagrid.Columns[9].Visible = false;
        //    datagrid.Columns[10].Visible = false;
        //    datagrid.Columns[11].Visible = false;
        //    datagrid.Columns[12].Visible = false;

        //    //datagrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        //    //datagrid.AutoResizeColumns();

        //    datagrid.Columns[0].Width = 30;
        //    datagrid.Columns[1].Width = 160;

        //}


        //private void GridMaker(DataGridView datagrid, DataTable dt, List<string> q1, List<string> q2)
        //{

        //    var ds = new BindingSource();
        //    ds.DataSource = dt;

        //    DataGridViewTextBoxCell noCell = new DataGridViewTextBoxCell();

        //    DataGridViewComboBoxCell firstCell = new DataGridViewComboBoxCell();
        //    firstCell.Items.AddRange(q1.ToArray());


        //    DataGridViewComboBoxCell secondCell = new DataGridViewComboBoxCell();
        //    secondCell.Items.AddRange(q2.ToArray());

        //    DataGridViewColumn noColumn = new DataGridViewColumn(noCell);
        //    DataGridViewColumn firstColumn = new DataGridViewColumn(firstCell);
        //    DataGridViewColumn secondColumn = new DataGridViewColumn(secondCell);


        //    datagrid.Columns.Add(noColumn);
        //    datagrid.Columns.Add(firstColumn);
        //    datagrid.Columns.Add(secondColumn);
        //    datagrid.DataSource = ds;
        //    datagrid.Columns[0].HeaderText = "No.";
        //    datagrid.Columns[1].HeaderText = datagrid.Columns[3].HeaderText;
        //    datagrid.Columns[2].HeaderText = datagrid.Columns[4].HeaderText;

        //    foreach (DataGridViewRow item in datagrid.Rows)
        //    {
        //        if (item.Cells[3].Value != null)
        //            item.Cells[0].Value = item.Index + 1;

        //        foreach (string itm1 in q1)
        //        {
        //            if (!(item.Cells[3].Value == null))
        //            {
        //                if (itm1.Length <= 3) 
        //                {
        //                    if (itm1 == item.Cells[3].Value.ToString())
        //                    item.Cells[1].Value = item.Cells[3].Value;
        //                }
        //                else if (itm1.Substring(0, 3).Equals(item.Cells[3].Value.ToString()))
        //                {
        //                    item.Cells[1].Value = itm1;
        //                    break;
        //                }
        //            }
        //        }

        //        foreach (string itm2 in q2)
        //        {
        //            if (!(item.Cells[4].Value == null))
        //            {
        //                if (itm2.Length <= 3) 
        //                {if(itm2 == item.Cells[4].Value.ToString())
        //                    item.Cells[2].Value = item.Cells[4].Value;
        //                }
        //                else if (itm2.Substring(0, 3).Equals(item.Cells[4].Value.ToString()))
        //                {
        //                    item.Cells[2].Value = itm2;
        //                    break;
        //                }
        //            }
        //        }
        //    }
        //    //datagrid.Columns[0].ReadOnly = true;
        //    //datagrid.Columns[6].ReadOnly = true;
        //    //datagrid.Columns[7].ReadOnly = true;
        //    //datagrid.Columns[8].ReadOnly = true;

        //    DisableColumn(datagrid, new int[] { 0, 6, 7, 8 });


        //    datagrid.Columns[3].Visible = false;
        //    datagrid.Columns[4].Visible = false;
        //    datagrid.Columns[8].Visible = false;

        //    datagrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        //    datagrid.AutoResizeColumns();



        //}

        //private List<Requisition> Requisitions()
        //{
        //    using (var dataBase = new SAP_Entities())
        //    {

        //        var reqList = (from a in dataBase.ODRFs
        //                       where a.U_ITWONo == jobNo && a.ObjType == "67"
        //                       join b in dataBase.DRF1 on a.DocEntry equals b.DocEntry
        //                       group b by new { seqNo = a.DocEntry, a.U_ITWONo, a.U_PartReqStatus, a.CreateDate } into bGrouped
        //                       select new Requisition
        //                       {
        //                           SeqNumber = bGrouped.Key.seqNo,
        //                           WorkOrderNumber = bGrouped.Key.U_ITWONo,
        //                           CreationDate = bGrouped.Key.CreateDate,
        //                           ItemsQuantity = bGrouped.Sum(x => x.Quantity),
        //                           ItemsType = bGrouped.ToList().Count(),
        //                           Status = bGrouped.Key.U_PartReqStatus
        //                       }).ToList();

        //        return reqList;
        //    }
        //}

        private void DisableColumn(DataGridView datagrid, int[] index)
        {
            foreach (var item in index)
            {
                datagrid.Columns[item].ReadOnly = true;
                datagrid.Columns[item].DefaultCellStyle.BackColor = Color.LightGray;
                // datagrid.Columns[item].DefaultCellStyle.ForeColor = Color.DarkGray;
                datagrid.Columns[item].DefaultCellStyle.SelectionBackColor = Color.LightGray;
                datagrid.Columns[item].DefaultCellStyle.SelectionForeColor = Color.Black;
            }
        }

    }

    class Requisition
    {
        public int? SeqNumber { get; set; }
        public int? WorkOrderNumber { get; set; }
        public int ItemsType { get; set; }
        public Decimal? ItemsQuantity { get; set; }
        public string Status { get; set; }
        public DateTime? CreationDate { get; set; }
        public int ReqisitionStatus { get; set; }

    }
}
