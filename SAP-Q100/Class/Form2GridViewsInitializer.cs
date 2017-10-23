using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SAP_Q100.Class
{
    class Form2GridViewsInitializer
    {

        //  private ODRFConnector odrfObject;
        private static SAPbobsCOM.Company pCompany;
        public SAPbobsCOM.Documents oInvTransDraft;
        private int jobNo;
        SqlData data;
        private int invDraftKey;
        private bool isAvailable;

        public Form2GridViewsInitializer(int jobNumber, int DraftObjectKey, SAPbobsCOM.Company company)
        {

            // odrfObject = new ODRFConnector(jobNumber,true);
            jobNo = jobNumber;
            pCompany = company;

            data = new SqlData();
            invDraftKey = DraftObjectKey;


            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            isAvailable=oInvTransDraft.GetByKey(invDraftKey);


        }

        public void DataGridView1Initializer(DataGridView datagrid1)
        {
           
            if (isAvailable)
            {

                //using (var dataBase = new SAP_Entities())
                //{
                  //  List<string> query1 = (from a in dataBase.OITMs where a.InvntItem == "Y" && a.SellItem == "Y" select a.ItemCode).ToList();


                    if (oInvTransDraft != null)//&& ordrObject.oOrder.DocumentStatus == SAPbobsCOM.BoStatus.bost_Open)
                    {
                        var dt1 = new DataTable();
                        dt1.Columns.Add("NO.", typeof(int));
                        dt1.Columns.Add("P/N", typeof(String));
                        dt1.Columns.Add("QTY", typeof(String));
                        dt1.Columns.Add("OH", typeof(String));
                        dt1.Columns.Add("SUBSTITUTE", typeof(String));
                        dt1.Columns.Add("MODEL", typeof(String));
                        dt1.Columns.Add("DESCRIPTION", typeof(string));
                        dt1.Columns.Add("VENDOR NO", typeof(String));
                        dt1.Columns.Add("ORGINAL REQUEST", typeof(String));
                        dt1.Columns.Add("AVAILABLE", typeof(String));

                        int counter = 1;
                        for (var i = 0; i < oInvTransDraft.Lines.Count; i++)
                        {
                            oInvTransDraft.Lines.SetCurrentLine(i);
                   
                            if (string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORRepCd").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORCause").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORSymCd").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORTestC").Value.ToString()) &&( oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value == "-"|| oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value == "Yes"|| oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value == ""))
                       
                            {
                                dt1.Rows.Add(new object[]
                            {
                           counter++,
                                oInvTransDraft.Lines.ItemCode=="UNKNOWN"? oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcItm").Value:oInvTransDraft.Lines.ItemCode,
                                oInvTransDraft.Lines.Quantity.ToString(),
                                System.Convert.ToInt32(data.ItemsByItemCode(oInvTransDraft.Lines.ItemCode).Rows[0]["OnHand"]),
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_Replaced").Value.ToString()=="1"?"Yes":"No",
                                data.ItemsByItemCode(oInvTransDraft.Lines.ItemCode).Rows[0]["ItemCode"].ToString(),
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value.ToString(),
                                oInvTransDraft.Lines.VisualOrder,
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_OrginalItemCode").Value.ToString(),
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value.ToString()

                            });
                            }
                        }
                        // /*oInvTransDraft.Lines.UserFields.Fields.Item("U_POAvail").Value,*/

                        datagrid1.DataSource = dt1;
                        DataGridViewComboBoxCell isAvailableCell = new DataGridViewComboBoxCell();
                        isAvailableCell.Items.AddRange("-","No", "Yes");
                        DataGridViewColumn availableColumn = new DataGridViewColumn(isAvailableCell);
                        datagrid1.Columns.Add(availableColumn);

                        DataGridViewComboBoxCell itemCodeCell = new DataGridViewComboBoxCell();
                        itemCodeCell.Items.AddRange(data.ItemCodesList());
                        DataGridViewColumn itemColumn = new DataGridViewColumn(itemCodeCell);
                        datagrid1.Columns.Add(itemColumn);

                     

                        datagrid1.Columns[10].HeaderText = "AVAILABLE";
                        datagrid1.Columns[11].HeaderText = "ALTERNATIVE";

                        //datagrid1.Columns[7].Visible = false;
                        //datagrid1.Columns[10].Visible = false;
                        datagrid1.Columns[9].Visible = false;

                    foreach (DataGridViewRow item in datagrid1.Rows)
                    {
                        if (item.Cells[5].Value != null)
                        {
                            if (!item.Cells[5].Value.ToString().Equals("UNKNOWN"))
                            {
                                item.Cells[6].ReadOnly = true;
                            }
                        }

                        if (item.Cells[9].Value != null)
                        {
                            if (item.Cells[9].Value.ToString().Equals("Yes"))
                                item.Cells[10].Value = item.Cells[9].Value.ToString();
                            else if (item.Cells[9].Value.ToString().Equals("No"))
                                item.Cells[10].Value = item.Cells[9].Value.ToString();
                            else { item.Cells[10].Value = "-"; }
                        }
                    }

                    datagrid1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        datagrid1.AutoResizeColumns();

                    }

                //}
            }

            DisableColumn(datagrid1, new int[] { 0,1,2, 3, 4, 5, 8 });
        }
        // ordrObject.oInvTransDraft.Lines.UserFields.Fields.Item("U_WORModel").Value,ordrObject.oInvTransDraft.Lines.ItemDescription




        public void DataGridView2Initializer(DataGridView datagrid2)
        {

            if (isAvailable)
            {

                //using (var dataBase = new SAP_Entities())
                //{
                  //  List<string> query1 = (from a in dataBase.OITMs where a.InvntItem == "Y" && a.SellItem == "Y" select a.ItemCode).ToList();

                  //  List<string> query2 = (from a in dataBase.OCRDs where a.CardType == "S" select a.CardCode).ToList();

                 //   var query3 = (from a in dataBase.OPORs join b in dataBase.ODRFs on a.U_RequisitionNumber equals b.DocEntry where b.ObjType == "67" && b.DocEntry==invDraftKey
                                           //&& b.U_ITWONo==jobNo join c in dataBase.DRF1 on b.DocEntry equals c.DocEntry  join d in 
                                           //dataBase.POR1 on c.VisOrder equals d.U_RequisitionLineNo where a.DocEntry == d.DocEntry   select new { c.VisOrder,a.DocNum }).ToList();

                    if (oInvTransDraft != null)//&& ordrObject.oOrder.DocumentStatus == SAPbobsCOM.BoStatus.bost_Open)
                    {
                        var dt2 = new DataTable();
                        dt2.Columns.Add("NO.", typeof(int));
                        dt2.Columns.Add("P/N", typeof(String));
                        dt2.Columns.Add("QTY", typeof(String));
                        dt2.Columns.Add("OH", typeof(String));
                        dt2.Columns.Add("SUBSTITUTE", typeof(String));
                        dt2.Columns.Add("MODEL", typeof(String));
                        dt2.Columns.Add("DESCRIPTION", typeof(string));
                        dt2.Columns.Add("VENDOR NO", typeof(String));
                        dt2.Columns.Add("ORGINAL REQUEST", typeof(String));
                        dt2.Columns.Add("PO Number", typeof(String));
                        dt2.Columns.Add("AVAILABLE", typeof(String));
                        dt2.Columns.Add("VENDOR CODE", typeof(String));
                        dt2.Columns.Add("ETA", typeof(String));
                    int counter = 1;
                        for (var i = 0; i < oInvTransDraft.Lines.Count; i++)
                        {
                            oInvTransDraft.Lines.SetCurrentLine(i);
                           
                            if (string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORRepCd").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORCause").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORSymCd").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORTestC").Value.ToString())&& oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value =="No")
                            {
                                dt2.Rows.Add(new object[]
                            {
                                counter++,
                                oInvTransDraft.Lines.ItemCode=="UNKNOWN"? oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcItm").Value:oInvTransDraft.Lines.ItemCode,
                                oInvTransDraft.Lines.Quantity.ToString(),
                                System.Convert.ToInt32(data.ItemsByItemCode(oInvTransDraft.Lines.ItemCode).Rows[0]["OnHand"]),
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_Replaced").Value.ToString()=="1"?"Yes":"No",
                                data.ItemsByItemCode(oInvTransDraft.Lines.ItemCode).Rows[0]["ItemCode"].ToString(),
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value.ToString(),
                                oInvTransDraft.Lines.VisualOrder,
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_OrginalItemCode").Value.ToString(),
                                "",
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value.ToString(),
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value.ToString(),
                                ""


                            });
                            }
                        }
                        // /*oInvTransDraft.Lines.UserFields.Fields.Item("U_POAvail").Value,*/





                        DataGridViewCheckBoxColumn myCheckedColumn = new DataGridViewCheckBoxColumn()
                        {
                            Name = "Checked Column",
                            FalseValue = 0,
                            TrueValue = 1,
                            Visible = true
                        };
                        // add the new column to your dataGridView 
                        datagrid2.Columns.Add(myCheckedColumn);



                        datagrid2.DataSource = dt2;



                        DataGridViewComboBoxCell isAvailableCell = new DataGridViewComboBoxCell();
                        isAvailableCell.Items.AddRange("No", "Yes");
                        DataGridViewColumn availableColumn = new DataGridViewColumn(isAvailableCell);
                        datagrid2.Columns.Add(availableColumn);


                        DataGridViewComboBoxCell itemCodeCell = new DataGridViewComboBoxCell();
                        itemCodeCell.Items.AddRange(data.ItemCodesList());
                        DataGridViewColumn itemColumn = new DataGridViewColumn(itemCodeCell);
                        datagrid2.Columns.Add(itemColumn);

                        DataGridViewComboBoxCell vendorCell = new DataGridViewComboBoxCell();
                    vendorCell.Items.AddRange(" ");
                    vendorCell.Items.AddRange(data.CardCodeList());
                        DataGridViewColumn vendor = new DataGridViewColumn(vendorCell);
                        datagrid2.Columns.Add(vendor);
                        

                        datagrid2.Columns[14].HeaderText = "AVAILABLE";
                        datagrid2.Columns[15].HeaderText = "ALTERNATIVE";
                        datagrid2.Columns[16].HeaderText = "VENDOR CODE";

                        datagrid2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        datagrid2.AutoResizeColumns();

                    //datagrid2.Columns[0].Visible = false;
                    //datagrid2.Columns[4].Visible = false;

                    datagrid2.Columns[8].Visible = false;
                    datagrid2.Columns[11].Visible = false;
                    datagrid2.Columns[12].Visible = false;
                    //datagrid2.Columns[13].Visible = false;
                    //datagrid2.Columns[14].Visible = false;

                    foreach (DataRow item in data.Form2Grid2Query1(invDraftKey).Rows)
                        {
                            foreach (DataGridViewRow itm in datagrid2.Rows)
                            {
                                if (System.Convert.ToInt32(item["VisOrder"]) == System.Convert.ToInt32(itm.Cells[8].Value))
                                {
                                    datagrid2.Rows[itm.Index].DefaultCellStyle.ForeColor = Color.Gray;
                                    datagrid2.Rows[itm.Index].ReadOnly = true;
                                    datagrid2.Rows[itm.Index].Cells[10].Value = item["DocNum"];
                                    datagrid2.Rows[itm.Index].Cells[13].Value = item["DocDueDate"];

                                //foreach (DataGridViewColumn it in dataGridView2.Columns)
                                //    dataGridView2.Rows[item].Cells[it.Index].ReadOnly = true;
                            }
                            }
                        }

                    foreach (DataGridViewRow item in datagrid2.Rows)
                    {
                        if (item.Cells[11].Value != null)
                        {
                            if (item.Cells[11].Value.ToString().Equals("Yes"))
                                item.Cells[14].Value = item.Cells[11].Value.ToString();
                            else if (item.Cells[11].Value.ToString().Equals("No"))
                                item.Cells[14].Value = item.Cells[11].Value.ToString();
                            else { item.Cells[14].Value = "-"; }
                        }
                    }


                    foreach (DataGridViewRow item in datagrid2.Rows)
                    {
                        if (item.Cells[12].Value != null)
                        {
                           
                                item.Cells[16].Value = item.Cells[12].Value.ToString();
                         
                        }
                    }

                    foreach (DataGridViewColumn item in datagrid2.Columns)
                        {

                            item.SortMode = DataGridViewColumnSortMode.NotSortable;
                        }

                    }



              //  }
            }

            DisableColumn(datagrid2, new int[] {1,2,3,4,5,6,8,9,10,13 });
        }

        private void DisableColumn(DataGridView datagrid, int[] index)
        {
            foreach (var item in index)
            {
                datagrid.Columns[item].ReadOnly = true;
                datagrid.Columns[item].DefaultCellStyle.BackColor = Color.LightGray;
              //  datagrid.Columns[item].DefaultCellStyle.ForeColor = Color.DarkGray;
                datagrid.Columns[item].DefaultCellStyle.SelectionBackColor = Color.LightGray;
                datagrid.Columns[item].DefaultCellStyle.SelectionForeColor = Color.Black;
            }
        }
    }
}
