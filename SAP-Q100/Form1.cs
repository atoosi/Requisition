using SAP_Q100.Class;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace SAP_Q100
{
    public partial class Form1 : Form
    {

        private int jobNo;
        private string userName;
        private int userId;
        private string tempUserName;

        private bool mouseClick1;
        private bool rowAdd1;
        private bool rowLeave1;


        private bool mouseClick2;
        private bool rowAdd2;
        private bool rowLeave2;

        private bool mouseClick3;
        private bool rowAdd3;
        private bool rowLeave3;


        public Form1GridViewsInitializer gd;
        private int delete_id;
        private string delete_Validator="";


        private static SAPbobsCOM.Company iCompany;
        private static SAPbobsCOM.Company pCompany;
        private SAPbouiCOM.Application SBO_Application;
        private int lresult;
        private string sresult;
        private SAPbobsCOM.Documents oOrder;
        public SAPbobsCOM.Documents oInvTransDraft;
        private int orderKey;
        private int invDraftKey;
        private SqlData data;
        private Form2 form2;
        private int reqNo;
        private bool autoComplete;
        private DataRow row;
        private DataRow rowR;
        public Dictionary<int, int> LineSelector;


        //BackgroundWorker backgroundWorker1 = new BackgroundWorker();


        //void Form1_Shown(object sender, EventArgs e)
        //{
        //    // Start the background worker
        //    backgroundWorker1.RunWorkerAsync();
        //}
        //// On worker thread so do our thing!
        //void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    // Your background task goes here
        //    for (int i = 0; i <= 100; i++)
        //    {
        //        // Report progress to 'UI' thread
        //        backgroundWorker1.ReportProgress(i);
        //        // Simulate long task
        //        System.Threading.Thread.Sleep(200);
        //    }
        //}
        // Back on the 'UI' thread so we can update the progress bar
        //void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        //{
        //    // The progress percentage is a property of e
        //    //progressBar1.Value = e.ProgressPercentage;

        //    tbProgress.TextChanged += new System.EventHandler(this.tbProgress_TextChanged);

          

        //}
    




    public Form1(int jobNO, int userID, int reqNO)
        {
    
            userId = userID;
            jobNo = jobNO;
            orderKey = 0;
            reqNo = reqNO;
   
            data = new SqlData();
            userName = data.UserName(userId);
            tempUserName = data.UserName(userId);
            invDraftKey = data.InventoryDraftKey(jobNo);
            sresult = "Okay";

           
            if (data.Form1WorkOrder1Query(jobNo) != null)
            {
                new Task(AsyncLoadingForExicting).Start();
                rowR = data.Form1WorkOrder1Query(jobNo).Rows[0];
                orderKey = System.Convert.ToInt32(rowR["DocEntry"]);
                
            }

            else
            {


                new Task(AsyncLoadingForNew).Start();


            }
            
            row = data.Form1MainQuery(jobNO).Rows[0];


            //Shown += new EventHandler(Form1_Shown);

            //// To report progress from the background worker we need to set this property
            //backgroundWorker1.WorkerReportsProgress = true;
            //// This event will be raised on the worker thread when the worker starts
            //backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            //// This event will be raised when we call ReportProgress
            //backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);

            //Cursor.Current = Cursors.WaitCursor;
            InitializeComponent();
            // Cursor.Current = Cursors.Default;

        }


        private void AsyncLoadingForNew()
        {
            SapConnector();
            oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            oOrder.DocNum = jobNo;
            oOrder.HandWritten = BoYesNoEnum.tYES;

            oOrder.CardCode = "C00950";
            oOrder.DocDueDate = DateTime.Now;

            oOrder.Lines.ItemCode = "_";
            oOrder.Lines.UserFields.Fields.Item("U_WORSymCd").Value = "505";
            oOrder.Lines.UserFields.Fields.Item("U_WORTestC").Value = "603";

            oOrder.Lines.Quantity = 1;
            oOrder.Lines.UnitPrice = 0;

            int IretCode = oOrder.Add();

            if (IretCode != 0)
            {
                MessageBox.Show(LastErrorMessage(IretCode));

            }

            else
            {
                rowR = data.Form1WorkOrder1Query(jobNo).Rows[0];
                orderKey = System.Convert.ToInt32(rowR["DocEntry"]);
                if (orderKey == 0)
                {

                    throw new System.InvalidOperationException("Work Order Part II cannot be Found");
                }
            }

            this.BeginInvoke(new MethodInvoker(() =>
            {

               gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                gd.DataGridView1Initializer(dataGridView1);
                gd.DataGridView2Initializer(dataGridView2);
                gd.DataGridView3Initializer(dataGridView3);

                this.LineSelector = gd.LineSelector;
            }));

            oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            if (reqNo != 0)
            {
                form2 = new Form2(jobNo, orderKey, reqNo, userId, pCompany, this);

                form2.ShowDialog();

                Close();
            }


            if (rowR["DocStatus"].Equals("C")/*|| data.ItemsByItemCode(userName).Rows.Count==0*/)
            {
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
                dataGridView3.Enabled = false;

            }
        }



        private void SapConnector()
        {

            #region Sap connection details

            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();

            // by following the steps specified above, the following
            // statment should be suficient for either development or run mode

            //sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            // connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString);

            // get an initialized application object

            SBO_Application = SboGuiApi.GetApplication(-1);

            iCompany = new SAPbobsCOM.Company();

            string sCookie = iCompany.GetContextCookie();
            string sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
            if (iCompany.Connected == true) iCompany.Disconnect();
            int setConnectionContext = iCompany.SetSboLoginContext(sConnectionContext);

            lresult = iCompany.Connect();

            //   if (iCompany.Connected) MessageBox.Show("connetec");

            pCompany = new SAPbobsCOM.Company();
            //Initialize the Company Object for the Connect method
            pCompany.Server = "10.10.1.8";
            pCompany.CompanyDB = iCompany.CompanyDB;
            pCompany.UserName = "ADO-SAP";
            pCompany.Password = "quest";
            pCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            pCompany.DbUserName = "sa1";
            pCompany.DbPassword = "s1ungod";
            pCompany.UseTrusted = false;
            pCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
            pCompany.LicenseServer = "10.10.1.8:30000";
            lresult = pCompany.Connect();
            #endregion

            if (lresult != 0)
            {
                sresult = ("ERROR: Connect has failed (Error " + lresult + ": " + pCompany.GetLastErrorDescription() + ")");
                //DialogResult result2 = MessageBox.Show(sResult, "Important Message", MessageBoxButtons.OK);
                throw new System.ArgumentException(sresult);

            }
        

        }

      

        private void AsyncLoadingForExicting()
        {
            //using (new CursorWait(true, true))
            //{

            SapConnector();

            this.BeginInvoke(new MethodInvoker(() =>
            {

               gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                gd.label = tbProgress;
                gd.DataGridView1Initializer(dataGridView1);
                gd.DataGridView2Initializer(dataGridView2);
                gd.DataGridView3Initializer(dataGridView3);

                this.LineSelector = gd.LineSelector;
            }));

            oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            if (reqNo != 0)
            {
                form2 = new Form2(jobNo, orderKey, reqNo, userId, pCompany, this);

                form2.ShowDialog();
                Close();

            }

            //     }
        }


        public class CursorWait : IDisposable
        {
            public CursorWait(bool appStarting = false, bool applicationCursor = false)
            {
                // Wait
                Cursor.Current = appStarting ? Cursors.AppStarting : Cursors.WaitCursor;
                if (applicationCursor) Application.UseWaitCursor = true;
            }

            public void Dispose()
            {
                // Reset
                Cursor.Current = Cursors.Default;
                Application.UseWaitCursor = false;
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

            mouseClick1 = false;
            rowAdd1 = false;
            rowLeave1 = false;

            mouseClick2 = false;
            rowAdd2 = false;
            rowLeave2 = false;

            mouseClick3 = false;
            rowAdd3 = false;
            rowLeave3 = false;
            autoComplete = false;
 

                label12.Text = string.IsNullOrEmpty(jobNo.ToString()) ? "" : jobNo.ToString();
                label13.Text = row["custmrName"] == null ? "" : row["custmrName"].ToString();
                label22.Text = row["U_BPCatNo"] == null ? "" : row["U_BPCatNo"].ToString();
                label23.Text = row["internalSN"] == null ? "" : row["internalSN"].ToString();
                label24.Text = row["FrgnName"] == null ? "" : row["FrgnName"].ToString();
                label25.Text = row["U_WOSalPer"] == null ? "" : row["U_WOSalPer"].ToString();
                label26.Text = row["U_WORcvdBy"] == null ? "" : row["U_WORcvdBy"].ToString();
                label27.Text = row["U_WORcvRmk"] == null ? "" : row["U_WORcvRmk"].ToString();
                label28.Text = row["U_WORecvDt"] == null ? "" : row["U_WORecvDt"].ToString();
                label44.Text = row["U_WOPONo"] == null ? "" : row["U_WOPONo"].ToString();
                label38.Text = row["ItemCode"] == null ? "" : row["ItemCode"].ToString();
                label34.Text = row["ItemName"] == null ? "" : row["ItemName"].ToString();
                label33.Text = row["FirmName"] == null ? "" : row["FirmName"].ToString();
                label32.Text = row["U_WOInbCar"] == null ? "" : row["U_WOInbCar"].ToString();
                label31.Text = row["U_WOArrDt"] == null ? "" : row["U_WOArrDt"].ToString();
                label30.Text = row["U_WODOM"] == null ? "" : row["U_WODOM"].ToString();
                label29.Text = row["closeDate"] == null ? "" : row["closeDate"].ToString();
                label120.Text = row["Name"] == null ? "" : row["Name"].ToString();
                label121.Text = row["U_WOSONo"] == null ? "" : row["U_WOSONo"].ToString();
                label122.Text = row["Expr2"] == null ? "" : row["Expr2"].ToString();
                label125.Text = userName;
                label128.Text = row["Name"].ToString();
                docDateLabel1.Text += "  " + row["subject"];

                if (row["U_WONaked"].Equals("Y"))
                    checkBox1.Checked = true;

                if (row["U_WOCabnt"].Equals("Y"))
                    checkBox2.Checked = true;

                if (row["U_WOSwivel"].Equals("Y"))
                    checkBox3.Checked = true;

                if (row["U_WOCord"].Equals("Y"))
                    checkBox4.Checked = true;

                if (row["U_WOVidCrd"].Equals("Y"))
                    checkBox5.Checked = true;

                if (row["U_WOBook"].Equals("Y"))
                    checkBox6.Checked = true;

                if (row["U_WORckCld"].Equals("Y"))
                    checkBox7.Checked = true;

                if (row["U_WORack"].Equals("Y"))
                    checkBox8.Checked = true;

                if (row["U_WOPCBMod"].Equals("Y"))
                    checkBox9.Checked = true;

                if (row["U_WOPatt"].Equals("Y"))
                    checkBox10.Checked = true;

                if (row["U_WOBroken"].Equals("Y"))
                    checkBox11.Checked = true;

                if (row["U_WOKeys"].Equals("Y"))
                    checkBox12.Checked = true;

                if (row["U_WOKeybrd"].Equals("Y"))
                    checkBox13.Checked = true;

                if (row["U_WOMouse"].Equals("Y"))
                    checkBox14.Checked = true;

                if (row["U_WOCmpSys"].Equals("Y"))
                    checkBox15.Checked = true;




                this.comboBox1.SelectedIndexChanged -= new EventHandler(comboBox1_SelectedIndexChanged);

                comboBox1.DataSource = data.UserCodeList();
                this.comboBox1.SelectedIndexChanged += new EventHandler(comboBox1_SelectedIndexChanged);

          

        }


        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (mouseClick1)
            {

                rowAdd1 = true;

            }
        }

        private void dataGridView1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (mouseClick1 && rowAdd1)
            {
                rowLeave1 = true;

            }

        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            mouseClick1 = true;
        }



        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                DataGridView datagrid = sender as DataGridView;
                if (datagrid.Rows[e.RowIndex].Cells[8].Value != null)
                    delete_id = Convert.ToInt32(datagrid.Rows[e.RowIndex].Cells[8].Value.ToString());

                if (datagrid.Rows[e.RowIndex].Cells[9].Value != null)
                    delete_Validator = datagrid.Rows[e.RowIndex].Cells[9].Value.ToString();

                this.contextMenuStrip1.Show(datagrid, e.Location);
                contextMenuStrip1.Show(Cursor.Position);
            }
        }
        #region Deleting section
        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool validator = true;
            if (!string.IsNullOrEmpty(delete_Validator))
            {

                foreach (DataGridViewRow item in dataGridView2.Rows)
                {
                    if (item.Cells[9].Value != null && item.Cells[9].Value.ToString().Equals(delete_Validator))
                    {
                        validator = false;

                    }
                }

                foreach (DataGridViewRow item in dataGridView3.Rows)
                {
                    if (item.Cells[13].Value != null && item.Cells[13].Value.ToString().Equals(delete_Validator))
                    {
                        validator = false;

                    }
                }
            }

                if (validator)
                {

                    // OrdrObject = new ORDRConnector(jobNo);
                    if (oOrder.GetByKey(orderKey))
                    {
                        oOrder.Lines.SetCurrentLine(delete_id);
                        int IretCode = 0;
                        if (oOrder.Lines.UserFields.Fields.Item("U_WORReqNo").Value.ToString() != "0")
                        {

                            if (form2 != null)
                                form2.Close();

                            oInvTransDraft = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                            oInvTransDraft.GetByKey(oOrder.Lines.UserFields.Fields.Item("U_WORReqNo").Value);
                            if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value.ToString() == "1")
                            {

                                int IretCode2;
                                if (oInvTransDraft.Lines.Count == 1)
                                {
                                    IretCode2 = oInvTransDraft.Remove();

                                }
                                else
                                {
                                    for (int i = 0; i < oInvTransDraft.Lines.Count; i++)
                                    {
                                        oInvTransDraft.Lines.SetCurrentLine(i);
                                        if (oInvTransDraft.Lines.LineNum == oOrder.Lines.UserFields.Fields.Item("U_U_RowIndex").Value)
                                        {
                                            oInvTransDraft.Lines.Delete();
                                            break;
                                        }
                                    }
                                    IretCode2 = oInvTransDraft.Update();
                                }
                                if (IretCode2 != 0)
                                {
                                    MessageBox.Show(LastErrorMessage(IretCode2));
                                    oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                    oOrder.GetByKey(orderKey);
                                }
                                //else
                                //{
                                //    if (form2 != null)
                                //        form2.Close();
                                //}

                                oOrder.Lines.Delete();
                                IretCode = oOrder.Update();

                            }
                            else
                            {

                                MessageBox.Show("You Cannot Remove Approved Item.");

                            }
                        }
                        else
                        {
                            oOrder.Lines.Delete();
                            IretCode = oOrder.Update();
                        }
                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                            oOrder.GetByKey(orderKey);
                        }
                        else
                        {
                            dataGridView1.DataSource = null;
                            dataGridView1.Columns.Clear();
                            dataGridView1.Refresh();

                            dataGridView2.DataSource = null;
                            dataGridView2.Columns.Clear();
                            dataGridView2.Refresh();

                            dataGridView3.DataSource = null;
                            dataGridView3.Columns.Clear();
                            dataGridView3.Refresh();

                            gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                            gd.DataGridView1Initializer(dataGridView1);
                            gd.DataGridView2Initializer(dataGridView2);
                            gd.DataGridView3Initializer(dataGridView3);
                        this.LineSelector = gd.LineSelector;
                    }

                    }
                }
                else
                {
                    MessageBox.Show("Delete Related Rows First");
                }
            
            delete_Validator = "";
        }
        #endregion
        private void dataGridView2_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                DataGridView datagrid = sender as DataGridView;
                if (datagrid.Rows[e.RowIndex].Cells[8].Value != null)
                    delete_id = Convert.ToInt32(datagrid.Rows[e.RowIndex].Cells[8].Value.ToString());
                this.contextMenuStrip1.Show(datagrid, e.Location);
                contextMenuStrip1.Show(Cursor.Position);

            }
        }

        private void dataGridView3_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                DataGridView datagrid = sender as DataGridView;
                if (datagrid.Rows[e.RowIndex].Cells[9].Value != null)
                    delete_id = Convert.ToInt32(datagrid.Rows[e.RowIndex].Cells[9].Value.ToString());
                this.contextMenuStrip1.Show(datagrid, e.Location);
                contextMenuStrip1.Show(Cursor.Position);

            }
        }

        private void dataGridView1_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {

            if (mouseClick1 && rowAdd1 && rowLeave1)
            {
                DialogResult result = MessageBox.Show("Add New Row?", "Important Question", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)

                {
                    DataGridView dgv = sender as DataGridView;
                    DataGridViewRow dr = dgv.Rows[e.RowIndex];
                    DialogResult popUpErrorMessage;

                    try
                    {
                        //using (var dataBase = new SAP_Entities())
                        //{
                        //    var oitmUserCheck = (from a in dataBase.OITMs where a.ItemCode == userName select a.ItemCode).FirstOrDefault();
                        var oitmUserCheck = data.UserCheck(tempUserName);
                        if (dr.Cells[1].Value == null)
                        {
                            popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[1].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgv.DataSource = null;
                                dgv.Columns.Clear();
                                dgv.Refresh();
                                gd.DataGridView1Initializer(dataGridView1);
                            }));
                        }
                        else if (dr.Cells[2].Value == null)
                        {
                            popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[2].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgv.DataSource = null;
                                dgv.Columns.Clear();
                                dgv.Refresh();
                                gd.DataGridView1Initializer(dataGridView1);
                            }));
                        }
                        //else if (dr.Cells[5].Value.ToString().Equals(""))
                        //{
                        //    popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[5].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                        //    this.BeginInvoke(new MethodInvoker(() =>
                        //    {
                        //        dgv.DataSource = null;
                        //        dgv.Columns.Clear();
                        //        dgv.Refresh();
                        //        gd.DataGridView1Initializer(dataGridView1);
                        //    }));
                        //}
                        else if (oitmUserCheck == null)
                        {
                            MessageBox.Show(" Please Define Your New User");
                            this.BeginInvoke(new MethodInvoker(() =>
                           {
                               dgv.DataSource = null;
                               dgv.Columns.Clear();
                               dgv.Refresh();
                               gd.DataGridView1Initializer(dataGridView1);
                           }));
                        }

                        else
                        {
                            var validate = true;

                            if (oOrder.GetByKey(orderKey))
                            {
                                if (oOrder.Lines.Count == 1)
                                {
                                    oOrder.Lines.SetCurrentLine(0);
                                if (oOrder.Lines.ItemCode == "_")
                                    {
                                        validate = false;
                                    }    
                                }
                                  
                                oOrder.Lines.Add();
                                //if (dgv.Rows[e.RowIndex].Cells[1].Value != null)
                                oOrder.Lines.UserFields.Fields.Item("U_WORSymCd").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString().Substring(0, 3);

                                //if (dgv.Rows[e.RowIndex].Cells[2].Value != null)
                                oOrder.Lines.UserFields.Fields.Item("U_WORTestC").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString().Substring(0, 3);
                                  //if (dgv.Rows[e.RowIndex].Cells[5].Value != null)
                                oOrder.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString().Equals("") ? 5 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[5].Value);

                                //oOrder.Lines.ItemCode = userName;
                                oOrder.Lines.ItemCode = oitmUserCheck;

                                oOrder.Lines.UserFields.Fields.Item("U_WORDate").Value = System.DateTime.Now;

                                int IretCode = oOrder.Update();
                                if (IretCode != 0)
                                {
                                    MessageBox.Show(LastErrorMessage(IretCode));
                                    oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                    oOrder.GetByKey(orderKey);

                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        dataGridView1.DataSource = null;
                                        dataGridView1.Columns.Clear();
                                        dataGridView1.Refresh();
                                        gd.DataGridView1Initializer(dataGridView1);
                                    }));

                                }
                                else
                                {
                                    if (!validate)
                                    {
                                        oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                        oOrder.GetByKey(orderKey);
                                        oOrder.Lines.SetCurrentLine(0);
                                        oOrder.Lines.Delete();
                                        
                                        oOrder.Update();
                                    }
                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        dataGridView1.DataSource = null;
                                        dataGridView1.Columns.Clear();
                                        dataGridView1.Refresh();

                                        dataGridView2.DataSource = null;
                                        dataGridView2.Columns.Clear();
                                        dataGridView2.Refresh();

                                        dataGridView3.DataSource = null;
                                        dataGridView3.Columns.Clear();
                                        dataGridView3.Refresh();

                                        gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                                        gd.DataGridView1Initializer(dataGridView1);
                                        gd.DataGridView2Initializer(dataGridView2);
                                        gd.DataGridView3Initializer(dataGridView3);
                                        this.LineSelector = gd.LineSelector;

                                    }));
                                }
                            }
                        }
                        //}
                    }
                    catch { }
                }
                else
                {
                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();
                        gd.DataGridView1Initializer(dataGridView1);
                    }));
                }
            }
            rowAdd1 = false;
            mouseClick1 = false;
            rowLeave1 = false;

        }



        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!rowAdd1 && mouseClick1)
            {
                DialogResult result = MessageBox.Show("Update Row?", "Important Question", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)

                {
                    DataGridView dgv = sender as DataGridView;
                    DataGridViewRow dr = dgv.Rows[e.RowIndex];
                    try
                    {
                        if (oOrder.GetByKey(orderKey))
                        {
                            oOrder.Lines.SetCurrentLine(System.Convert.ToInt32(dr.Cells[8].Value));
                            switch (e.ColumnIndex)
                            {
                                case 1:
                                    oOrder.Lines.UserFields.Fields.Item("U_WORSymCd").Value = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Length >= 3 ? dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Substring(0, 3) : dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                                    break;
                                case 2:
                                    oOrder.Lines.UserFields.Fields.Item("U_WORTestC").Value = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString().Substring(0, 3);
                                    break;
                                case 5:
                                    oOrder.Lines.Quantity = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "" ? 5 : System.Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[5].Value);
                                    break;
                            }

                            int IretCode = oOrder.Update();
                            if (IretCode != 0)
                            {
                                MessageBox.Show(LastErrorMessage(IretCode));
                                oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                oOrder.GetByKey(orderKey);
                            }
                        }
                    }
                    catch { }

                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();

                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Refresh();

                        dataGridView3.DataSource = null;
                        dataGridView3.Columns.Clear();
                        dataGridView3.Refresh();

                        gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                        gd.DataGridView1Initializer(dataGridView1);
                        gd.DataGridView2Initializer(dataGridView2);
                        gd.DataGridView3Initializer(dataGridView3);
                        this.LineSelector = gd.LineSelector;
                    }));
                }


                else
                {
                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();
                        gd.DataGridView1Initializer(dataGridView1);
                    }));


                }
            }

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            if (!rowAdd2 && mouseClick2)
            {
                DialogResult result = MessageBox.Show("Update Row?", "Important Question", MessageBoxButtons.YesNo);
                DataGridView dgv = sender as DataGridView;
                DataGridViewRow dr = dgv.Rows[e.RowIndex];
                if (result == DialogResult.Yes)

                {
                    try
                    {
                        if (oOrder.GetByKey(orderKey))
                        {
                            oOrder.Lines.SetCurrentLine(System.Convert.ToInt32(dr.Cells[8].Value));
                            switch (e.ColumnIndex)
                            {
                               
                                case 0:
                                    {
                                        foreach (var item in LineSelector)
                                        {
                                            if(item.Key.ToString().Equals(dgv.Rows[e.RowIndex].Cells[0].Value.ToString()))
                                            oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value = item.Value.ToString();
                                        }
                                    }
                                    break;
                                case 1:
                                    oOrder.Lines.UserFields.Fields.Item("U_WORRepCd").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString().Substring(0, 3);
                                    break;
                                case 2:

                                    oOrder.Lines.UserFields.Fields.Item("U_WORCause").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString().Length >= 3 ? dgv.Rows[e.RowIndex].Cells[2].Value.ToString().Substring(0, 3) : dgv.Rows[e.RowIndex].Cells[2].Value.ToString();
                                    break;

                                case 5:
                                    oOrder.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString() == "" ? 5 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[5].Value);
                                    break;

                                case 9:
                                    oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value = dgv.Rows[e.RowIndex].Cells[9].Value.ToString();
                                    break;
                            }
                        }
                    }
                    catch { }
                    int IretCode = oOrder.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        oOrder.GetByKey(orderKey);

                        this.BeginInvoke(new MethodInvoker(() =>
                        {
                            dataGridView2.DataSource = null;
                            dataGridView2.Columns.Clear();
                            dataGridView2.Refresh();
                            gd.DataGridView2Initializer(dataGridView2);
                        }));
                    }


                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();

                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Refresh();

                        dataGridView3.DataSource = null;
                        dataGridView3.Columns.Clear();
                        dataGridView3.Refresh();

                        gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                        gd.DataGridView1Initializer(dataGridView1);
                        gd.DataGridView2Initializer(dataGridView2);
                        gd.DataGridView3Initializer(dataGridView3);
                    }));
                }


                else
                {
                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dgv.DataSource = null;
                        dgv.Columns.Clear();
                        dgv.Refresh();
                        gd.DataGridView2Initializer(dgv);
                    }));


                }
            }

        }

        private void dataGridView2_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {

            if (mouseClick2 && rowAdd2 && rowLeave2)
            {

                DialogResult result = MessageBox.Show("Add New Row?", "Important Question", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)

                {
                    DataGridView dgv = sender as DataGridView;
                    DataGridViewRow dr = dgv.Rows[e.RowIndex];
                    DialogResult popUpErrorMessage;


                    //try
                    //{
                        //using (var dataBase = new SAP_Entities())
                        //{
                        //    var oitmUserCheck = (from a in dataBase.OITMs where a.ItemCode == userName select a.ItemCode).FirstOrDefault();
                        var oitmUserCheck = data.UserCheck(tempUserName);

                        if (dr.Cells[0].Value == null)
                        {
                            popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[0].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgv.DataSource = null;
                                dgv.Columns.Clear();
                                dgv.Refresh();
                                gd.DataGridView2Initializer(dataGridView2);
                            }));
                        }
                        else if (dr.Cells[1].Value == null)
                        {
                            popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[1].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgv.DataSource = null;
                                dgv.Columns.Clear();
                                dgv.Refresh();
                                gd.DataGridView2Initializer(dataGridView2);
                            }));
                        }
                        else if (dr.Cells[2].Value == null)
                        {
                            popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[2].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgv.DataSource = null;
                                dgv.Columns.Clear();
                                dgv.Refresh();
                                gd.DataGridView2Initializer(dataGridView2);
                            }));
                        }

                        else if (dr.Cells[5].Value.ToString().Equals(""))
                        {
                            popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[5].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgv.DataSource = null;
                                dgv.Columns.Clear();
                                dgv.Refresh();
                                gd.DataGridView2Initializer(dataGridView2);
                            }));
                        }

                        else if (oitmUserCheck == null)
                        {
                            MessageBox.Show(" Please Define Your New User");
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgv.DataSource = null;
                                dgv.Columns.Clear();
                                dgv.Refresh();
                                gd.DataGridView2Initializer(dataGridView2);
                            }));
                        }

                        else
                        {
                            if (oOrder.GetByKey(orderKey))
                            {
                                oOrder.Lines.Add();
                                //if (dgv.Rows[e.RowIndex].Cells[1].Value != null)
                                oOrder.Lines.UserFields.Fields.Item("U_WORRepCd").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString().Substring(0, 3);

                                //if (dgv.Rows[e.RowIndex].Cells[2].Value != null)
                                oOrder.Lines.UserFields.Fields.Item("U_WORCause").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString().Length>3? dgv.Rows[e.RowIndex].Cells[2].Value.ToString().Substring(0, 3): dgv.Rows[e.RowIndex].Cells[2].Value.ToString();

                                //if (dgv.Rows[e.RowIndex].Cells[5].Value != null)
                                oOrder.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString().Equals("") ? 5 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[5].Value);

                                // oOrder.Lines.ItemCode = userName;
                                oOrder.Lines.ItemCode = oitmUserCheck;

                          
                                    foreach(KeyValuePair<int,int> item in LineSelector)
                                    if (System.Convert.ToInt32(dgv.Rows[e.RowIndex].Cells[0].Value)==item.Key)
                                    {
                                            oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value = item.Value.ToString();
                                    }
                               


                                oOrder.Lines.UserFields.Fields.Item("U_WORDate").Value = System.DateTime.Now;



                                int IretCode = oOrder.Update();
                                if (IretCode != 0)
                                {
                                    MessageBox.Show(LastErrorMessage(IretCode));
                                    oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                    oOrder.GetByKey(orderKey);

                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        dataGridView2.DataSource = null;
                                        dataGridView2.Columns.Clear();
                                        dataGridView2.Refresh();
                                        gd.DataGridView2Initializer(dataGridView2);
                                    }));
                                }
                                else
                                {

                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        dataGridView1.DataSource = null;
                                        dataGridView1.Columns.Clear();
                                        dataGridView1.Refresh();

                                        dataGridView2.DataSource = null;
                                        dataGridView2.Columns.Clear();
                                        dataGridView2.Refresh();

                                        dataGridView3.DataSource = null;
                                        dataGridView3.Columns.Clear();
                                        dataGridView3.Refresh();

                                        gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                                        gd.DataGridView1Initializer(dataGridView1);
                                        gd.DataGridView2Initializer(dataGridView2);
                                        gd.DataGridView3Initializer(dataGridView3);
                                    }));
                                }
                            }
                        }
                        //}
                    //}
                    //catch { }
                }
                else
                {
                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Refresh();
                        gd.DataGridView2Initializer(dataGridView2);
                    }));
                }
            }
            rowAdd2 = false;
            mouseClick2 = false;
            rowLeave2 = false;
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            mouseClick2 = true;
        }

        private void dataGridView3_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            mouseClick3 = true;
        }

        #region adding to DataGridView 3
        private void dataGridView3_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            var validator = true;

            int IretCode;
            if (mouseClick3 && rowAdd3 && rowLeave3)
            {


                DataGridView dgv = sender as DataGridView;
                DataGridViewRow dr = dgv.Rows[e.RowIndex];
                DialogResult popUpErrorMessage;

                //try
                //{
                if (dr.Cells[0].Value == null)
                {
                   
                        e.Cancel = true;
                        dataGridView3.CurrentCell = dr.Cells[0];
                        validator = false;


                        popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[0].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);

                    
                }
                else if (dr.Cells[1].Value.ToString().Equals(""))
                {
                    e.Cancel = true;
                    dataGridView3.CurrentCell = dr.Cells[1];
                    validator = false;


                    popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[1].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                    //this.BeginInvoke(new MethodInvoker(() =>
                    //{
                    //    dgv.DataSource = null;
                    //    dgv.Columns.Clear();
                    //    dgv.Refresh();
                    //    gd.DataGridView3Initializer(dgv);
                    //}));
                }
                else if (dr.Cells[2].Value.ToString().Equals(""))
                {
                    e.Cancel = true;
                    dataGridView3.CurrentCell = dr.Cells[2];
                    validator = false;

                    popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[2].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                    //this.BeginInvoke(new MethodInvoker(() =>
                    //{
                    //    dgv.DataSource = null;
                    //    dgv.Columns.Clear();
                    //    dgv.Refresh();
                    //    gd.DataGridView3Initializer(dgv);
                    //}));
                }
                //else if (dr.Cells[4].Value.ToString().Equals(""))
                //{
                //    e.Cancel = true;
                //    dataGridView3.CurrentCell = dr.Cells[4];
                //    validator = false;
                //    if (!autoComplete)
                //        popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[4].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);

                //    //this.BeginInvoke(new MethodInvoker(() =>
                //    //{
                //    //    dgv.DataSource = null;
                //    //    dgv.Columns.Clear();
                //    //    dgv.Refresh();
                //    //    gd.DataGridView3Initializer(dgv);
                //    //}));

                //    autoComplete = false;

                //}
                else if (dr.Cells[5].Value.ToString().Equals(""))
                {
                    e.Cancel = true;
                    dataGridView3.CurrentCell = dr.Cells[5];
                    validator = false;
                    if (!autoComplete)
                        popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[5].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);
                    //this.BeginInvoke(new MethodInvoker(() =>
                    //{
                    //    dgv.DataSource = null;
                    //    dgv.Columns.Clear();
                    //    dgv.Refresh();
                    //    gd.DataGridView3Initializer(dgv);
                    //}));
                    autoComplete = false;
                }

                //else if (dr.Cells[13].Value == null)
                //{
                //    e.Cancel = true;
                //    dataGridView3.CurrentCell = dr.Cells[13];
                //    validator = false;
                //    popUpErrorMessage = MessageBox.Show("The " + dgv.Columns[13].HeaderText + " column cannot be empty ", "Important Message", MessageBoxButtons.OK);

                //    //this.BeginInvoke(new MethodInvoker(() =>
                //    //{
                //    //    dgv.DataSource = null;
                //    //    dgv.Columns.Clear();
                //    //    dgv.Refresh();
                //    //    gd.DataGridView3Initializer(dgv);
                //    //}));

                //}

                else
                {

                    DialogResult result = MessageBox.Show("Add New Row?", "Important Question", MessageBoxButtons.YesNo);

                    if (result == DialogResult.Yes)

                    {

                        #region draft status
                        //  OdrfObject = new ODRFConnector(jobNo, false);

                        oInvTransDraft = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                        if (oInvTransDraft.GetByKey(invDraftKey) && oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value.ToString() == "1")
                        {
                            oInvTransDraft.Lines.Add();
                            if (dr.Cells[1].Value != null)
                            {
                                var selectedItem = data.ItemsByItemCode(dr.Cells[1].Value.ToString());
                                if (selectedItem.Rows.Count > 0)
                                {
                                    oInvTransDraft.Lines.ItemCode = selectedItem.Rows[0]["ItemCode"].ToString();
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcItm").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();
                                }
                                else
                                {
                                    oInvTransDraft.Lines.ItemCode = "UNKNOWN";
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcItm").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();

                                }
                                if (selectedItem.Rows.Count == 0)
                                {
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_UnknownItem").Value = "1";
                                }
                                else
                                {
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_UnknownItem").Value = "0";

                                }
                            }
                            if (dr.Cells[2].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString();

                            oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value = "-";

                            if (dr.Cells[4].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_WORLoc").Value = string.IsNullOrEmpty(dr.Cells[4].Value.ToString()) ? "" : dgv.Rows[e.RowIndex].Cells[4].Value.ToString();

                            if (dr.Cells[5].Value != null)
                                oInvTransDraft.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString() == "" ? 1 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[5].Value);

                            if (dr.Cells[13].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_WORDis").Value = string.IsNullOrEmpty(dgv.Rows[e.RowIndex].Cells[13].Value.ToString()) ? "N" : dgv.Rows[e.RowIndex].Cells[13].Value.ToString();

                            if (dr.Cells[14].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_WORFix").Value = dgv.Rows[e.RowIndex].Cells[14].Value.ToString().Equals("YES") ? "Y" : "N";

                            oInvTransDraft.Lines.UserFields.Fields.Item("U_WORDate").Value = System.DateTime.Now;

                            oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "1";

                            //
                            IretCode = oInvTransDraft.Update();
                            if (IretCode != 0)
                            {
                                MessageBox.Show(LastErrorMessage(IretCode));
                                oInvTransDraft = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                this.BeginInvoke(new MethodInvoker(() =>
                                {
                                    dataGridView3.DataSource = null;
                                    dataGridView3.Columns.Clear();
                                    dataGridView3.Refresh();
                                    gd.DataGridView3Initializer(dataGridView3);
                                }));
                            }
                        }
                        else
                        {
                            oInvTransDraft = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                            oInvTransDraft.DocDate = System.DateTime.Now;
                            //this is id
                            oInvTransDraft.UserFields.Fields.Item("U_ITWONo").Value = jobNo.ToString();
                            oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "1";
                            oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value = "-";

                            //   OdrfObject.oInvTransDraft.Lines.UserFields.Fields.Item("U_WORReqNo").Value = OrdrObject.key;
                            //
                            if (dr.Cells[1].Value != null)
                            {
                                var selectedItem = data.ItemsByItemCode(dr.Cells[1].Value.ToString());
                                if (selectedItem.Rows.Count > 0)
                                {
                                    oInvTransDraft.Lines.ItemCode = selectedItem.Rows[0]["ItemCode"].ToString();
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcItm").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();
                                }
                                else
                                {
                                    oInvTransDraft.Lines.ItemCode = "UNKNOWN";
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcItm").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();

                                }
                            }
                            if (dr.Cells[2].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString();

                            if (dr.Cells[4].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_WORLoc").Value = dgv.Rows[e.RowIndex].Cells[4].Value.ToString();

                            if (dr.Cells[5].Value != null)
                                oInvTransDraft.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString() == "" ? 1 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[5].Value);

                            if (dr.Cells[13].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_WORDis").Value = string.IsNullOrEmpty(dgv.Rows[e.RowIndex].Cells[13].Value.ToString()) ? "N" : dgv.Rows[e.RowIndex].Cells[13].Value.ToString();

                            if (dr.Cells[14].Value != null)
                                oInvTransDraft.Lines.UserFields.Fields.Item("U_WORFix").Value = dgv.Rows[e.RowIndex].Cells[14].Value.ToString().Equals("YES") ? "Y" : "N";

                            oInvTransDraft.Lines.UserFields.Fields.Item("U_WORDate").Value = System.DateTime.Now;


                            IretCode = oInvTransDraft.Add();
                            if (IretCode != 0)
                            {
                                MessageBox.Show(LastErrorMessage(IretCode));
                                oInvTransDraft = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                            }
                            else
                            {
                                //using (var dataBase = new SAP_Entities())
                                //{
                                //    invDraftKey = (from a in dataBase.ODRFs where a.U_ITWONo == jobNo && a.U_PartReqStatus == "1" && a.ObjType == "67" select a.DocEntry).FirstOrDefault();
                                //}
                                invDraftKey = data.InventoryDraftKey(jobNo);
                                oInvTransDraft.GetByKey(invDraftKey);
                                oInvTransDraft.UserFields.Fields.Item("U_RequisitionNumber").Value = invDraftKey;
                                oInvTransDraft.Update();
                            }
                        }
                        if (IretCode == 0)
                        {
                            #endregion

                            if (form2 != null)
                            {
                                form2.Close();
                            }

                            if (oOrder.GetByKey(orderKey))
                            {
                                oOrder.Lines.Add();
                                //if (dr.Cells[1].Value != null)
                                var selectedItem = data.ItemsByItemCode(dr.Cells[1].Value.ToString());
                                if (selectedItem.Rows.Count > 0)
                                {
                                    oOrder.Lines.ItemCode = selectedItem.Rows[0]["ItemCode"].ToString();
                                    oOrder.Lines.UserFields.Fields.Item("U_POSrvcItm").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();
                                }
                                else
                                {
                                    oOrder.Lines.ItemCode = "UNKNOWN";
                                    oOrder.Lines.UserFields.Fields.Item("U_POSrvcItm").Value = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();

                                }

                                if (dr.Cells[2].Value != null)
                                    oOrder.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString();

                                //            if (dr.Cells[5].Value != null)
                                oOrder.Lines.UserFields.Fields.Item("U_WORLoc").Value = dgv.Rows[e.RowIndex].Cells[4].Value.ToString();
                                //            if (dr.Cells[6].Value != null)
                                oOrder.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString() == "" ? 1 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[5].Value);
                                if (dr.Cells[13].Value != null)
                                    oOrder.Lines.UserFields.Fields.Item("U_WORDis").Value = string.IsNullOrEmpty(dgv.Rows[e.RowIndex].Cells[13].Value.ToString()) ? "N" : dgv.Rows[e.RowIndex].Cells[13].Value.ToString();

                                if (dr.Cells[14].Value != null)
                                    oOrder.Lines.UserFields.Fields.Item("U_WORFix").Value = dgv.Rows[e.RowIndex].Cells[14].Value.ToString().Equals("YES") ? "Y" : "N";
                                //    
                                oOrder.Lines.UserFields.Fields.Item("U_WORDate").Value = System.DateTime.Now;

                                oOrder.Lines.UserFields.Fields.Item("U_WORReqNo").Value = oInvTransDraft.DocEntry;
                                oOrder.Lines.UserFields.Fields.Item("U_WORReqQty").Value = dgv.Rows[e.RowIndex].Cells[11].Value.ToString();
                                oOrder.Lines.UserFields.Fields.Item("U_WORDelQty").Value = "0";
                                oOrder.Lines.UserFields.Fields.Item("U_U_RowIndex").Value = oInvTransDraft.Lines.LineNum;


                                foreach (KeyValuePair<int, int> item in LineSelector)
                                    if (System.Convert.ToInt32(dgv.Rows[e.RowIndex].Cells[0].Value) == item.Key)
                                    {
                                        oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value = item.Value.ToString();
                                    }

                                IretCode = oOrder.Update();


                                if (IretCode != 0)
                                {
                                    MessageBox.Show(LastErrorMessage(IretCode));
                                    oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                    oOrder.GetByKey(orderKey);

                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        dataGridView3.DataSource = null;
                                        dataGridView3.Columns.Clear();
                                        dataGridView3.Refresh();
                                        gd.DataGridView3Initializer(dataGridView3);
                                    }));

                                }
                                else
                                {

                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        dataGridView1.DataSource = null;
                                        dataGridView1.Columns.Clear();
                                        dataGridView1.Refresh();

                                        dataGridView2.DataSource = null;
                                        dataGridView2.Columns.Clear();
                                        dataGridView2.Refresh();

                                        dataGridView3.DataSource = null;
                                        dataGridView3.Columns.Clear();
                                        dataGridView3.Refresh();

                                        gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                                        gd.DataGridView1Initializer(dataGridView1);
                                        gd.DataGridView2Initializer(dataGridView2);
                                        gd.DataGridView3Initializer(dataGridView3);

                                    }));

                                }

                            }

                        }

                    }


                    else
                    {
                        this.BeginInvoke(new MethodInvoker(() =>
                        {
                            dataGridView3.DataSource = null;
                            dataGridView3.Columns.Clear();
                            dataGridView3.Refresh();
                            gd.DataGridView3Initializer(dataGridView3);
                        }));

                    }

                }
                //}


                //catch { }



            }
            if (validator)
            {
                rowAdd3 = false;
                mouseClick3 = false;
                rowLeave3 = false;
            }
        }
        #endregion

        #region Edit DataGridView 3
        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!rowAdd3 && mouseClick3)
            {
                DialogResult result = MessageBox.Show("Update Row?", "Important Question", MessageBoxButtons.YesNo);
                DataGridView dgv = sender as DataGridView;
                DataGridViewRow dr = dgv.Rows[e.RowIndex];

                if (result == DialogResult.Yes)

                {
                    if (oOrder.GetByKey(orderKey))
                    {

                        oOrder.Lines.SetCurrentLine(System.Convert.ToInt32(dr.Cells[9].Value));
                        try
                        {
                            switch (e.ColumnIndex)
                            {
                                case 0:
                                    {
                                        foreach (var item in LineSelector)
                                        {
                                            if (item.Key.ToString().Equals(dgv.Rows[e.RowIndex].Cells[0].Value.ToString()))
                                                oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value = item.Value.ToString();
                                        }
                                    }
                                    break;

                                case 1:
                                    oOrder.Lines.ItemCode = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();
                                    oInvTransDraft.Lines.ItemCode = dgv.Rows[e.RowIndex].Cells[1].Value.ToString();
                                    break;

                                case 2:
                                    if (!string.IsNullOrEmpty(dr.Cells[2].Value.ToString()))

                                    oOrder.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString();
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[2].Value.ToString();
                                    break;

                                case 4:
                                    oOrder.Lines.UserFields.Fields.Item("U_WORLoc").Value = dgv.Rows[e.RowIndex].Cells[4].Value.ToString();
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_WORLoc").Value = dgv.Rows[e.RowIndex].Cells[4].Value.ToString();
                                    break;
                                case 5:
                                    oOrder.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString() == "" ? 1 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[6].Value);
                                    oInvTransDraft.Lines.Quantity = dgv.Rows[e.RowIndex].Cells[5].Value.ToString() == "" ? 1 : System.Convert.ToDouble(dgv.Rows[e.RowIndex].Cells[6].Value);


                                    break;

                                case 9:
                                    oOrder.Lines.UserFields.Fields.Item("U_LineNumber").Value = dgv.Rows[e.RowIndex].Cells[9].Value.ToString();
                                    break;

                                case 13:
                                    if (!string.IsNullOrEmpty(dr.Cells[13].Value.ToString()))

                                        oOrder.Lines.UserFields.Fields.Item("U_WORDis").Value = dgv.Rows[e.RowIndex].Cells[13].Value.ToString();
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_WORDis").Value = dgv.Rows[e.RowIndex].Cells[13].Value.ToString();
                                    break;

                                case 14:
                                    oOrder.Lines.UserFields.Fields.Item("U_WORFix").Value = dgv.Rows[e.RowIndex].Cells[14].Value.ToString().Substring(0, 1);
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_WORFix").Value = dgv.Rows[e.RowIndex].Cells[14].Value.ToString().Substring(0, 1);

                                    break;
                            }
                        }
                        catch { }
                    }
                    int IretCode = oOrder.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oOrder = (SAPbobsCOM.Documents)iCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        oOrder.GetByKey(orderKey);
                    }

                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();

                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Refresh();

                        dataGridView3.DataSource = null;
                        dataGridView3.Columns.Clear();
                        dataGridView3.Refresh();

                        gd = new Form1GridViewsInitializer(jobNo, orderKey, pCompany);gd.label = tbProgress;
                        gd.DataGridView1Initializer(dataGridView1);
                        gd.DataGridView2Initializer(dataGridView2);
                        gd.DataGridView3Initializer(dataGridView3);
                    }));
                }

                else
                {
                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        dgv.DataSource = null;
                        dgv.Columns.Clear();
                        dgv.Refresh();
                        gd.DataGridView3Initializer(dgv);
                    }));
                }
            }
        }

        #endregion
        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

            if (mouseClick2)
            {

                rowAdd2 = true;
            }
        }

        private void dataGridView2_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (mouseClick2 && rowAdd2)
            {
                rowLeave2 = true;
            }
        }

        private void dataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (mouseClick3)
            {

                rowAdd3 = true;
            }
        }

        private void dataGridView3_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (mouseClick3 && rowAdd3)
            {
                rowLeave3 = true;
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (itemControl1.SelectedTab == itemControl1.TabPages["tabPage2"])//your specific tabname
            {
                gd.StatusChangerForm2();
                gd.DataGridView4Initializer(dataGridView4);
            }

            if (itemControl1.SelectedTab == itemControl1.TabPages["tabPage3"])//your specific tabname
            {
                gd.DataGridView5Initializer(dataGridView5);
            }


        }

        private void dataGridView4_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            DataGridViewRow dr = dgv.Rows[e.RowIndex];
            if (form2 == null)
            {
                form2 = new Form2(jobNo, orderKey, System.Convert.ToInt32(dr.Cells[1].Value), userId, pCompany, this);
                form2.Show();
                //this.Hide();
            }
            else
            {
                form2.Close();
                form2 = new Form2(jobNo, orderKey, System.Convert.ToInt32(dr.Cells[1].Value), userId, pCompany, this);
                form2.Show();
            }
        }

        public string LastErrorMessage(int IretCode)
        {
            string sErr = "";
            iCompany.GetLastError(out IretCode, out sErr);
            return sErr;
        }



        private void FillItems(AutoCompleteStringCollection autoItemsPopUp)
        {

            autoItemsPopUp.AddRange(data.ItemCodesList());
        }



        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //string header = dataGridView3.Columns[1].HeaderText;

            //if (header.Equals("NAME OR P/N"))
            if (dataGridView3.CurrentCell.ColumnIndex == 1)
            {
                TextBox auto = e.Control as TextBox;
                if (auto != null)
                {
                    auto.AutoCompleteMode = AutoCompleteMode.Suggest;
                    auto.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    AutoCompleteStringCollection autoItems = new AutoCompleteStringCollection();
                    FillItems(autoItems);
                    auto.AutoCompleteCustomSource = autoItems;
                    autoComplete = true;
                }
            }
            else
            {
                TextBox prodCode = e.Control as TextBox;
                if (prodCode != null)
                {
                    prodCode.AutoCompleteMode = AutoCompleteMode.None;
                }
            }
        }



        private void dataGridView3_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {

                if (data.ItemsByItemCode(dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()).Rows.Count != 0)
                {
                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Value = data.ItemsByItemCode(dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()).Rows[0]["FrgnName"];
                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].ReadOnly = true;
                    //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Style.BackColor = Color.LightGray;
                    //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Style.SelectionBackColor = Color.LightGray;
                    //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Style.SelectionForeColor = Color.Black;


                }
            }
       
        }
         

        //private void textBox1_TextChanged(object sender, EventArgs e)
        //{
        //    this.textBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
        //    this.textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;

        //    TextBox t = sender as TextBox;
        //    if (t != null)
        //    {
        //        //say you want to do a search when user types 3 or more chars
        //        if (t.Text.Length >= 3)
        //        {
        //            //SuggestStrings will have the logic to return array of strings either from cache/db
        //            string[] arr = data.UserCodeList();

        //            AutoCompleteStringCollection collection = new AutoCompleteStringCollection();
        //            collection.AddRange(arr);

        //            this.textBox1.AutoCompleteCustomSource = collection;
        //        }
        //    }
        //}

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    tempUserName = textBox1.Text;
        //}

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            tempUserName = comboBox1.SelectedValue.ToString();
        }

        private void tbProgress_TextChanged(object sender, EventArgs e)
        {

          
            switch (tbProgress.Text)
            {
                case "Loading":
                    tbProgress.Text = "Loading.";
                    tbProgress.Refresh();
                    break;
                case "Loading.":
                    tbProgress.Text = "Loading..";
                    tbProgress.Refresh();
                    break;
                case "Loading..":
                    tbProgress.Text = "Loading...";
                    tbProgress.Refresh();
                    break;
                case "Loading...":
                    tbProgress.Text = "Loading";
                    tbProgress.Refresh();
                    break;
            }
        }

        private void dataGridView3_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            itemControl1.Enabled = true;
        }
    }
}
