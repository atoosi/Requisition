using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAP_Q100.Class;
using SAPbobsCOM;
using System.IO;
using System.Data.SqlClient;
using System.Drawing.Printing;
using SAP_Q100.B1WS_Login;
using SAP_Q100.B1WS_Messages;

namespace SAP_Q100
{
    public partial class Form2 : Form
    {
        private int jobNumber;
        private int ORDRkey;
        private int ODRFkey;
        private string userName;
        private int userId;
        //private ODRFConnector odrfObject;
        private Form2GridViewsInitializer gridBuilder;

        public static SAPbobsCOM.CompanyService oCmpSrv;
        private static SAPbobsCOM.Company pCompany;
        public SAPbobsCOM.Documents oInvTransDraft;
        public SAPbobsCOM.Documents oPoDraft;
        private SAPbobsCOM.Documents oOrder;

        const string imageLocation = @"\\Sap\b1_shr\Inventory Pictures\";
        private bool isAvailable;
        private bool cellMouseClick1 = false;
        private bool cellValueChange1 = false;

        private bool cellMouseClick2 = false;
        private bool cellValueChange2 = false;
        private Form1 form1;
        private SAPbouiCOM.Application SBO_Application;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.EditText oEdit;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.LinkedButton linkedButton;
        private SqlData data;
        private SAPbouiCOM.Matrix oMatrix;
        private Form3 form3;
        public int existingPo;
        public bool display;
        private int delete_id;
        Bitmap memoryImage;


        public Form2(int JobNo, int workOrderPart1Key, int workOrderPart2Key, int userID, SAPbobsCOM.Company company, Form1 frm)
        {
            InitializeComponent();
            jobNumber = JobNo;
            ORDRkey = workOrderPart1Key;
            ODRFkey = workOrderPart2Key;
            userId = userID;
            pCompany = company;
            form1 = frm;


            data = new SqlData();
            //using (var dataBase = new SAP_Entities())
            //{
            //    userName = userName = (from i in dataBase.OUSRs where i.USERID == userId select i.U_NAME).FirstOrDefault();
            //}

            userName = data.UserName(userId);
            // odrfObject = new ODRFConnector(ODRFkey,true);



        }


        private void Form2_Load(object sender, EventArgs e)
        {

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


            gridBuilder = new Form2GridViewsInitializer(jobNumber, ODRFkey, pCompany);
            gridBuilder.DataGridView1Initializer(dataGridView1);

            gridBuilder.DataGridView2Initializer(dataGridView2);







            FieldsLoader();
            Validator();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            //using (var db = new SAP_Entities())
            //{

            //    var passQuery = (from a in db.OHEMs where a.passportNo == textBox2.Text && a.userId == userId && a.userId == 40 join b in db.HTM1 on a.empID equals b.empID select new { a.empID, a.firstName, a.passportNo, b.teamID }).FirstOrDefault();
            var passQuery = data.Form2Query1(userId, textBox2.Text);

            if (passQuery.Rows.Count > 0 && userId == 40)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {
                    oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "4";
                    oInvTransDraft.UserFields.Fields.Item("U_POReqAp2").Value = "1";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApDa2").Value = (DateTime?)DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApTm2").Value = DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApIm2").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApUsr2").Value = userName;
                    oInvTransDraft.DocumentsOwner = 34;

                    oInvTransDraft.Update();
                    //
                    int IretCode = oInvTransDraft.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                        isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                    }

                    else
                    {
                        //    button1.Visible = false;
                        //    button1.Enabled = false;

                        //    textBox2.Enabled = false;
                        //    textBox2.Visible = false;

                        //    label23.Visible = false;


                        //    textBox7.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApDa2").Value.ToString("MMM ddd d yyyy");
                        //    textBox8.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr2").Value;
                        //    textBox9.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApTm2").Value.ToString("HH : mm");
                        //    pictureBox1.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm2").Value;
                        //    pictureBox1.Visible = true;

                        form1.gd.DataGridView4Initializer(form1.dataGridView4);
                        FieldsLoader();
                        Validator();


                        SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "34", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());

                        LocationClass.Changer(pCompany, jobNumber, 34, 26);


                    }
                }

            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
            //}

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //using (var db = new SAP_Entities())
            //{

            //    var passQuery = (from a in db.OHEMs where a.passportNo == textBox4.Text && a.userId == userId join b in db.HTM1 on a.empID equals b.empID select new { a.empID, a.firstName, a.passportNo, b.teamID }).FirstOrDefault();
            //    if (passQuery != null)
            //    {
            var passQuery = data.Form2Query1(userId, textBox4.Text);
            if (passQuery.Rows.Count > 0)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {
                    oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "3";
                    oInvTransDraft.UserFields.Fields.Item("U_POReqAp1").Value = "1";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApDa1").Value = (DateTime?)DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApTm1").Value = DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApUsr1").Value = userName;
                    oInvTransDraft.DocumentsOwner = 40;

                    //
                    int IretCode = oInvTransDraft.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                        isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                    }

                    else
                    {


                        form1.gd.DataGridView4Initializer(form1.dataGridView4);
                        FieldsLoader();
                        Validator();

                        sendNotification("test**New Requisition Form has been Sent**test", "test**New Requisition form has been sent and needs your approval**test ", "40", "tet**Requisition**test", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());


                        // SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval ", "40", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());


                        LocationClass.Changer(pCompany, jobNumber, 40);
                    }

                }
            }
            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }


        }



        private void sendNotification(string msgSubject, string msgBody, string msgUserID, string msgDraft, string msgFirstLine, string objectType, string objectKey, string jobNumber)

        {
            var msgUserName = GetUserName(msgUserID);
            LoginServiceSoapClient login = new LoginServiceSoapClient();
            string sessionID = login.Login("10.10.1.8", "Quest_train", "dst_MSSQL2005", "sa", "s1ungod", "ado-sap", "quest", "ln_English", "10.10.1.8:30000");

            B1WS_Messages.Message message = new B1WS_Messages.Message();
            MessagesServiceSoapClient messageService = new MessagesServiceSoapClient();
            SendMessage sendMessage = new SendMessage();
            B1WS_Messages.MsgHeader msgHeader = new B1WS_Messages.MsgHeader();

            //add message header and lines
            message.User = -1;
            message.Subject = msgSubject;
            message.AttachmentSpecified = true;

            message.MessageDataColumns = new MessageMessageDataColumn[4];
            message.MessageDataColumns[2] = new MessageMessageDataColumn();
            message.MessageDataColumns[2].ColumnName = "Customer approval for";
            message.MessageDataColumns[2].LinkSpecified = false;
            message.MessageDataColumns[2].MessageDataLines = new MessageMessageDataColumnMessageDataLine[1];

            message.MessageDataColumns[3] = new MessageMessageDataColumn();
            message.MessageDataColumns[3].ColumnName = "Office Instruction";
            message.MessageDataColumns[3].LinkSpecified = false;
            message.MessageDataColumns[3].MessageDataLines = new MessageMessageDataColumnMessageDataLine[1];

            message.MessageDataColumns[0] = new MessageMessageDataColumn();
            message.MessageDataColumns[0].ColumnName = "Requisition No.";
            message.MessageDataColumns[0].LinkSpecified = false;
            message.MessageDataColumns[0].Link = MessageMessageDataColumnLink.tYES;
            message.MessageDataColumns[0].MessageDataLines = new MessageMessageDataColumnMessageDataLine[1];

            message.MessageDataColumns[1] = new MessageMessageDataColumn();
            message.MessageDataColumns[1].ColumnName = "Job No.";
            message.MessageDataColumns[1].LinkSpecified = true;
            message.MessageDataColumns[1].Link = MessageMessageDataColumnLink.tYES;
            message.MessageDataColumns[1].MessageDataLines = new MessageMessageDataColumnMessageDataLine[1];


            message.MessageDataColumns[2].MessageDataLines[0] = new MessageMessageDataColumnMessageDataLine();
            message.MessageDataColumns[2].MessageDataLines[0].Value = "test";
            message.MessageDataColumns[3].MessageDataLines[0] = new MessageMessageDataColumnMessageDataLine();
            message.MessageDataColumns[3].MessageDataLines[0].Value = "test";
            message.MessageDataColumns[0].MessageDataLines[0] = new MessageMessageDataColumnMessageDataLine();
            message.MessageDataColumns[0].MessageDataLines[0].Object = "23";
            message.MessageDataColumns[0].MessageDataLines[0].ObjectKey = objectKey;
            message.MessageDataColumns[0].MessageDataLines[0].Value = objectKey;
            message.MessageDataColumns[1].MessageDataLines[0] = new MessageMessageDataColumnMessageDataLine();
            message.MessageDataColumns[1].MessageDataLines[0].Object = "191";
            message.MessageDataColumns[1].MessageDataLines[0].ObjectKey = jobNumber;
            message.MessageDataColumns[1].MessageDataLines[0].Value = jobNumber;


            message.RecipientCollection = new MessageRecipient[1];
            message.RecipientCollection[0] = new MessageRecipient();
            message.RecipientCollection[0].UserCode = msgUserName;
            message.RecipientCollection[0].SendInternal = MessageRecipientSendInternal.tYES;
            message.RecipientCollection[0].SendInternalSpecified = true;
            message.RecipientCollection[0].SendFaxSpecified = false;
            message.RecipientCollection[0].SendEmailSpecified = false;

             

            sendMessage.Message = message;
            msgHeader.SessionID = sessionID;
            msgHeader.ServiceNameSpecified = true;
            msgHeader.ServiceName = B1WS_Messages.MsgHeaderServiceName.MessagesService;

            SendMessageResponse sendMessageResponse = new SendMessageResponse();
            sendMessageResponse.MessageHeader = messageService.SendMessage(msgHeader, sendMessage);
            //MessageBox.Show("Final Message# Returned by SAP:" + sendMessageResponse.MessageHeader.Code.ToString());


        }



        public void SendMessage(string msgSubject, string msgBody, string msgUserID, string msgDraft, string msgFirstLine, string objectType, string objectKey, string jobNumber)
        {
            SAPbobsCOM.Message oMessage = null;
            MessageDataColumns pMessageDataColumns = null;
            MessageDataColumn pMessageDataColumn = null;

            MessageDataLines oLines = null;
            MessageDataLine oLine = null;
            RecipientCollection oRecipientCollection = null;

            try
            {
                oCmpSrv = pCompany.GetCompanyService();
                MessagesService oMessageService = (SAPbobsCOM.MessagesService)oCmpSrv.GetBusinessService(ServiceTypes.MessagesService);


                // get the data interface for the new message
                oMessage = ((SAPbobsCOM.Message)(oMessageService.GetDataInterface(MessagesServiceDataInterfaces.msdiMessage)));

                // fill subject
                oMessage.Subject = msgSubject;

                oMessage.Text = objectKey + "@" + jobNumber + "@" + msgUserID + "@";

                // Add Recipient 
                oRecipientCollection = oMessage.RecipientCollection;

                oRecipientCollection.Add();

                // send internal message
                oRecipientCollection.Item(0).SendInternal = BoYesNoEnum.tYES;

                // add existing user name
                oRecipientCollection.Item(0).UserCode = GetUserCode(data.UserName(System.Convert.ToInt32(msgUserID)));

                // oRecipientCollection.Item(0).UserCode = msgUserCode;

                // get columns data
                pMessageDataColumns = oMessage.MessageDataColumns;

                // get column
                pMessageDataColumn = pMessageDataColumns.Add();
                // set column name
                pMessageDataColumn.ColumnName = msgDraft;

                // set link to a real object in the application
                pMessageDataColumn.Link = BoYesNoEnum.tNO;


                //******************************************************************
                // get lines
                oLines = pMessageDataColumn.MessageDataLines;
                // add new line
                oLine = oLines.Add();
                // set the line value
                oLine.Value = msgFirstLine;

                // set the link to BusinessPartner (the object type for Bp is 2)
                oLine.Object = objectType;
                // set the Bp code
                oLine.ObjectKey = objectKey;


                //*************************************************************

                // get column
                pMessageDataColumn = pMessageDataColumns.Add();

                // set column name
                pMessageDataColumn.ColumnName = "Job Number";

                // set link to a real object in the application
                pMessageDataColumn.Link = BoYesNoEnum.tYES;

                // send the message

                //*************************************************************

                // get lines
                oLines = pMessageDataColumn.MessageDataLines;
                // add new line
                oLine = oLines.Add();
                // set the line value
                oLine.Value = jobNumber;

                // set the link to BusinessPartner (the object type for Bp is 2)
                oLine.Object = "191";
                // set the Bp code
                oLine.ObjectKey = jobNumber;

                //*****************************************************************
                oMessageService.SendMessage(oMessage);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, null);
            }
        }

        private string GetUserCode(string sUserName)
        {

            Users oUsers;
            SAPbobsCOM.Recordset oRecordSet;

            try
            {
                //get users object
                oUsers = (SAPbobsCOM.Users)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

                // Get a new Recordset object
                oRecordSet = (SAPbobsCOM.Recordset)pCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Perform the SELECT statement.
                // The query result will be loaded
                // into the Recordset object
                oRecordSet.DoQuery("Select USER_CODE from OUSR");

                // Asign (link) the Recordset object
                // to the Browser.Recordset property
                oUsers.Browser.Recordset = oRecordSet;

                //find the username
                while (oUsers.Browser.EoF == false)
                {
                    if (oUsers.UserName == sUserName)
                    {
                        break;
                    }
                    oUsers.Browser.MoveNext();
                }

                //return the User Code

                return oUsers.UserCode;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);

            }

            return "";
        }



        private string GetUserName(string UserID)
        {

            Users oUsers;
            SAPbobsCOM.Recordset oRecordSet;

            try
            {
                //get users object
                oUsers = (SAPbobsCOM.Users)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

                // Get a new Recordset object
                oRecordSet = (SAPbobsCOM.Recordset)pCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Perform the SELECT statement.
                // The query result will be loaded
                // into the Recordset object
                oRecordSet.DoQuery("Select USER_CODE from OUSR");

                // Asign (link) the Recordset object
                // to the Browser.Recordset property
                oUsers.Browser.Recordset = oRecordSet;

                //find the username
                while (oUsers.Browser.EoF == false)
                {

                    if (oUsers.InternalKey == System.Convert.ToInt32(UserID))
                    {
                        break;
                    }
                    oUsers.Browser.MoveNext();
                }

                //return the User Code

                return oUsers.UserCode;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);

            }

            return "";
        }



        public string LastErrorMessage(int IretCode)
        {
            string sErr = "";
            pCompany.GetLastError(out IretCode, out sErr);
            return sErr;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            //using (var db = new SAP_Entities())
            //{

            //    var passQuery = (from a in db.OHEMs where a.passportNo == textBox5.Text && a.userId == userId && a.userId == 34 join b in db.HTM1 on a.empID equals b.empID select new { a.empID, a.firstName, a.passportNo, b.teamID }).FirstOrDefault();
            //    if (passQuery != null)
            //    {
            bool Approval = true;
            var passQuery = data.Form2Query1(userId, textBox5.Text);
            if (passQuery.Rows.Count > 0 && userId == 34)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {


                    foreach (DataGridViewRow item in dataGridView1.Rows)
                    {

                        if (item.Cells[10].Value.ToString() == "-")
                        {
                            Approval = false;
                        }
                    }

                    if (Approval)
                    {
                        var validate = true;
                        for (var i = 0; i < oInvTransDraft.Lines.Count; i++)
                        {
                            oInvTransDraft.Lines.SetCurrentLine(i);
                            if (string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORRepCd").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORCause").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORSymCd").Value.ToString()) && string.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_WORTestC").Value.ToString()) && oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value == "No")
                            {
                                validate = false;
                            }


                        }
                        if (validate)
                        {
                            oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "10";
                            LocationClass.Changer(pCompany, jobNumber, 34);

                            SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "34", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());
                            oInvTransDraft.DocumentsOwner = 34;
                        }
                        else
                        {

                            oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "5";
                            LocationClass.Changer(pCompany, jobNumber, 34, 21);

                            oInvTransDraft.DocumentsOwner = 50;

                            SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "50", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());

                        }


                        oInvTransDraft.UserFields.Fields.Item("U_POReqAp3").Value = "1";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApDa3").Value = (DateTime?)DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApTm3").Value = DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApIm3").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApUsr3").Value = userName;


                        int IretCode = oInvTransDraft.Update();
                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                        }

                        else
                        {
                            form1.gd.DataGridView4Initializer(form1.dataGridView4);
                            FieldsLoader();
                            Validator();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Define Item is Available or Not");

                    }
                }
                else
                {

                }
            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
            // }

        }


        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            cellMouseClick2 = true;

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (cellMouseClick2)
            {
                cellValueChange2 = true;
            }
        }

        private void dataGridView2_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            int IretCode;
            if ((e.ColumnIndex == 7 || e.ColumnIndex == 12 /*|| e.ColumnIndex == 13*/|| e.ColumnIndex == 15 || e.ColumnIndex == 16) && dgv.Rows[e.RowIndex].Cells[8].Value != null && cellMouseClick2 && cellValueChange2)
            {
                DialogResult result = MessageBox.Show("Update Row?", "Important Question", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)

                {

                    oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                    isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                    if (isAvailable)
                    {
                        oInvTransDraft.Lines.SetCurrentLine(System.Convert.ToInt32(dgv.Rows[e.RowIndex].Cells[8].Value));
                        if (dgv.Rows[e.RowIndex].Cells[7].Value != null)
                        {
                            //  if (oInvTransDraft.Lines.ItemCode == "_")
                            oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[7].Value.ToString();
                            oOrder = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                            oOrder.GetByKey(ORDRkey);
                            for (int i = 0; i < oOrder.Lines.Count; i++)
                            {
                                oOrder.Lines.SetCurrentLine(i);
                                if (System.Convert.ToInt32(oOrder.Lines.UserFields.Fields.Item("U_WORReqNo").Value) == ODRFkey && System.Convert.ToInt32(oOrder.Lines.UserFields.Fields.Item("U_U_RowIndex").Value) == oInvTransDraft.Lines.LineNum)
                                    oOrder.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[7].Value.ToString();
                            }
                            oOrder.Update();

                            form1.Refresh();
                        }



                        if (dgv.Rows[e.RowIndex].Cells[14].Value != null)
                        {

                            oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value = dgv.Rows[e.RowIndex].Cells[14].Value.ToString();

                        }

                        if (dgv.Rows[e.RowIndex].Cells[15].Value != null)
                        {
                            oInvTransDraft.Lines.UserFields.Fields.Item("U_OrginalItemCode").Value = oInvTransDraft.Lines.ItemCode;
                            oInvTransDraft.Lines.ItemCode = dgv.Rows[e.RowIndex].Cells[15].Value.ToString();
                            oInvTransDraft.Lines.UserFields.Fields.Item("U_Replaced").Value = "1";
                        }

                        if (dgv.Rows[e.RowIndex].Cells[16].Value != null)
                        {

                            oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value = dgv.Rows[e.RowIndex].Cells[16].Value.ToString();

                        }


                        IretCode = oInvTransDraft.Update();

                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                        }

                        else
                        {
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                gridBuilder = new Form2GridViewsInitializer(jobNumber, ODRFkey, pCompany);

                                dataGridView1.DataSource = null;
                                dataGridView1.Columns.Clear();
                                dataGridView1.Refresh();

                                dataGridView2.DataSource = null;
                                dataGridView2.Columns.Clear();
                                dataGridView2.Refresh();

                                gridBuilder.DataGridView1Initializer(dataGridView1);
                                gridBuilder.DataGridView2Initializer(dataGridView2);
                                FieldsLoader();
                                Validator();
                            }));
                        }
                    }
                    cellMouseClick2 = false;
                    cellValueChange2 = false;

                }
                else
                {

                    this.BeginInvoke(new MethodInvoker(() =>
                    {

                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Refresh();

                        gridBuilder.DataGridView2Initializer(dataGridView2);
                        FieldsLoader();
                        Validator();

                    }));
                    cellMouseClick2 = false;
                    cellValueChange2 = false;
                }
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            cellMouseClick1 = true;
        }

        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            int IretCode;
            if ((e.ColumnIndex == 6 /*|| e.ColumnIndex == 10 */|| e.ColumnIndex == 11) && dgv.Rows[e.RowIndex].Cells[7].Value != null && cellMouseClick1 && cellValueChange1)
            {
                DialogResult result = MessageBox.Show("Update Row?", "Important Question", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)

                {

                    oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                    isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                    if (isAvailable)
                    {
                        oInvTransDraft.Lines.SetCurrentLine(System.Convert.ToInt32(dgv.Rows[e.RowIndex].Cells[7].Value));
                        if (dgv.Rows[e.RowIndex].Cells[6].Value != null)
                        {
                            //if (oInvTransDraft.Lines.ItemCode == "_")

                            oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[6].Value.ToString();
                        }

                        oOrder = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        oOrder.GetByKey(ORDRkey);
                        for (int i = 0; i < oOrder.Lines.Count; i++)
                        {
                            oOrder.Lines.SetCurrentLine(i);
                            if (System.Convert.ToInt32(oOrder.Lines.UserFields.Fields.Item("U_WORReqNo").Value) == ODRFkey && System.Convert.ToInt32(oOrder.Lines.UserFields.Fields.Item("U_U_RowIndex").Value) == oInvTransDraft.Lines.LineNum)
                                oOrder.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = dgv.Rows[e.RowIndex].Cells[6].Value.ToString();
                        }
                        oOrder.Update();


                        if (dgv.Rows[e.RowIndex].Cells[10].Value != null)
                        {

                            oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value = dgv.Rows[e.RowIndex].Cells[10].Value.ToString();

                        }

                        if (dgv.Rows[e.RowIndex].Cells[11].Value != null)
                        {
                            oInvTransDraft.Lines.UserFields.Fields.Item("U_OrginalItemCode").Value = oInvTransDraft.Lines.ItemCode;
                            oInvTransDraft.Lines.ItemCode = dgv.Rows[e.RowIndex].Cells[11].Value.ToString();
                            oInvTransDraft.Lines.UserFields.Fields.Item("U_Replaced").Value = "1";
                            var desc = data.ItemsByItemCode(dgv.Rows[e.RowIndex].Cells[11].Value.ToString()).Rows[0]["FrgnName"].ToString();
                            oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = String.IsNullOrEmpty(desc) ? "" : desc;

                        }
                        IretCode = oInvTransDraft.Update();

                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                        }

                        else
                        {
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                gridBuilder = new Form2GridViewsInitializer(jobNumber, ODRFkey, pCompany);

                                dataGridView1.DataSource = null;
                                dataGridView1.Columns.Clear();
                                dataGridView1.Refresh();

                                dataGridView2.DataSource = null;
                                dataGridView2.Columns.Clear();
                                dataGridView2.Refresh();

                                gridBuilder.DataGridView1Initializer(dataGridView1);
                                gridBuilder.DataGridView2Initializer(dataGridView2);
                                FieldsLoader();
                                Validator();
                            }));
                        }
                    }
                    cellMouseClick1 = false;
                    cellValueChange1 = false;

                }
                else
                {


                    this.BeginInvoke(new MethodInvoker(() =>
                    {

                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();

                        gridBuilder.DataGridView1Initializer(dataGridView1);
                        FieldsLoader();
                        Validator();

                    }));
                    cellMouseClick1 = false;
                    cellValueChange1 = false;

                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (cellMouseClick1)
            {
                cellValueChange1 = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            //using (var db = new SAP_Entities())
            //{
            bool Approval = true;
            //var passQuery = (from a in db.OHEMs where a.passportNo == textBox13.Text && a.userId == userId join b in db.HTM1 on a.empID equals b.empID select new { a.empID, a.firstName, a.passportNo, b.teamID }).FirstOrDefault();
            //if (passQuery != null)
            //{
            var passQuery = data.Form2Query1(userId, textBox13.Text);
            if (passQuery.Rows.Count > 0 && userId == 50)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {
                    foreach (DataGridViewRow item in dataGridView2.Rows)
                    {
                        if (item.Cells[10].Value.ToString() == "")
                        {
                            Approval = false;
                        }
                    }


                    if (Approval)
                    {
                        oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "6";
                        oInvTransDraft.UserFields.Fields.Item("U_POReqAp4").Value = "1";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApDa4").Value = (DateTime?)DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApTm4").Value = DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApUsr4").Value = userName;
                        oInvTransDraft.DocumentsOwner = 34;

                        //
                        int IretCode = oInvTransDraft.Update();
                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                        }

                        else
                        {

                            form1.gd.DataGridView4Initializer(form1.dataGridView4);
                            FieldsLoader();
                            Validator();
                            LocationClass.Changer(pCompany, jobNumber, 34, 25);
                            SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "34", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());

                        }

                    }
                    else
                    {
                        MessageBox.Show("Items are Needed to Send for Purchasing First!!!");
                    }
                }


            }
            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
            //  }


        }

        private void button5_Click(object sender, EventArgs e)
        {
            DataGridView dgv = dataGridView2;
            var updated = false;
            var valid = true;
            var vendor = "";
            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            isAvailable = oInvTransDraft.GetByKey(ODRFkey);

            var Approval = true;
            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                if (item.Cells[16].Value.ToString() == "")
                {
                    Approval = false;
                }
            }

            List<KeyValuePair<int, int>> validate = new List<KeyValuePair<int, int>>();

            if (Approval)
            {
                foreach (DataGridViewRow item in dgv.Rows)
                {


                    if (System.Convert.ToInt32(item.Cells[0].Value) == 1)
                    {
                        if (item.Cells[16].Value != null)
                        {

                            validate.Add(new KeyValuePair<int, int>(item.Index, System.Convert.ToInt32(item.Cells[8].Value)));
                            updated = true;

                        }
                    }

                }
                if (updated)
                {

                    foreach (var item in validate)
                    {

                        if (!dgv.Rows[item.Key].Cells[16].Value.ToString().Equals(dgv.Rows[validate[0].Key].Cells[16].Value.ToString()) || dgv.Rows[item.Key].Cells[6].Value.ToString().Equals("_"))

                        {
                            valid = false;

                        }
                    }
                    if (valid)
                    {
                        foreach (DataGridViewRow item in dgv.Rows)
                        {
                            if (System.Convert.ToInt32(item.Cells[0].Value) == 1)
                            {
                                if (item.Cells[16].Value != null)
                                {
                                    oInvTransDraft.Lines.SetCurrentLine(System.Convert.ToInt32(item.Cells[8].Value));
                                    oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value = item.Cells[16].Value.ToString();
                                    vendor = item.Cells[16].Value.ToString();

                                }
                            }
                        }

                        var IretCode = oInvTransDraft.Update();

                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                        }

                        else
                        {
                            form3 = new Form3(this, pCompany, vendor);
                            var result = form3.ShowDialog(this);
                            var vendorcode = data.CardCodeByName(vendor);
                            var bpvalidate = true;
                            var row = data.Form1MainQuery(jobNumber).Rows[0];
                            if (result == DialogResult.No)
                            {

                                if (data.BpCodeList2(vendorcode) != null)
                                {

                                    foreach (DataGridViewRow item in dgv.Rows)
                                    {
                                        if (System.Convert.ToInt32(item.Cells[0].Value) == 1)
                                        {
                                            if (item.Cells[16].Value != null)
                                            {
                                                oInvTransDraft.Lines.SetCurrentLine(System.Convert.ToInt32(item.Cells[8].Value));
                                                if (string.IsNullOrEmpty(data.BpCodeList(oInvTransDraft.Lines.ItemCode, vendorcode)))
                                                {
                                                    bpvalidate = false;
                                                }

                                            }
                                        }
                                    }

                                    if (bpvalidate)
                                    {
                                        //oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                        //oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                        //oInvTransDraft.GetByKey(ODRFkey);

                                        ////Open a specific Form 
                                        //SBO_Application.ActivateMenuItem("2305");

                                        ////initial Form instance on the current active SAP Form
                                        //oForm = SBO_Application.Forms.GetForm("142", 0);

                                        ////intial a Edit text object with specific item
                                        //oEdit = oForm.Items.Item("4").Specific;

                                        ////Click on the a item
                                        //oForm.Items.Item("4").Click();

                                        //// set a value to the item
                                        //oInvTransDraft.Lines.SetCurrentLine(validate[0].Value);

                                        //oEdit.Value = data.CardCodeByName(oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value.ToString());

                                        //oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                                        //oOrder = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                        //oOrder.GetByKey(ORDRkey);


                                        //for (int i = 1; i <= validate.Count; i++)
                                        //{

                                        //    oInvTransDraft.Lines.SetCurrentLine(validate[i - 1].Value);

                                        //    oEdit = oMatrix.Columns.Item("2").Cells.Item(i).Specific;
                                        //    oEdit.Value = data.BpCodeList(oInvTransDraft.Lines.ItemCode, vendorcode);

                                        //    oEdit = oMatrix.Columns.Item("11").Cells.Item(i).Specific;
                                        //    oEdit.Value = oInvTransDraft.Lines.Quantity.ToString();

                                        //    oEdit = oMatrix.Columns.Item("14").Cells.Item(i).Specific;
                                        //    oEdit.Value = oInvTransDraft.Lines.UnitPrice.ToString();

                                        //    oEdit = oMatrix.Columns.Item("U_POJobNo").Cells.Item(i).Specific;
                                        //    oEdit.Value = jobNumber.ToString();

                                        //    oEdit = oMatrix.Columns.Item("U_RequisitionLineNo").Cells.Item(i).Specific;
                                        //    oEdit.Value = validate[i - 1].Value.ToString();

                                        //    oEdit = oMatrix.Columns.Item("U_POReqNo").Cells.Item(i).Specific;
                                        //    oEdit.Value = ODRFkey.ToString();

                                        //    oEdit = oMatrix.Columns.Item("U_WOITNo").Cells.Item(i).Specific;
                                        //    oEdit.Value = jobNumber.ToString();

                                        //    oEdit = oMatrix.Columns.Item("U_RequisitionNo").Cells.Item(i).Specific;
                                        //    oEdit.Value = oInvTransDraft.DocEntry.ToString();

                                        //    oEdit = oMatrix.Columns.Item("U_POProblem").Cells.Item(i).Specific;
                                        //    oEdit.Value = row["subject"].ToString();

                                        //    oEdit = oMatrix.Columns.Item("U_POSONum").Cells.Item(i).Specific;
                                        //    oEdit.Value = row["U_WOSONo"].ToString();

                                        //}

                                        //MessageBox.Show("Please Go to SAP to Complete the Process. ");

                                        //this.Close();


                                        //form here
                                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                        oInvTransDraft.GetByKey(ODRFkey);

                                        oPoDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                        oPoDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;

                                        oInvTransDraft.Lines.SetCurrentLine(validate[0].Value);

                                        oPoDraft.CardCode = data.CardCodeByName(oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value.ToString());
                                        //   oPoDraft.UserFields.Fields.Item("U_Requisit").Value = oInvTransDraft.DocEntry.ToString();

                                        oPoDraft.Lines.SupplierCatNum = data.BpCodeList(oInvTransDraft.Lines.ItemCode, vendorcode);
                                        oPoDraft.Lines.Quantity = oInvTransDraft.Lines.Quantity;
                                        oPoDraft.Lines.WarehouseCode = "A01";
                                        oPoDraft.Lines.UnitPrice = oInvTransDraft.Lines.UnitPrice;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POJobNo").Value = jobNumber.ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionLineNo").Value = validate[0].Value;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionNo").Value = oInvTransDraft.DocEntry.ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_WOITNo").Value = jobNumber;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POProblem").Value = row["subject"].ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_SOSerNo").Value = oInvTransDraft.Lines.SerialNum;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POCustNm").Value = row["custmrName"].ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POSONum").Value = row["U_WOSONo"].ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POReqNo").Value = ODRFkey.ToString();
                                        oPoDraft.Lines.AccountCode = "_SYS00000000026";

                                        for (int i = 2; i <= validate.Count; i++)
                                        {

                                            oInvTransDraft.Lines.SetCurrentLine(validate[i - 1].Value);
                                            oPoDraft.Lines.Add();
                                            oPoDraft.Lines.SupplierCatNum = data.BpCodeList(oInvTransDraft.Lines.ItemCode, vendorcode);

                                            oPoDraft.Lines.Quantity = oInvTransDraft.Lines.Quantity;
                                            oPoDraft.Lines.WarehouseCode = "A01";
                                            oPoDraft.Lines.UnitPrice = oInvTransDraft.Lines.UnitPrice;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POJobNo").Value = jobNumber.ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionLineNo").Value = validate[i - 1].Value;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionNo").Value = oInvTransDraft.DocEntry.ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_WOITNo").Value = jobNumber;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POProblem").Value = row["subject"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_SOSerNo").Value = oInvTransDraft.Lines.SerialNum;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POCustNm").Value = row["custmrName"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POSONum").Value = row["U_WOSONo"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POReqNo").Value = ODRFkey.ToString();
                                            oPoDraft.Lines.AccountCode = "_SYS00000000026";

                                        }
                                        IretCode = oPoDraft.Add();
                                        if (IretCode != 0)
                                        {
                                            MessageBox.Show(LastErrorMessage(IretCode));
                                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                                        }

                                        else
                                        {

                                            ///////////////////

                                            SAPbouiCOM.FormCreationParams oCreationParams = null;

                                            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                                            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                                            oCreationParams.UniqueID = "Form1";

                                            oForm = SBO_Application.Forms.AddEx(oCreationParams);

                                            // set the form properties
                                            oForm.Title = "Form link to draft";
                                            oForm.Left = 400;
                                            oForm.Top = 100;
                                            oForm.ClientHeight = 80;
                                            oForm.ClientWidth = 350;

                                            oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                            oItem.Left = 157;
                                            oItem.Width = 163;
                                            oItem.Top = 8;
                                            oItem.Height = 14;

                                            oItem.LinkTo = "linkb1";

                                            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
                                            //key for existing object
                                            oEdit.String = data.LastPo();

                                            oItem = oForm.Items.Add("linkb1", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                            oItem.Left = 135;
                                            oItem.Top = 8;
                                            oItem.LinkTo = "EditText1";
                                            linkedButton = ((SAPbouiCOM.LinkedButton)(oItem.Specific));
                                            //linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Drafts;
                                            linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseOrder;

                                            //Click link button icon
                                            oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked);

                                            //Make Form visible
                                            //oForm.Visible = true;

                                            oForm.Close();

                                            MessageBox.Show("Please Go to SAP to Complete the Process. ");
                                            this.Close();

                                        }

                                    }


                                    else
                                    {
                                        MessageBox.Show("Please Define Item for BP Catalog First");
                                    }

                                }
                                else
                                {
                                    //oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                    //oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                    //oInvTransDraft.GetByKey(ODRFkey);

                                    ////Open a specific Form 
                                    //SBO_Application.ActivateMenuItem("2305");

                                    ////initial Form instance on the current active SAP Form
                                    //oForm = SBO_Application.Forms.GetForm("142", 0);

                                    ////intial a Edit text object with specific item
                                    //oEdit = oForm.Items.Item("4").Specific;

                                    ////Click on the a item
                                    //oForm.Items.Item("4").Click();

                                    //// set a value to the item
                                    //oInvTransDraft.Lines.SetCurrentLine(validate[0].Value);

                                    //oEdit.Value = data.CardCodeByName(oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value.ToString());


                                    ////oItem = oForm.Items.Item("U_RequisitionNumber").Specific;

                                    ////oForm.Items.Item("U_Requisit").Specific.Value = oInvTransDraft.DocEntry.ToString();

                                    //oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                                    //oOrder = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                    //oOrder.GetByKey(ORDRkey);

                                    //for (int i = 1; i <= validate.Count; i++)
                                    //{
                                    //    oInvTransDraft.Lines.SetCurrentLine(validate[i - 1].Value);

                                    //    oEdit = oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                                    //    oEdit.Value = oInvTransDraft.Lines.ItemCode;

                                    //    oEdit = oMatrix.Columns.Item("11").Cells.Item(i).Specific;
                                    //    oEdit.Value = oInvTransDraft.Lines.Quantity.ToString();

                                    //    oEdit = oMatrix.Columns.Item("14").Cells.Item(i).Specific;
                                    //    oEdit.Value = oInvTransDraft.Lines.UnitPrice.ToString();

                                    //    oEdit = oMatrix.Columns.Item("U_POJobNo").Cells.Item(i).Specific;
                                    //    oEdit.Value = jobNumber.ToString();

                                    //    oEdit = oMatrix.Columns.Item("U_RequisitionLineNo").Cells.Item(i).Specific;
                                    //    oEdit.Value = validate[i - 1].Value.ToString();

                                    //    oEdit = oMatrix.Columns.Item("U_WOITNo").Cells.Item(i).Specific;
                                    //    oEdit.Value = jobNumber.ToString();

                                    //    oEdit = oMatrix.Columns.Item("U_POReqNo").Cells.Item(i).Specific;
                                    //    oEdit.Value = ODRFkey.ToString();

                                    //    oEdit = oMatrix.Columns.Item("U_RequisitionNo").Cells.Item(i).Specific;
                                    //    oEdit.Value = oInvTransDraft.DocEntry.ToString();

                                    //    oEdit = oMatrix.Columns.Item("U_POProblem").Cells.Item(i).Specific;
                                    //    oEdit.Value = row["subject"].ToString();

                                    //    oEdit = oMatrix.Columns.Item("U_POSONum").Cells.Item(i).Specific;
                                    //    oEdit.Value = row["U_WOSONo"].ToString();
                                    //}

                                    //MessageBox.Show("Please Go to SAP to Complete the Process. ");

                                    //this.Close();

                                    #region Create Purchase Order

                                    oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                    oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                    oInvTransDraft.GetByKey(ODRFkey);

                                    oPoDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                    oPoDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;

                                    oInvTransDraft.Lines.SetCurrentLine(validate[0].Value);
                                    // oPoDraft.UserFields.Fields.Item("U_Requisit").Value = oInvTransDraft.DocEntry.ToString();

                                    oPoDraft.CardCode = data.CardCodeByName(oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value.ToString());
                                    oPoDraft.Lines.ItemCode = oInvTransDraft.Lines.ItemCode;
                                    oPoDraft.Lines.Quantity = oInvTransDraft.Lines.Quantity;
                                    oPoDraft.Lines.WarehouseCode = "A01";
                                    oPoDraft.Lines.UnitPrice = oInvTransDraft.Lines.UnitPrice;
                                    oPoDraft.Lines.UserFields.Fields.Item("U_POJobNo").Value = jobNumber.ToString();
                                    oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionLineNo").Value = validate[0].Value;
                                    oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionNo").Value = oInvTransDraft.DocEntry.ToString();
                                    oPoDraft.Lines.UserFields.Fields.Item("U_WOITNo").Value = jobNumber;
                                    oPoDraft.Lines.UserFields.Fields.Item("U_POProblem").Value = row["subject"].ToString();
                                    oPoDraft.Lines.UserFields.Fields.Item("U_SOSerNo").Value = oInvTransDraft.Lines.SerialNum;
                                    oPoDraft.Lines.UserFields.Fields.Item("U_POCustNm").Value = row["custmrName"].ToString();
                                    oPoDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value;
                                    oPoDraft.Lines.UserFields.Fields.Item("U_POSONum").Value = row["U_WOSONo"].ToString();
                                    oPoDraft.Lines.UserFields.Fields.Item("U_POReqNo").Value = ODRFkey.ToString();
                                    oPoDraft.Lines.AccountCode = "_SYS00000000026";

                                    for (int i = 2; i <= validate.Count; i++)
                                    {

                                        oInvTransDraft.Lines.SetCurrentLine(validate[i - 1].Value);

                                        oPoDraft.Lines.Add();

                                        oPoDraft.Lines.ItemCode = oInvTransDraft.Lines.ItemCode;
                                        oPoDraft.Lines.Quantity = oInvTransDraft.Lines.Quantity;
                                        oPoDraft.Lines.WarehouseCode = "A01";
                                        oPoDraft.Lines.UnitPrice = oInvTransDraft.Lines.UnitPrice;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POJobNo").Value = jobNumber.ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionLineNo").Value = validate[i - 1].Value;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionNo").Value = oInvTransDraft.DocEntry.ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_WOITNo").Value = jobNumber;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_SOSerNo").Value = oInvTransDraft.Lines.SerialNum;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POCustNm").Value = row["custmrName"].ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value;
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POSONum").Value = row["U_WOSONo"].ToString();
                                        oPoDraft.Lines.UserFields.Fields.Item("U_POReqNo").Value = ODRFkey.ToString();
                                        oPoDraft.Lines.AccountCode = "_SYS00000000026";

                                    }
                                    IretCode = oPoDraft.Add();
                                    if (IretCode != 0)
                                    {
                                        MessageBox.Show(LastErrorMessage(IretCode));
                                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                        isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                                    }

                                    else
                                    {

                                        ///////////////////

                                        SAPbouiCOM.FormCreationParams oCreationParams = null;

                                        oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                                        oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                                        oCreationParams.UniqueID = "Form1";

                                        oForm = SBO_Application.Forms.AddEx(oCreationParams);

                                        // set the form properties
                                        oForm.Title = "Form link to draft";
                                        oForm.Left = 400;
                                        oForm.Top = 100;
                                        oForm.ClientHeight = 80;
                                        oForm.ClientWidth = 350;

                                        oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                        oItem.Left = 157;
                                        oItem.Width = 163;
                                        oItem.Top = 8;
                                        oItem.Height = 14;

                                        oItem.LinkTo = "linkb1";

                                        oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
                                        //key for existing object
                                        oEdit.String = data.LastPo();

                                        oItem = oForm.Items.Add("linkb1", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                        oItem.Left = 135;
                                        oItem.Top = 8;
                                        oItem.LinkTo = "EditText1";
                                        linkedButton = ((SAPbouiCOM.LinkedButton)(oItem.Specific));
                                        //linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Drafts;
                                        linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseOrder;

                                        //Click link button icon
                                        oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked);

                                        //Make Form visible
                                        //oForm.Visible = true;

                                        oForm.Close();

                                        MessageBox.Show("Please Go to SAP to Complete the Process. ");
                                        this.Close();


                                        #endregion
                                    }
                                }

                            }
                            else if (result == DialogResult.Yes)
                            {


                                //   var count = oPoDraft.Lines.Count;


                                if (data.BpCodeList2(vendorcode) != null)
                                {

                                    foreach (DataGridViewRow item in dgv.Rows)
                                    {
                                        if (System.Convert.ToInt32(item.Cells[0].Value) == 1)
                                        {
                                            if (item.Cells[16].Value != null)
                                            {
                                                oInvTransDraft.Lines.SetCurrentLine(System.Convert.ToInt32(item.Cells[8].Value));
                                                var test = data.BpCodeList(oInvTransDraft.Lines.ItemCode, vendorcode);

                                                if (string.IsNullOrEmpty(data.BpCodeList(oInvTransDraft.Lines.ItemCode, vendorcode)))
                                                {
                                                    bpvalidate = false;
                                                }

                                            }
                                        }
                                    }

                                    if (bpvalidate)
                                    {

                                        for (int i = 1; i <= validate.Count; i++)
                                        {

                                            oInvTransDraft.Lines.SetCurrentLine(validate[i - 1].Value);
                                            oPoDraft.Lines.Add();
                                            oPoDraft.Lines.SupplierCatNum = data.BpCodeList(oInvTransDraft.Lines.ItemCode, vendorcode);
                                            oPoDraft.Lines.Quantity = oInvTransDraft.Lines.Quantity;
                                            oPoDraft.Lines.WarehouseCode = "A01";
                                            oPoDraft.Lines.UnitPrice = oInvTransDraft.Lines.UnitPrice;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POJobNo").Value = jobNumber.ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionLineNo").Value = validate[i - 1].Value;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionNo").Value = oInvTransDraft.DocEntry.ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_WOITNo").Value = jobNumber;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POProblem").Value = row["subject"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_SOSerNo").Value = oInvTransDraft.Lines.SerialNum;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POCustNm").Value = row["custmrName"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POSONum").Value = row["U_WOSONo"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POReqNo").Value = ODRFkey.ToString();
                                            oPoDraft.Lines.AccountCode = "_SYS00000000026";

                                        }
                                        IretCode = oPoDraft.Update();
                                        if (IretCode != 0)
                                        {
                                            MessageBox.Show(LastErrorMessage(IretCode));
                                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                                        }

                                        else
                                        {

                                            ///////////////////

                                            SAPbouiCOM.FormCreationParams oCreationParams = null;

                                            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                                            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                                            oCreationParams.UniqueID = "Form1";

                                            oForm = SBO_Application.Forms.AddEx(oCreationParams);

                                            // set the form properties
                                            oForm.Title = "Form link to draft";
                                            oForm.Left = 400;
                                            oForm.Top = 100;
                                            oForm.ClientHeight = 80;
                                            oForm.ClientWidth = 350;

                                            oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                            oItem.Left = 157;
                                            oItem.Width = 163;
                                            oItem.Top = 8;
                                            oItem.Height = 14;

                                            oItem.LinkTo = "linkb1";

                                            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
                                            //key for existing object
                                            oEdit.String = existingPo.ToString();

                                            oItem = oForm.Items.Add("linkb1", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                            oItem.Left = 135;
                                            oItem.Top = 8;
                                            oItem.LinkTo = "EditText1";
                                            linkedButton = ((SAPbouiCOM.LinkedButton)(oItem.Specific));
                                            //linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Drafts;
                                            linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseOrder;

                                            //Click link button icon
                                            oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked);

                                            //Make Form visible
                                            //oForm.Visible = true;

                                            oForm.Close();

                                            MessageBox.Show("Please Go to SAP to Complete the Process. ");
                                            this.Close();

                                        }

                                    }
                                    //bp catalog is not true
                                    else
                                    {
                                        for (int i = 1; i <= validate.Count; i++)
                                        {

                                            oInvTransDraft.Lines.SetCurrentLine(validate[i - 1].Value);

                                            oPoDraft.Lines.Add();
                                            oPoDraft.Lines.ItemCode = oInvTransDraft.Lines.ItemCode;

                                            oPoDraft.Lines.Quantity = oInvTransDraft.Lines.Quantity;
                                            oPoDraft.Lines.UnitPrice = oInvTransDraft.Lines.UnitPrice;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POJobNo").Value = jobNumber.ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionLineNo").Value = validate[i - 1].Value;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_RequisitionNo").Value = oInvTransDraft.DocEntry.ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_WOITNo").Value = jobNumber;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_SOSerNo").Value = oInvTransDraft.Lines.SerialNum;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POCustNm").Value = row["custmrName"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value = oInvTransDraft.Lines.UserFields.Fields.Item("U_POSrvcDsc").Value;
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POSONum").Value = row["U_WOSONo"].ToString();
                                            oPoDraft.Lines.UserFields.Fields.Item("U_POReqNo").Value = ODRFkey.ToString();

                                        }
                                        IretCode = oPoDraft.Update();
                                        if (IretCode != 0)
                                        {
                                            MessageBox.Show(LastErrorMessage(IretCode));
                                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                                        }

                                        else
                                        {

                                            ///////////////////

                                            SAPbouiCOM.FormCreationParams oCreationParams = null;

                                            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                                            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                                            oCreationParams.UniqueID = "Form1";

                                            oForm = SBO_Application.Forms.AddEx(oCreationParams);

                                            // set the form properties
                                            oForm.Title = "Form link to draft";
                                            oForm.Left = 400;
                                            oForm.Top = 100;
                                            oForm.ClientHeight = 80;
                                            oForm.ClientWidth = 350;

                                            oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                            oItem.Left = 157;
                                            oItem.Width = 163;
                                            oItem.Top = 8;
                                            oItem.Height = 14;

                                            oItem.LinkTo = "linkb1";

                                            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
                                            //key for existing object
                                            oEdit.String = existingPo.ToString();

                                            oItem = oForm.Items.Add("linkb1", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                            oItem.Left = 135;
                                            oItem.Top = 8;
                                            oItem.LinkTo = "EditText1";
                                            linkedButton = ((SAPbouiCOM.LinkedButton)(oItem.Specific));
                                            //linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Drafts;
                                            linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseOrder;

                                            //Click link button icon
                                            oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked);

                                            //Make Form visible
                                            //oForm.Visible = true;

                                            oForm.Close();

                                            MessageBox.Show("Please Go to SAP to Complete the Process. ");
                                            this.Close();

                                        }


                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Select Same Vendors or Create Item Code for Unknown Items First.");
                    }
                }

                else
                {
                    MessageBox.Show("You Have to Select Vendor for Items.");

                }
            }


            else { MessageBox.Show("Please define vendors before creating PO"); }


        }


        private void Validator()
        {
            var oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            oInvTransDraft.GetByKey(ODRFkey);

            if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "1")
            {
                oInvTransDraft.DocumentsOwner = userId;
                oInvTransDraft.Update();
                pictureBoxSignProd.Visible = false;
                label49.Visible = false;

                button1.Visible = false;
                button1.Enabled = false;
                button9.Visible = false;
                button9.Enabled = false;
                textBox2.Enabled = false;
                textBox2.Visible = false;
                label23.Visible = false;

                button3.Visible = false;
                button3.Enabled = false;
                textBox5.Enabled = false;
                textBox5.Visible = false;
                label25.Visible = false;
                button6.Visible = false;
                button6.Enabled = false;


                button4.Visible = false;
                button4.Enabled = false;
                button8.Visible = false;
                button8.Enabled = false;
                button5.Visible = false;
                button5.Enabled = false;
                textBox13.Enabled = false;
                textBox13.Visible = false;
                label29.Visible = false;


                button11.Visible = false;
                button11.Enabled = false;
                textBox20.Enabled = false;
                textBox20.Visible = false;
                label41.Visible = false;
                button10.Visible = false;
                button10.Enabled = false;


                button7.Visible = false;
                button7.Enabled = false;
                textBox16.Enabled = false;
                textBox16.Visible = false;
                label34.Visible = false;


                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;


                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                dataGridView2.Columns[14].Visible = false;
                dataGridView2.Columns[15].Visible = false;
                label39.Text = "Draft";

            }

            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "2")
            {
                //  oInvTransDraft.DocumentsOwner = 56;

                button3.Visible = false;
                button3.Enabled = false;
                textBox5.Enabled = false;
                textBox5.Visible = false;
                label25.Visible = false;
                button6.Visible = false;
                button6.Enabled = false;


                button4.Visible = false;
                button4.Enabled = false;
                button8.Visible = false;
                button8.Enabled = false;
                button5.Visible = false;
                button5.Enabled = false;
                textBox13.Enabled = false;
                textBox13.Visible = false;
                label29.Visible = false;


                button1.Visible = false;
                button1.Enabled = false;
                button9.Visible = false;
                button9.Enabled = false;
                textBox2.Enabled = false;
                textBox2.Visible = false;
                label23.Visible = false;


                button2.Visible = false;
                button2.Enabled = false;
                textBox4.Enabled = false;
                textBox4.Visible = false;
                label24.Visible = false;


                button11.Visible = false;
                button11.Enabled = false;
                textBox20.Enabled = false;
                textBox20.Visible = false;
                label41.Visible = false;
                button10.Visible = false;
                button10.Enabled = false;

                button7.Visible = false;
                button7.Enabled = false;
                textBox16.Enabled = false;
                textBox16.Visible = false;
                label34.Visible = false;

                if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                {
                    pictureBoxSignProd.Visible = true;
                    label49.Visible = false;
                }
                else
                {
                    pictureBoxSignProd.Visible = false;
                    label49.Visible = true;
                }
                //pictureBox1.Visible = true;
                //pictureBox2.Visible = true;
                label39.Visible = true;
                label39.BringToFront();

                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;



                dataGridView1.Columns[7].Visible = false;
                //dataGridView1.Columns[10].Visible = false;
                //dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                //dataGridView2.Columns[11].Visible = false;
                //dataGridView2.Columns[13].Visible = false;
                dataGridView2.Columns[15].Visible = false;

                label39.Text = "Rejected/Void";

            }


            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "3")
            {
                if (userId == 40)
                {
                    //    oInvTransDraft.DocumentsOwner = 36;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;

                    button1.Visible = true;
                    button1.Enabled = true;
                    button9.Visible = true;
                    button9.Enabled = true;
                    button9.Visible = true;
                    button9.Enabled = true;
                    textBox2.Enabled = true;
                    textBox2.Visible = true;
                    label23.Visible = true;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;

                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }


                    dataGridView1.Enabled = true;
                    dataGridView2.Enabled = true;

                }
                else
                {
                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;

                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;

                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = false;
                    pictureBox2.Visible = false;

                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                }


                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                //dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                //dataGridView2.Columns[13].Visible = false;
                dataGridView2.Columns[15].Visible = false;
                label39.Text = "Production Pending";

            }
            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "4")
            {

                if (userId == 34)
                {
                    //    oInvTransDraft.DocumentsOwner = 50;

                    button3.Visible = true;
                    button3.Enabled = true;
                    textBox5.Enabled = true;
                    textBox5.Visible = true;
                    label25.Visible = true;
                    button6.Enabled = false;
                    button12.Enabled = true;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    pictureBox1.Visible = true;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;

                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;

                    dataGridView1.Enabled = true;
                    dataGridView2.Enabled = true;
                }

                else
                {

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;

                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;

                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                }


                dataGridView1.Columns[7].Visible = false;
                //dataGridView1.Columns[10].Visible = false;
                //dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                //dataGridView2.Columns[11].Visible = false;
                //dataGridView2.Columns[13].Visible = false;
                dataGridView2.Columns[15].Visible = false;

                label39.Text = "Inventory Pending";

            }
            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "5")
            {

                if (userId == 50)
                {
                    //  oInvTransDraft.DocumentsOwner = userId;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;

                    button4.Visible = true;
                    button4.Enabled = true;
                    button8.Visible = true;
                    button8.Enabled = true;
                    button5.Visible = true;
                    button5.Enabled = true;
                    textBox13.Enabled = true;
                    textBox13.Visible = true;
                    label29.Visible = true;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;


                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;

                    dataGridView1.Enabled = true;
                    dataGridView2.Enabled = true;

                }
                else
                {
                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;

                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;

                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;



                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;
                    pictureBox3.Visible = false;

                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;

                }

                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                // dataGridView1.Columns[11].Visible = false;

                //dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                // dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                dataGridView2.Columns[14].Visible = false;
                //dataGridView2.Columns[14].Visible = false;

                label39.Text = "Purchase Pending";

            }
            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "6")
            {
                if (userId == 34)
                {
                    // oInvTransDraft.DocumentsOwner = userId;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = true;
                    button6.Enabled = true;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;

                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value))
                    {
                        pictureBox3.Visible = true;
                    }
                    else
                    {
                        pictureBox3.Visible = false;
                    }


                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;

                }
                else
                {
                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;


                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;

                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;


                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;


                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value))
                    {
                        pictureBox3.Visible = true;
                    }
                    else
                    {
                        pictureBox3.Visible = false;
                    }


                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                }


                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                //dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                //dataGridView2.Columns[13].Visible = false;
                dataGridView2.Columns[15].Visible = false;

                label39.Text = "Good Receipt";

            }



            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "7")
            {
                if (userId == 40)
                {
                    //      oInvTransDraft.DocumentsOwner = userId;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;


                    button11.Visible = true;
                    button11.Enabled = true;
                    textBox20.Enabled = true;
                    textBox20.Visible = true;
                    label41.Visible = true;
                    button10.Visible = true;
                    button10.Enabled = true;


                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;


                    dataGridView1.Enabled = true;
                    dataGridView2.Enabled = true;

                }
                else
                {
                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;

                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;

                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;


                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;

                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                }


                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                //dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                //dataGridView2.Columns[13].Visible = false;
                dataGridView2.Columns[15].Visible = false;
                label39.Text = "Escalation";

            }


            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "8")
            {


                if (userId == data.UserId(oInvTransDraft.UserFields.Fields.Item("U_EstApUsr1").Value.ToString()))
                {
                    //   oInvTransDraft.DocumentsOwner = userId;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;

                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;


                    button7.Visible = true;
                    button7.Enabled = true;
                    textBox16.Enabled = true;
                    textBox16.Visible = true;
                    label34.Visible = true;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value))
                    {
                        pictureBox3.Visible = true;
                    }
                    else
                    {
                        pictureBox3.Visible = false;
                    }

                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                }


                else
                {
                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;

                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;


                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value))
                    {
                        pictureBox3.Visible = true;
                    }
                    else
                    {
                        pictureBox3.Visible = false;
                    }

                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;

                }

                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                dataGridView2.Columns[14].Visible = false;
                dataGridView2.Columns[15].Visible = false;


                label39.Text = "Sending Items Approval";
            }

            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "9")
            {
                // oInvTransDraft.DocumentsOwner = 56;

                button3.Visible = false;
                button3.Enabled = false;
                textBox5.Enabled = false;
                textBox5.Visible = false;
                label25.Visible = false;
                button6.Visible = false;
                button6.Enabled = false;


                button4.Visible = false;
                button4.Enabled = false;
                button8.Visible = false;
                button8.Enabled = false;
                button5.Visible = false;
                button5.Enabled = false;
                textBox13.Enabled = false;
                textBox13.Visible = false;
                label29.Visible = false;


                button1.Visible = false;
                button1.Enabled = false;
                button9.Visible = false;
                button9.Enabled = false;
                textBox2.Enabled = false;
                textBox2.Visible = false;
                label23.Visible = false;


                button2.Visible = false;
                button2.Enabled = false;
                textBox4.Enabled = false;
                textBox4.Visible = false;
                label24.Visible = false;

                button11.Visible = false;
                button11.Enabled = false;
                textBox20.Enabled = false;
                textBox20.Visible = false;
                label41.Visible = false;
                button10.Visible = false;
                button10.Enabled = false;


                button7.Visible = false;
                button7.Enabled = false;
                textBox16.Enabled = false;
                textBox16.Visible = false;
                label34.Visible = false;

                if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                {
                    pictureBoxSignProd.Visible = true;
                    label49.Visible = false;
                }
                else
                {
                    pictureBoxSignProd.Visible = false;
                    label49.Visible = true;
                }
                pictureBox1.Visible = true;
                pictureBox2.Visible = true;

                if (oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value.ToString() != "")
                {
                    pictureBox3.Visible = true;
                }
                else
                {
                    pictureBox3.Visible = false;
                }
                pictureBox4.Visible = true;

                if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm5").Value))
                {
                    pictureBox5.Visible = true;
                    label52.Visible = false;
                }
                else
                {
                    pictureBox5.Visible = false;
                    label52.Visible = true;
                }



                label39.Visible = true;
                label39.BringToFront();
                label39.Text = "Completed";

                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;



                dataGridView1.Columns[7].Visible = false;
                //dataGridView1.Columns[10].Visible = false;
                //dataGridView1.Columns[11].Visible = false;

                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                //dataGridView2.Columns[11].Visible = false;
                //dataGridView2.Columns[13].Visible = false;
                dataGridView2.Columns[15].Visible = false;


            }

            else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "10")
            {
                if (userId == 34)
                {
                    // oInvTransDraft.DocumentsOwner = userId;

                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = true;
                    button6.Enabled = true;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;

                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;


                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value))
                    {
                        pictureBox3.Visible = true;
                    }
                    else
                    {
                        pictureBox3.Visible = false;
                    }
                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm5").Value))
                    {
                        pictureBox5.Visible = true;
                    }
                    else
                    {
                        pictureBox5.Visible = false;
                    }


                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                }
                else
                {
                    button3.Visible = false;
                    button3.Enabled = false;
                    textBox5.Enabled = false;
                    textBox5.Visible = false;
                    label25.Visible = false;
                    button6.Visible = false;
                    button6.Enabled = false;

                    button1.Visible = false;
                    button1.Enabled = false;
                    button9.Visible = false;
                    button9.Enabled = false;
                    textBox2.Enabled = false;
                    textBox2.Visible = false;
                    label23.Visible = false;

                    button2.Visible = false;
                    button2.Enabled = false;
                    textBox4.Enabled = false;
                    textBox4.Visible = false;
                    label24.Visible = false;


                    button4.Visible = false;
                    button4.Enabled = false;
                    button8.Visible = false;
                    button8.Enabled = false;
                    button5.Visible = false;
                    button5.Enabled = false;
                    textBox13.Enabled = false;
                    textBox13.Visible = false;
                    label29.Visible = false;

                    button11.Visible = false;
                    button11.Enabled = false;
                    textBox20.Enabled = false;
                    textBox20.Visible = false;
                    label41.Visible = false;
                    button10.Visible = false;
                    button10.Enabled = false;


                    button7.Visible = false;
                    button7.Enabled = false;
                    textBox16.Enabled = false;
                    textBox16.Visible = false;
                    label34.Visible = false;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                    {
                        pictureBoxSignProd.Visible = true;
                        label49.Visible = false;
                    }
                    else
                    {
                        pictureBoxSignProd.Visible = false;
                        label49.Visible = true;
                    }
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value))
                    {
                        pictureBox3.Visible = true;
                    }
                    else
                    {
                        pictureBox3.Visible = false;
                    }

                    if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm5").Value))
                    {
                        pictureBox5.Visible = true;
                    }
                    else
                    {
                        pictureBox5.Visible = false;
                    }

                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;


                }

                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                // dataGridView1.Columns[11].Visible = false;

                //dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                //  dataGridView2.Columns[13].Visible = false;
                //dataGridView2.Columns[14].Visible = false;

                label39.Text = "Issue Part";

            }

            else
            {
                button1.Visible = false;
                button1.Enabled = false;
                button9.Visible = false;
                button9.Enabled = false;
                textBox2.Enabled = false;
                textBox2.Visible = false;
                label23.Visible = false;

                button3.Visible = false;
                button3.Enabled = false;
                textBox5.Enabled = false;
                textBox5.Visible = false;
                label25.Visible = false;
                button6.Visible = false;
                button6.Enabled = false;


                button2.Visible = false;
                button2.Enabled = false;
                textBox4.Enabled = false;
                textBox4.Visible = false;
                label24.Visible = false;


                button4.Visible = false;
                button4.Enabled = false;
                button8.Visible = false;
                button8.Enabled = false;
                button5.Visible = false;
                button5.Enabled = false;
                textBox13.Enabled = false;
                textBox13.Visible = false;
                label29.Visible = false;

                button11.Visible = false;
                button11.Enabled = false;
                textBox20.Enabled = false;
                textBox20.Visible = false;
                label41.Visible = false;
                button10.Visible = false;
                button10.Enabled = false;


                button7.Visible = false;
                button7.Enabled = false;
                textBox16.Enabled = false;
                textBox16.Visible = false;
                label34.Visible = false;

                if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
                {
                    pictureBoxSignProd.Visible = true;
                    label49.Visible = false;
                }
                else
                {
                    pictureBoxSignProd.Visible = false;
                    label49.Visible = true;
                }
                pictureBox1.Visible = true;
                pictureBox2.Visible = true;
                //  pictureBox3.Visible = true;
                pictureBox4.Visible = true;
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {


            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            isAvailable = oInvTransDraft.GetByKey(ODRFkey);

            //oInvTransDraft.GetByKey(approval_id);
            if (isAvailable)
            {

                oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "8";
                oInvTransDraft.DocumentsOwner = data.UserId(oInvTransDraft.UserFields.Fields.Item("U_EstApUsr1").Value.ToString());


                oInvTransDraft.Update();
                //
                int IretCode = oInvTransDraft.Update();
                if (IretCode != 0)
                {
                    MessageBox.Show(LastErrorMessage(IretCode));
                    oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                    isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                }

                else
                {
                    LocationClass.Changer(pCompany, jobNumber, 40, 0, true);

                    SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "40", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());


                    SAPbouiCOM.FormCreationParams oCreationParams = null;

                    oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                    oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                    oCreationParams.UniqueID = "Form1";

                    oForm = SBO_Application.Forms.AddEx(oCreationParams);

                    // set the form properties
                    oForm.Title = "Form link to draft";
                    oForm.Left = 400;
                    oForm.Top = 100;
                    oForm.ClientHeight = 80;
                    oForm.ClientWidth = 350;

                    oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oItem.Left = 157;
                    oItem.Width = 163;
                    oItem.Top = 8;
                    oItem.Height = 14;

                    oItem.LinkTo = "linkb1";

                    oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));

                    oEdit.String = ODRFkey.ToString();

                    oItem = oForm.Items.Add("linkb1", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oItem.Left = 135;
                    oItem.Top = 8;
                    oItem.LinkTo = "EditText1";
                    linkedButton = ((SAPbouiCOM.LinkedButton)(oItem.Specific));
                    linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Drafts;
                    //linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseOrder;

                    //Click link button icon
                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked);

                    //Make Form visible
                    oForm.Visible = true;


                    oForm.Close();


                    MessageBox.Show("Please Go to SAP to Complete the Process. ");
                    this.Close();
                }
            }
        }

        public void FieldsLoader()
        {
            var row = data.Form1MainQuery(jobNumber).Rows[0];
            var oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            oInvTransDraft.GetByKey(ODRFkey);
            label5.Text = oInvTransDraft.DocEntry.ToString();
            label4.Text = oInvTransDraft.CreationDate.ToLongDateString();
            label51.Text = oInvTransDraft.DocumentsOwner.ToString();
            label54.Text = row["custmrName"].ToString();
            label55.Text = row["U_WOPONo"].ToString();
            label56.Text = row["U_WOSONo"].ToString();
            label57.Text = jobNumber.ToString();

            textBox1.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApDa1").Value.ToString("MMM ddd d yyyy");
            textBox3.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr1").Value.ToString();
            textBox6.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApTm1").Value.ToString("HH : mm");
            if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value))
            {
                pictureBoxSignProd.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm1").Value;

            }
            else
            {
                label49.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr1").Value.ToString();

            }


            textBox8.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr2").Value.ToString();
            textBox7.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApDa2").Value.ToString("MMM ddd d yyyy");
            pictureBox1.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm2").Value;
            textBox9.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApTm2").Value.ToString("HH : mm");



            textBox11.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApDa3").Value.ToString("MMM ddd d yyyy");
            textBox10.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr3").Value.ToString();
            pictureBox2.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm3").Value;
            textBox12.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApTm3").Value.ToString("HH : mm");


            textBox29.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApDa4").Value.ToString("MMM ddd d yyyy");
            textBox28.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr4").Value.ToString();
            pictureBox3.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value;
            textBox14.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApTm4").Value.ToString("HH : mm");

            textBox22.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApDa5").Value.ToString("MMM ddd d yyyy");
            textBox21.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr5").Value.ToString();
            pictureBox5.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm5").Value;
            textBox19.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApTm5").Value.ToString("HH : mm");

            textBox33.Text = oInvTransDraft.UserFields.Fields.Item("U_comment").Value;



            textBox17.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApDa6").Value.ToString("MMM ddd d yyyy");
            textBox18.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr6").Value.ToString();
            pictureBox4.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm6").Value;
            textBox15.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApTm6").Value.ToString("HH : mm");

            if (File.Exists(imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm6").Value))
            {
                pictureBox4.ImageLocation = imageLocation + oInvTransDraft.UserFields.Fields.Item("U_EstApIm6").Value;

            }
            else
            {
                label52.Text = oInvTransDraft.UserFields.Fields.Item("U_EstApUsr6").Value.ToString();

            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            var passQuery = data.Form2Query1(userId, textBox16.Text);
            if (passQuery.Rows.Count > 0 && userId == data.UserId(oInvTransDraft.UserFields.Fields.Item("U_EstApUsr1").Value.ToString()))
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {

                    oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "9";
                    oInvTransDraft.UserFields.Fields.Item("U_POReqAp6").Value = "1";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApDa6").Value = (DateTime?)DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApTm6").Value = DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApIm6").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApUsr6").Value = userName;

                    oInvTransDraft.DocumentsOwner = 56;

                    oInvTransDraft.Update();
                    //
                    int IretCode = oInvTransDraft.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                        isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                    }

                    else
                    {
                        form1.gd.DataGridView4Initializer(form1.dataGridView4);
                        FieldsLoader();
                        Validator();
                        LocationClass.Changer(pCompany, jobNumber, data.UserId(userName));
                        SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", data.UserId(userName).ToString(), "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());
                    }
                }
                else
                {


                }
            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var passQuery = data.Form2Query1(userId, textBox13.Text);
            if (passQuery.Rows.Count > 0 && userId == 50)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {
                    var validation = false;

                    foreach (DataGridViewRow item in dataGridView2.Rows)
                    {

                        if (string.IsNullOrEmpty(item.Cells[16].Value.ToString()))
                        {

                            validation = true;
                        }
                    }
                    if (validation)
                    {
                        oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "7";
                        oInvTransDraft.UserFields.Fields.Item("U_POReqAp4").Value = "1";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApDa4").Value = (DateTime?)DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApTm4").Value = DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApIm4").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApUsr4").Value = userName;
                        oInvTransDraft.UserFields.Fields.Item("U_comment").Value = textBox33.Text;
                        oInvTransDraft.DocumentsOwner = 40;

                        oInvTransDraft.Update();
                        //
                        int IretCode = oInvTransDraft.Update();
                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                        }

                        else
                        {
                            form1.gd.DataGridView4Initializer(form1.dataGridView4);
                            FieldsLoader();
                            Validator();
                            LocationClass.Changer(pCompany, jobNumber, 40);


                            SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "40", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());

                            ///////

                            //SAPbouiCOM.FormCreationParams oCreationParams = null;

                            //oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                            //oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                            //oCreationParams.UniqueID = "Form1";

                            //oForm = SBO_Application.Forms.AddEx(oCreationParams);

                            //// set the form properties
                            //oForm.Title = "Form link to draft";
                            //oForm.Left = 400;
                            //oForm.Top = 100;
                            //oForm.ClientHeight = 80;
                            //oForm.ClientWidth = 350;

                            //oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //oItem.Left = 157;
                            //oItem.Width = 163;
                            //oItem.Top = 8;
                            //oItem.Height = 14;

                            //oItem.LinkTo = "linkb1";

                            //oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));

                            //oEdit.String = ODRFkey.ToString();

                            //oItem = oForm.Items.Add("linkb1", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            //oItem.Left = 135;
                            //oItem.Top = 8;
                            //oItem.LinkTo = "EditText1";
                            //linkedButton = ((SAPbouiCOM.LinkedButton)(oItem.Specific));
                            //linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Drafts;
                            ////linkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseOrder;

                            ////Click link button icon
                            //oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked);

                            ////Make Form visible
                            //oForm.Visible = true;


                            //oForm.Close();


                            //MessageBox.Show("Please Go to SAP to Complete the Process. ");
                            //this.Close();


                        }
                    }
                    else
                    {
                        MessageBox.Show("If Vendors are Available you can not Esclate Requisition");
                    }
                }
                else
                {


                }
            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            var passQuery = data.Form2Query1(userId, textBox2.Text);

            if (passQuery.Rows.Count > 0 && userId == 40)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {
                    oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "1";
                    oInvTransDraft.UserFields.Fields.Item("U_POReqAp2").Value = "1";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApDa2").Value = (DateTime?)DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApTm2").Value = DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApIm2").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApUsr2").Value = userName;
                    oInvTransDraft.DocumentsOwner = 56;


                    oInvTransDraft.Update();
                    //
                    int IretCode = oInvTransDraft.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                        isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                    }

                    else
                    {

                        form1.gd.DataGridView4Initializer(form1.dataGridView4);
                        FieldsLoader();
                        Validator();

                        LocationClass.Changer(pCompany, jobNumber, data.UserId(userName));

                        SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", data.UserId(oInvTransDraft.UserFields.Fields.Item("U_EstApUsr1").Value.ToString()).ToString(), "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());

                    }
                }

            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
            //}
        }

        private void button10_Click(object sender, EventArgs e)
        {
            var passQuery = data.Form2Query1(userId, textBox5.Text);
            if (passQuery.Rows.Count > 0 && userId == 34)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {

                    oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "9";
                    oInvTransDraft.UserFields.Fields.Item("U_POReqAp3").Value = "1";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApDa3").Value = (DateTime?)DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApTm3").Value = DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApIm3").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApUsr3").Value = userName;

                    oInvTransDraft.Update();
                    //
                    int IretCode = oInvTransDraft.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                        isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                    }

                    else
                    {
                        form1.gd.DataGridView4Initializer(form1.dataGridView4);
                        FieldsLoader();
                        Validator();
                        LocationClass.Changer(pCompany, jobNumber, data.UserId(userName));
                        SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", data.UserId(userName).ToString(), "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());
                    }
                }
                else
                {


                }
            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            var passQuery = data.Form2Query1(userId, textBox20.Text);

            if (passQuery.Rows.Count > 0 && userId == 40)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {
                    oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "2";
                    oInvTransDraft.UserFields.Fields.Item("U_POReqAp5").Value = "1";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApDa5").Value = (DateTime?)DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApTm5").Value = DateTime.Now;
                    oInvTransDraft.UserFields.Fields.Item("U_EstApIm5").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                    oInvTransDraft.UserFields.Fields.Item("U_EstApUsr5").Value = userName;

                    oInvTransDraft.DocumentsOwner = 56;


                    oInvTransDraft.Update();
                    //
                    int IretCode = oInvTransDraft.Update();
                    if (IretCode != 0)
                    {
                        MessageBox.Show(LastErrorMessage(IretCode));
                        oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                        isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                    }

                    else
                    {


                        form1.gd.DataGridView4Initializer(form1.dataGridView4);
                        FieldsLoader();
                        Validator();

                        LocationClass.Changer(pCompany, jobNumber, data.UserId(userName));
                        SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", data.UserId(userName).ToString(), "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());

                    }
                }

            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            var passQuery = data.Form2Query1(userId, textBox20.Text);

            if (passQuery.Rows.Count > 0 && userId == 40)
            {
                oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                isAvailable = oInvTransDraft.GetByKey(ODRFkey);

                //oInvTransDraft.GetByKey(approval_id);
                if (isAvailable)
                {
                    var validate = true;
                    for (int i = 0; i < oInvTransDraft.Lines.Count; i++)
                    {
                        oInvTransDraft.Lines.SetCurrentLine(i);

                        oInvTransDraft.Lines.SetCurrentLine(i);
                        if (oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value.ToString().Equals("No") && String.IsNullOrEmpty(oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemVendor").Value.ToString()))
                        {
                            validate = false;
                        }
                    }
                    if (validate)
                    {
                        oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "10";
                        for (int i = 0; i < oInvTransDraft.Lines.Count; i++)
                        {
                            oInvTransDraft.Lines.SetCurrentLine(i);
                            if (oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value.ToString().Equals("No"))
                            {
                                oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value = "5";
                            }

                        }

                        if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value.ToString().Equals("10"))
                        {
                            LocationClass.Changer(pCompany, jobNumber, 34);
                            SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "34", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());
                            oInvTransDraft.DocumentsOwner = 34;
                        }
                        else if (oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value.ToString().Equals("5"))
                        {
                            LocationClass.Changer(pCompany, jobNumber, 50, 21);
                            SendMessage("New Requisition Form has been Sent", "New Requisition form has been sent and needs your approval", "50", "Requisition", oInvTransDraft.DocEntry.ToString(), "112", oInvTransDraft.DocEntry.ToString(), jobNumber.ToString());
                            oInvTransDraft.DocumentsOwner = 50;
                        }

                        oInvTransDraft.UserFields.Fields.Item("U_POReqAp5").Value = "1";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApDa5").Value = (DateTime?)DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApTm5").Value = DateTime.Now;
                        oInvTransDraft.UserFields.Fields.Item("U_EstApIm5").Value = passQuery.Rows[0]["empID"].ToString() + ".png";
                        oInvTransDraft.UserFields.Fields.Item("U_EstApUsr5").Value = userName;




                        oInvTransDraft.Update();
                        //
                        int IretCode = oInvTransDraft.Update();
                        if (IretCode != 0)
                        {
                            MessageBox.Show(LastErrorMessage(IretCode));
                            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                            isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                        }

                        else
                        {



                            form1.gd.DataGridView4Initializer(form1.dataGridView4);
                            FieldsLoader();
                            Validator();


                        }
                    }
                    else { MessageBox.Show("Please Eliminate Items Without Vendor"); }
                }

            }

            else
            {
                MessageBox.Show("Your Password Is Not Valid!!!");
            }
        }

        private void dataGridView2_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            //oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            //oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            //isAvailable = oInvTransDraft.GetByKey(ODRFkey);
            // && oInvTransDraft.UserFields.Fields.Item("U_PartReqStatus").Value == "7"
            if (e.Button == MouseButtons.Right)
            {
                DataGridView datagrid = sender as DataGridView;
                if (datagrid.Rows[e.RowIndex].Cells[8].Value != null)
                    delete_id = Convert.ToInt32(datagrid.Rows[e.RowIndex].Cells[8].Value.ToString());
                this.contextMenuStrip1.Show(datagrid, e.Location);
                contextMenuStrip1.Show(Cursor.Position);

            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            if (oInvTransDraft.GetByKey(ODRFkey))
            {
                oInvTransDraft.Lines.SetCurrentLine(delete_id);
                int IretCode = 0;


                oInvTransDraft.Lines.Delete();
                IretCode = oInvTransDraft.Update();


                if (IretCode != 0)
                {
                    MessageBox.Show(LastErrorMessage(IretCode));
                    oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                    oInvTransDraft.GetByKey(ODRFkey);
                }
                else
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.Refresh();

                    dataGridView2.DataSource = null;
                    dataGridView2.Columns.Clear();
                    dataGridView2.Refresh();


                    gridBuilder = new Form2GridViewsInitializer(jobNumber, ODRFkey, pCompany);
                    gridBuilder.DataGridView1Initializer(dataGridView1);
                    gridBuilder.DataGridView2Initializer(dataGridView2);

                }

            }
        }



        private void button12_Click(object sender, EventArgs e)
        {
            DataGridView dgv1 = dataGridView1;
            DataGridView dgv2 = dataGridView2;
            int IretCode;


            oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
            isAvailable = oInvTransDraft.GetByKey(ODRFkey);

            if (isAvailable)
            {
                for (int i = 0; i < dgv1.RowCount; i++)
                {
                    oInvTransDraft.Lines.SetCurrentLine(System.Convert.ToInt32(dgv1.Rows[i].Cells[7].Value));


                    if (dgv1.Rows[i].Cells[10].Value != null)
                    {

                        oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value = dgv1.Rows[i].Cells[10].Value.ToString();

                    }
                }



                for (int i = 0; i < dgv2.RowCount; i++)
                {
                    oInvTransDraft.Lines.SetCurrentLine(System.Convert.ToInt32(dgv2.Rows[i].Cells[8].Value));


                    if (dgv2.Rows[i].Cells[14].Value != null)
                    {

                        oInvTransDraft.Lines.UserFields.Fields.Item("U_ItemAvailable").Value = dgv2.Rows[i].Cells[14].Value.ToString();

                    }
                }


                IretCode = oInvTransDraft.Update();

                if (IretCode != 0)
                {
                    MessageBox.Show(LastErrorMessage(IretCode));
                    oInvTransDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    oInvTransDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                    isAvailable = oInvTransDraft.GetByKey(ODRFkey);
                }

                else
                {
                    this.BeginInvoke(new MethodInvoker(() =>
                    {
                        gridBuilder = new Form2GridViewsInitializer(jobNumber, ODRFkey, pCompany);

                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();

                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Refresh();

                        gridBuilder.DataGridView1Initializer(dataGridView1);
                        gridBuilder.DataGridView2Initializer(dataGridView2);
                        FieldsLoader();
                        Validator();
                    }));
                }

            }

        }

        private void button13_Click(object sender, EventArgs e)
        {

            if (form1 != null)
                form1.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {

            CaptureScreen();
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
            //if (printPreviewDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    //printDocument1.Print();
            //printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);

            //}

        }
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern long BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);
        private void CaptureScreen()
        {
            Graphics myGraphics = this.CreateGraphics();
            Size s = this.Size;
            memoryImage = new Bitmap(s.Width, s.Height, myGraphics);
            Graphics memoryGraphics = Graphics.FromImage(memoryImage);
            memoryGraphics.CopyFromScreen(this.Location.X, this.Location.Y, 0, 0, s);
            IntPtr dc1 = myGraphics.GetHdc();
            IntPtr dc2 = memoryGraphics.GetHdc();
            BitBlt(dc2, 0, 0, this.ClientRectangle.Width,
            this.ClientRectangle.Height, dc1, 0, 0,
            13369376);
            myGraphics.ReleaseHdc(dc1);
            memoryGraphics.ReleaseHdc(dc2);
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(memoryImage, 0, 0);
        }


    }
}
