using SAP_Q100.Class;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SAP_Q100
{
    public partial class Form3 : Form
    {
        private static SAPbobsCOM.Company pCompany;
        private SAPbobsCOM.Documents oPoDraft;
        private Form2 form2;
        private SqlData data;
        private string vendor;
        public Form3(Form2 senderForm, SAPbobsCOM.Company company,string vendorName)
        {
            vendor = vendorName;
            form2 = senderForm;
            pCompany = company;
            data = new SqlData();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show(" Please Input PO Number. ");
            }
            else
            {
                oPoDraft = (SAPbobsCOM.Documents)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                oPoDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
            if (oPoDraft.GetByKey(data.CheckPoNumber(textBox1.Text)))
                {
                    if (oPoDraft.CardName == vendor )
                    {
                        form2.existingPo = data.CheckPoNumber(textBox1.Text);
                        form2.oPoDraft = oPoDraft;
                        this.DialogResult = DialogResult.Yes;
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show(" The Vendor Code Is not Matched with This Purchase Order ");
                    }
                }
                else
                {
                    MessageBox.Show(" This PO Is Not Exist Or Is Close. ");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
