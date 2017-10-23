using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SAP_Q100.Class
{
    static class LocationClass
    {
        public static void Changer(SAPbobsCOM.Company pCompany, int key, int userID, int statusID, bool first)
        {
            SAPbobsCOM.ServiceCalls oServiceCall = (SAPbobsCOM.ServiceCalls)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls);


            oServiceCall.GetByKey(key);
            var data = new SqlData();
            bool validation = false;
            foreach(var item in data.Requisitions(key))
            {
                if (item.ReqisitionStatus > statusID)
                {
                    validation = true;
                }
            }
            if (validation)
            {

                if (!first)
                {

                    oServiceCall.UserFields.Fields.Item("U_FirstLocation").Value = oServiceCall.Status;
                    oServiceCall.Status = statusID;
                }
                else
                {

                    oServiceCall.Status = oServiceCall.UserFields.Fields.Item("U_FirstLocation").Value;
                }
                oServiceCall.AssigneeCode = userID;
                int IretCode = oServiceCall.Update();

                if (IretCode != 0)
                {
                    string sErr = "";
                    pCompany.GetLastError(out IretCode, out sErr);
                    MessageBox.Show(sErr);
                }
            }
        }

        public static void Changer(SAPbobsCOM.Company pCompany, int key, int userID, int statusID)
        {
            SAPbobsCOM.ServiceCalls oServiceCall = (SAPbobsCOM.ServiceCalls)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls);


            oServiceCall.GetByKey(key);

            //oServiceCall.Lines.ItemCode = ".001MX1000BF";
            //oServiceCall.Lines.Quantity = 1;
            //oServiceCall.Lines.WarehouseCode = "02";
            oServiceCall.Status = statusID;
            oServiceCall.AssigneeCode = userID;


            int IretCode = oServiceCall.Update();

            if (IretCode != 0)
            {
                string sErr = "";
                pCompany.GetLastError(out IretCode, out sErr);
                MessageBox.Show(sErr);
            }
        }

        public static void Changer(SAPbobsCOM.Company pCompany, int key, int userID)
        {
            SAPbobsCOM.ServiceCalls oServiceCall = (SAPbobsCOM.ServiceCalls)pCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls);


            oServiceCall.GetByKey(key);


            oServiceCall.AssigneeCode = userID;

            int IretCode = oServiceCall.Update();

            if (IretCode != 0)
            {
                string sErr = "";
                pCompany.GetLastError(out IretCode, out sErr);
                MessageBox.Show(sErr);
            }
        }
    }
}
