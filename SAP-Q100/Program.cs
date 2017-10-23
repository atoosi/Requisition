using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SAP_Q100
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            int jobNo;
            int userId;
            int reqNo;
       
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length == 3)
            {
                jobNo = int.Parse(args[0]);
                userId = int.Parse(args[1]);
                reqNo = int.Parse(args[2]);
                Application.Run(new Form1(jobNo, userId, reqNo));
            }
            else if(args.Length == 1)
            {
                int start = args[0].IndexOf("@");
                reqNo = System.Convert.ToInt32(args[0].Substring(0, start));
                int next  = args[0].IndexOf("@",start+1);
                jobNo = System.Convert.ToInt32(args[0].Substring(start + 1, next - start - 1));
                int end = args[0].IndexOf("@",next+1);
                userId = System.Convert.ToInt32(args[0].Substring(next+1,end-next-1));
                Application.Run(new Form1(jobNo, userId, reqNo));
            }
            else
            {
                jobNo = int.Parse(args[0]);
                userId = int.Parse(args[1]);
                Application.Run(new Form1(jobNo, userId, 0));
            }

        }
    }
}
