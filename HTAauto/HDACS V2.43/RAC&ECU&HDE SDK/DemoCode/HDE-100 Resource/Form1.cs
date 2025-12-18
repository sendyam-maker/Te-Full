using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DEMO
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int result = -1;
            int timeout = 2000;
            UInt32 handle = 0;
            string address=Convert.ToString(gvwRAC2000.CurrentRow.Cells["clAddress"].Value);
            int port =Convert.ToInt32(gvwRAC2000.CurrentRow.Cells["clPort"].Value);
            resultCode = TRAC2000ELDLL.OpenChannel( ref handle,address, port);
            if (resultCode == 0)
            {
                FrmRAC2000EL frmRac2000 = new FrmRAC2000EL();
                frmRac2000.RAC2000Handle = handle;
                frmRac2000.Address = address;
                frmRac2000.Port = port;
                frmRac2000.RAC2000Id = Convert.ToInt32(gvwRAC2000.CurrentRow.Cells["clId"].Value);
                frmRac2000.ShowDialog();
                resultCode = TRAC2000ELDLL.CloseChannel(handle);
                if (resultCode != 0)
                {
                    MessageBox.Show("CloseChannel error!error code :" + Convert.ToInt32(resultCode));
                }
            }
            else
            {
                MessageBox.Show("OpenChannel error!error code :"+Convert.ToInt32(resultCode));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void gvwRAC2000_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["clAddress"].Value = "172.29.10.221";
            e.Row.Cells["clPort"].Value = "4660";
            e.Row.Cells["clId"].Value = "1";
            e.Row.Cells["clStatus"].Value = "";
            e.Row.Cells["clResult"].Value = "";
        }
    }
}