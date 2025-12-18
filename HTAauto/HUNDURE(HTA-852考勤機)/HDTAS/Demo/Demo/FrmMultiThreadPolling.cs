using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;

using System.Threading;

namespace Demo
{
    public partial class FrmMultiThreadPolling : Form
    {
        ArrayList HTAList;
        ArrayList showMsg;
        public DataGridView grvHTAList;
        public FrmMultiThreadPolling()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            HTAList.Clear();
            int compress = 0;
            if (cbCompress.Checked)
                compress = 1;
            for (int i = 0; i < grvHTAList.Rows.Count - 1; i++)
            {
                THTAPolling hta = new THTAPolling(Convert.ToString(grvHTAList.Rows[i].Cells[0].Value), Convert.ToInt32(grvHTAList.Rows[i].Cells[1].Value), Convert.ToInt32(grvHTAList.Rows[i].Cells[2].Value));
                hta.compress = compress;
                hta.Onhint = new OnHintEvent(showMessage);
                if (hta.connectHTA())
                {
                    HTAList.Add(hta);
                }
            }
            for (int i = 0; i < HTAList.Count; i++)
            {
                THTAPolling hta = (THTAPolling)HTAList[i];
                hta.start();
            }
        }

        private void FrmMultiThreadPolling_Load(object sender, EventArgs e)
        {
            HTAList =new ArrayList();
            showMsg = new ArrayList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            for (int i = 0; i < HTAList.Count; i++)
            {
                THTAPolling hta = (THTAPolling)HTAList[i];
                hta.stop();
            }
            HTAList.Clear();
        }

        public void showMessage(string hint)
        {
            try
            {
                if (rtbResult.InvokeRequired)
                {
                    OnHintEvent d = new OnHintEvent(showMessage);
                    rtbResult.Invoke(d, new object[] { hint });
                }
                else
                {
                    rtbResult.AppendText(hint + "\n");
                }
            }
            catch (Exception ex)
            { 
                //
            }
        }

        private void FrmMultiThreadPolling_FormClosed(object sender, FormClosedEventArgs e)
        {
            button2_Click(null,null);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            rtbResult.Clear();
        }

    }
}