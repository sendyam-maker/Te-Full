using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Demo
{
    public partial class FrmMain : Form
    {
        ArrayList alHTAList;
        public FrmMain()
        {
            InitializeComponent();
            alHTAList = new ArrayList();
        }

        public FrmHTA findHTA(string address,int port,int id)
        {
            foreach (Object obj in alHTAList)
            {
                FrmHTA htaForm = obj as FrmHTA;
                if ((htaForm.address == address) && (htaForm.port == port) && (htaForm.htaId == id))
                {
                    return htaForm;
                }
            }
            return null;
        }

        public void deleteHTA(string address, int port, int id)
        {
            Object objTemp=null;
            foreach (Object obj in alHTAList)
            {
                FrmHTA htaForm = obj as FrmHTA;
                if ((htaForm.address == address) && (htaForm.port == port) && (htaForm.htaId == id))
                {
                    objTemp = obj;
                }
            }
            if(objTemp!=null)
            {
                alHTAList.Remove(objTemp);
            }
        }

        public void UnconnectHTA830(string address, int port, int id)
        {
            try
            {
                FrmHTA htaForm = null;
                htaForm = findHTA(address,port,id);
                if (htaForm == null)
                {
                    MessageBox.Show("device:" + Convert.ToString(gvwHTAList.CurrentRow.Cells[0].Value) + "is unconnect status!");
                    return;
                }
                int resultCode;
                resultCode = THTA830DLL.HUNHTACloseSocket(htaForm.htaHandle);
                if (resultCode == 0)
                {
                    for (int i = 0; i < gvwHTAList.Rows.Count - 1; i++)
                    {
                        if (gvwHTAList.Rows[i].Cells[0].Value == htaForm.address && gvwHTAList.Rows[i].Cells[1].Value.ToString() ==Convert.ToString(htaForm.port) && gvwHTAList.Rows[i].Cells[2].Value.ToString() == Convert.ToString(htaForm.htaId))
                        {
                            gvwHTAList.Rows[i].Cells[3].Value = "unconnect";
                            gvwHTAList.Rows[i].Cells[4].Value = "unconnect " + htaForm.address + " OK";
                        }
                    }
                    deleteHTA(htaForm.address, htaForm.port, htaForm.htaId);
                }
                else
                {
                    gvwHTAList.CurrentRow.Cells[4].Value = "unconnect " + htaForm.address + " error!return:" + resultCode.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void connectHTAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                gvwHTAList.EndEdit();
                FrmHTA htaForm = null;
                htaForm = findHTA(Convert.ToString(gvwHTAList.CurrentRow.Cells[0].Value), Convert.ToInt32(Convert.ToString(gvwHTAList.CurrentRow.Cells[1].Value).Trim()), Convert.ToInt32(Convert.ToString(gvwHTAList.CurrentRow.Cells[2].Value).Trim()));
                if (htaForm != null)
                {
                    htaForm.Show();
                    return;
                }
                htaForm = new FrmHTA(this);
                htaForm.htaHandle = 0;
                htaForm.address = Convert.ToString(gvwHTAList.CurrentRow.Cells[0].Value);
                htaForm.port = Convert.ToInt32(Convert.ToString(gvwHTAList.CurrentRow.Cells[1].Value).Trim());
                htaForm.htaId = Convert.ToInt32(Convert.ToString(gvwHTAList.CurrentRow.Cells[2].Value).Trim());
                int resultCode;
                resultCode = THTA830DLL.HUNHTAOpenSocket(ref htaForm.htaHandle,htaForm.address, htaForm.port);
                if (resultCode == 0)
                {
                    gvwHTAList.CurrentRow.Cells[3].Value = "connecting";
                    gvwHTAList.CurrentRow.Cells[4].Value = "open " + htaForm.address + " OK";
                    alHTAList.Add(htaForm);
                    htaForm.Show();
                }
                else
                {
                    gvwHTAList.CurrentRow.Cells[4].Value = "Open " + htaForm.address + ":error! return:" + resultCode.ToString();
                    return;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void unconnectHTAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                UnconnectHTA830(Convert.ToString(gvwHTAList.CurrentRow.Cells[0].Value), Convert.ToInt32(Convert.ToString(gvwHTAList.CurrentRow.Cells[1].Value).Trim()), Convert.ToInt32(Convert.ToString(gvwHTAList.CurrentRow.Cells[2].Value).Trim()));
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void gvwHTAList_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            gvwHTAList.CurrentRow.Cells[1].Value = "4660";
            gvwHTAList.CurrentRow.Cells[2].Value = "1";
            gvwHTAList.CurrentRow.Cells[3].Value = "unconnect";
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            connectHTAToolStripMenuItem_Click(null,null);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            unconnectHTAToolStripMenuItem_Click(null,null);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            gvwHTAList.CurrentRow.Cells[0].Value = "172.16.39.4";
            gvwHTAList.CurrentRow.Cells[1].Value = "4660";
            gvwHTAList.CurrentRow.Cells[2].Value = "1";
            gvwHTAList.CurrentRow.Cells[3].Value = "unconnect";
        }

    }
}