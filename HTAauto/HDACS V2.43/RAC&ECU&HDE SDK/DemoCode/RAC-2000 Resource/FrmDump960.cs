using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace DEMO
{
    public partial class FrmDump960 : Form
    {
        public UInt32 _RAC2000Handle;
        public string Address;
        public int Port;
        public int _RAC2000Id;
        Dump96 dump960 = new Dump96();
        delegate void Callback(string s);
        delegate void CallbackEvent(clPubevent cPubEvent);
        string sTxtFile = "";
        public FrmDump960(int RAC2000Id, UInt32 RAC2000Handle)
        {
            InitializeComponent();
            _RAC2000Id = RAC2000Id;
            _RAC2000Handle = RAC2000Handle;
            dump960.Callbackmsg += new Dump96.callbackmsg(dump960_callbackmsg);
            dump960.CallbackPubEvent += new Dump96.callbackpubevent(dump960_CallbackPubEvent);
            dump960.RAC2000Id = RAC2000Id;
            dump960.RAC2000Handle = RAC2000Handle;
            sTxtFile = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "dumpdata.txt";
            if (File.Exists(sTxtFile))
                File.Delete(sTxtFile);
        }

        public void WriteText(string sFileName, string sText)
        {
            StreamWriter writer = new StreamWriter(sFileName, true, Encoding.Unicode);
            writer.WriteLine(sText);
            writer.Close();

        }

        void dump960_CallbackPubEvent(clPubevent cEvent)
        {
            string Str = "Card No:" + cEvent.EventCard + ",EventTime:" + cEvent.EventDate + " " + cEvent.EventTime + ",Event Code:" + cEvent.EventCode;
            WriteText(sTxtFile, Str);
        }

        void dump960_callbackmsg(string s)
        {
            if (label1.InvokeRequired)
            {
                Callback d = new Callback(dump960_callbackmsg);
                label1.Invoke(d, new object[] { s });
            }
            else
            {
                label1.Text = s;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            byte[] b = new byte[dump960.MaxMemorySize];
            byte[] r = new byte[1024];
            dump960.BeginDump();
            int ioffset = 0;
            int ircvlen = 0;
            int inodatacount = 0;
            int iretcode = 0;
            progressBar1.Maximum = 1000;
            progressBar1.Value = 0;
            do
            {
                iretcode = dump960.KeepDump(ref b, ref r, ioffset , ref ircvlen);
                if (iretcode == 0xff)
                {
                    inodatacount++;
                }
                else
                {
                    dump960.DecodeSwip(r);
                    inodatacount = 0;
                    ioffset += ircvlen;
                }
                if (inodatacount > 8)
                    break;

                if (progressBar1.Value < progressBar1.Maximum)
                    progressBar1.Value += 1;
                label1.Text = progressBar1.Value.ToString();
                //dump960.KeepDump(ref b, ioffset + ircvlen, ref ircvlen);
                //dump960.KeepDump(ref b, ioffset + ircvlen, ref ircvlen);
            }
            while (ioffset < dump960.MaxMemorySize);
            using (var s = File.OpenText(sTxtFile))
            {
                string line;
                while ((line = s.ReadLine()) != null)
                {
                    listBox1.Items.Add(line);
                }
            }
            MessageBox.Show("Finish Job.");
        }
    }
}
