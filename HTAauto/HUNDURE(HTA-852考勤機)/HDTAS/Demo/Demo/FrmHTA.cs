using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Demo
{
    public partial class FrmHTA : Form
    {
        //存放打开socket的句柄


        public int htaHandle;
        //the HTA IP or COM
        public string address = "172.29.10.164";
        //the HTA Rate/Port
        public int port;
        //The DeviceID
        public int htaId = 1;
        //the foregone polling Record items
        int iprevLen = 0;

        //main user interface handle
        FrmMain mainUI;

        public FrmHTA(FrmMain mainUI)
        {
            InitializeComponent();
            this.mainUI = mainUI;
        }

        /// <summary>
        /// 将byte数组转换成ASCII字符串,比如说byte[0]=48,输出的就是"0"
        /// </summary>
        /// <param name="ReturnBytes">要转换的byte数组</param>
        /// <returns></returns>
        private string bytesToASCIIString(byte[] ReturnBytes)
        {
            string tmpReturn = "";
            foreach (byte btmp in ReturnBytes)
            {
                tmpReturn += Convert.ToChar(btmp);
            }
            return tmpReturn;
        }

        private void FrmHTA_FormClosing(object sender, FormClosingEventArgs e)
        {
            mainUI.UnconnectHTA830(address, port, htaId);
        }

        /// <summary>
        /// add normal card NO
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                string cardNO = tbCardNO.Text.Trim();
                int cardLen = cardNO.Length;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNAddHTACard(htaHandle, htaId, cardNO, cardLen, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Add Normal Card:" + cardNO + " OK");
                }
                else
                {
                    MessageBox.Show("Add Normal Card:" + cardNO + "error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// delete normal card NO.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                string cardNO = tbCardNO.Text.Trim();
                int cardLen = cardNO.Length;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNDelHTACard(htaHandle, htaId, cardNO, cardLen, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Delete Normal Card:" + cardNO + " OK");
                }
                else
                {
                    MessageBox.Show("Delete Normal Card:" + cardNO + "error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// add compress card NO.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                string cardNO = tbCpCardNO.Text.Trim();
                int cardLen = cardNO.Length;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNAddHTAZCard(htaHandle, htaId, cardNO, cardLen, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Add Compressed Card:" + cardNO + " OK");
                }
                else
                {
                    MessageBox.Show("Add Compressed Card:" + cardNO + "error!return:" + resultCode.ToString());
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        /// <summary>
        /// delete compress card NO.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                string cardNO = tbCpCardNO.Text.Trim();
                int cardLen = cardNO.Length;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNDelHTAZCard(htaHandle, htaId, cardNO, cardLen, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Delete Compressed Card:" + cardNO + " OK");
                }
                else
                {
                    MessageBox.Show("Delete Compressed Card:" + cardNO + "error!return:" + resultCode.ToString());
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        /// <summary>
        /// Disabled Illegal Card
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                byte[] dataByte = new byte[2];
                dataByte[0] = 0;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNSetHTAMemoryData(htaHandle, htaId, dataByte, 231, 1, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Disabled legal Card: OK");
                }
                else
                {
                    MessageBox.Show("Disabled legal Card:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Enabled Illegal Card
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                byte[] dataByte = new byte[2];
                dataByte[0] = 1;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNSetHTAMemoryData(htaHandle, htaId, dataByte, 231, 1, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Enabled legal Card: OK");
                }
                else
                {
                    MessageBox.Show("Enabled legal Card:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// delete HTA all legal card
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNDelHTAAllCard(htaHandle, htaId, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Delete all Card: OK");
                }
                else
                {
                    MessageBox.Show("Delete all Card:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// delete HTA log data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 5000;
                resultCode = THTA830DLL.HUNDeleteHTAAllLog(htaHandle, htaId, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Delete all Log: OK");
                }
                else
                {
                    MessageBox.Show("Delete all Log:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// get HTA all legal card NO. data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                byte[] dataByte = new byte[800];
                int dataLength = 0;
                int bank = 3;
                int compress = 0;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNGetHTACardData(htaHandle, htaId, dataByte, ref dataLength, bank, compress, timeout);
                if (resultCode == 0)
                {
                    rtbCardData.Clear();
                    for (int i = 0; i < dataLength; i++)
                        rtbCardData.AppendText(string.Format("{0,-2}", dataByte[i].ToString("X2")) + " ");
                    MessageBox.Show("Read CardData: OK");
                }
                else
                {
                    MessageBox.Show("Read CardData:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// get all HTA all log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                int compress = 0;
                //log Module NO.0-319
                int bank = 0;
                byte[] dataByte = new byte[800];
                int dataLength = 0;
                uint timeout = 1000;
                resultCode = THTA830DLL.HUNGetHTALogData(htaHandle, htaId, dataByte, ref dataLength, bank, compress, timeout);
                if (resultCode == 0)
                {
                    rtbLogData.Clear();
                    for (int i = 0; i < dataLength; i++)
                        rtbLogData.AppendText(string.Format("{0,-2}", dataByte[i].ToString("X2")) + " ");
                    MessageBox.Show("Read LogData:OK!");
                }
                else
                {
                    MessageBox.Show("Read LogData: error! return:" + resultCode.ToString());
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }


        /// <summary>
        /// restart HTA
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRestart_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 3000;
                resultCode = THTA830DLL.HUNRestartHTA(htaHandle, htaId, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Restart HTA: OK");
                }
                else
                {
                    MessageBox.Show("Restart HTA:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// initial HTA
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInitial_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 3000;
                resultCode = THTA830DLL.HUNEraseHTAFlash(htaHandle, htaId, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Initial HTA: OK");
                }
                else
                {
                    MessageBox.Show("Initial HTA:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// set HTA compress mode
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWriteCompress_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                byte[] dataByte = new byte[280];
                //set the compress 
                dataByte[0] = 1;
                resultCode = THTA830DLL.HUNSetHTAMemoryData(htaHandle, htaId, dataByte, 230, 1, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Set Compress: OK");
                }
                else
                {
                    MessageBox.Show("Set Compress:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// set HTA uncompress mode
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWriteUnCompress_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                byte[] dataByte = new byte[280];
                //set the compress 
                dataByte[0] = 0;
                resultCode = THTA830DLL.HUNSetHTAMemoryData(htaHandle, htaId, dataByte, 230, 1, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Set Uncompress: OK");
                }
                else
                {
                    MessageBox.Show("Set Uncompress:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// get HTA version
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetVersion_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                byte[] dataByte = new byte[28];
                resultCode = THTA830DLL.HUNReadHTAVersion(htaHandle, htaId, dataByte, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Get Version: " + bytesToASCIIString(dataByte));
                }
                else
                {
                    MessageBox.Show("Get Version: error! return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// change HTA ID
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChangeID_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                byte[] dataByte = new byte[280];
                //set the new ID of the HTA 
                dataByte[0] = Convert.ToByte(tbHTAId.Text.Trim());
                resultCode = THTA830DLL.HUNSetHTAMemoryData(htaHandle, htaId, dataByte, 233, 1, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Change HTA ID:" + tbHTAId.Text.Trim() + " OK");
                }
                else
                {
                    MessageBox.Show("Change HTA ID:" + tbHTAId.Text.Trim() + " error!return:" + resultCode.ToString());
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        /// <summary>
        /// get HTA time
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetDateTime_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                byte[] dateByte = new byte[9];
                byte[] timeByte = new byte[6];
                resultCode = THTA830DLL.HUNReadHTADateTime(htaHandle, htaId, dateByte, timeByte, timeout);
                if (resultCode == 0)
                {
                    tbDateTime.Text = "DateTime: " + bytesToASCIIString(dateByte) + " " + bytesToASCIIString(timeByte);
                    MessageBox.Show("Read Time:OK!");
                }
                else
                {
                    MessageBox.Show("Read Time:Error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSetDateTime_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                string date, time;
                date = DateTime.Now.Year.ToString().PadLeft(4, '0') + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + Convert.ToInt16(DateTime.Now.DayOfWeek).ToString();
                time = DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');
                resultCode = THTA830DLL.HUNWriteHTADateTime(htaHandle, htaId, date, time, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Set DateTime:Ok");
                }
                else
                {
                    MessageBox.Show("Set DateTime: error! return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnReadCompress_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 1000;
                byte[] htaMemoryData = new byte[280];
                int dataLength = 0;
                resultCode = THTA830DLL.HUNGetHTAMemoryData(htaHandle, htaId, htaMemoryData, ref dataLength, 230, 8, timeout);
                if (resultCode == 0)
                {
                    rtbMemoryContent.Clear();
                    for (int i = 0; i < dataLength; i++)
                        rtbMemoryContent.AppendText(string.Format("{0,-2}", htaMemoryData[i].ToString("X2")) + " ");
                    MessageBox.Show("Read Compress: OK");
                }
                else
                {
                    MessageBox.Show("Read Compress:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnReadMemory_Click(object sender, EventArgs e)
        {
            try
            {
                int resultCode = -1;
                uint timeout = 2000;
                byte[] htaMemoryData = new byte[280];
                int dataLength = 0;
                int iaddr = Convert.ToInt32(tbMemoryStart.Text.Trim());
                int ilen = Convert.ToInt32(tbMemoryLength.Text.Trim());
                resultCode = THTA830DLL.HUNGetHTAMemoryData(htaHandle, htaId, htaMemoryData, ref dataLength, iaddr, ilen, timeout);
                if (resultCode == 0)
                {
                    rtbMemoryContent.Clear();
                    for (int i = 0; i < dataLength; i++)
                        rtbMemoryContent.AppendText(string.Format("{0,-2}", htaMemoryData[i].ToString("X2")) + " ");
                    MessageBox.Show("Read Compress: OK");
                }
                else
                {
                    MessageBox.Show("Read Compress:error!return:" + resultCode.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnPolling_Click(object sender, EventArgs e)
        {
            if (cbUninterrupted.Checked)
            {
                tmrPolling.Enabled = true;
            }
            else
            {
                pollingHTA();
            }
        }

        public void pollingHTA()
        {
            try
            {
                int compress = 0;
                if (cbPollCompress.Checked)
                    compress = 1;
                else
                    compress = 0;
                int resultCode = -1;
                uint timeout = 1000;
                int dataLength = 0;
                // EventFormat[] EventData = new EventFormat[255];
                byte[] EventData = new byte[255 * 60];
                string show = "";
                resultCode = THTA830DLL.HUNHTAPolling(htaHandle, htaId, iprevLen, EventData, ref dataLength, compress, timeout);
                if (resultCode == 0)
                {
                    EventFormat tmpData;
                    iprevLen = dataLength;
                    for (int i = 0; i < dataLength; i++)
                    {
                        tmpData = new EventFormat();
                        for (int j = 3; j >= 0; j--)
                            tmpData.ClassCode += Convert.ToString(EventData[i * 60 + j]);
                        for (int k = 7; k >= 4; k--)
                            tmpData.IllegalCode += Convert.ToString(EventData[i * 60 + k]);
                        for (int p = 8; p <= 27; p++)
                        {
                            if (Convert.ToChar(EventData[i * 60 + p]) == '\0')
                                break;
                            tmpData.sDateTime += Convert.ToChar(EventData[i * 60 + p]);
                        }
                        for (int q = 28; q <= 47; q++)
                        {
                            if (Convert.ToChar(EventData[i * 60 + q]) == '\0')
                                break;
                            tmpData.sCard += Convert.ToChar(EventData[i * 60 + q]);
                        }
                        for (int m = 48; m <= 57; m++)
                        {
                            if (Convert.ToChar(EventData[i * 60 + m]) == '\0')
                                break;
                            tmpData.sDeviceID += Convert.ToChar(EventData[i * 60 + m]);
                        }
                        show = "No.:" + Convert.ToString(i + 1) + " " + "Class Code: " + tmpData.ClassCode + "  " +
                            "legal Code: " + tmpData.IllegalCode + "  " +
                            "Date time: " + tmpData.sDateTime + "  " +
                            "Card NO: " + tmpData.sCard + "  " +
                            "Device ID: " + tmpData.sDeviceID;
                        rtbPollContent.AppendText(show + "\n");
                    }
                }
                //the HTA has no data
                else if (htaHandle != 0 && resultCode == 1010)
                {
                    rtbPollContent.AppendText("the HTA:" + address + " has no data!\n");
                }
                else
                {
                    rtbPollContent.AppendText("Polling HTA:" + address + " error!return:" + resultCode.ToString() + "\n");
                }
            }
            catch (Exception ex)
            {
                rtbPollContent.AppendText("System error:" + ex.Message + "\n");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            tmrPolling.Enabled = false;
        }

        private void btnClearDisPlay_Click(object sender, EventArgs e)
        {
            rtbPollContent.Clear();
        }

        private void tmrPolling_Tick(object sender, EventArgs e)
        {
            try
            {
                try
                {
                    tmrPolling.Enabled = false;
                    pollingHTA();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                tmrPolling.Enabled = true;
            }
        }

        private void FrmHTA_Load(object sender, EventArgs e)
        {
            Text = "HAT(" + address + ")";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string sFileName = saveFileDialog1.FileName;
                StreamWriter sw;
                string filename = saveFileDialog1.FileName;
                if (System.IO.File.Exists(filename))
                {
                    sw = new StreamWriter(filename, true, Encoding.Unicode);
                }
                else
                {
                    sw = new StreamWriter(filename, false, Encoding.Unicode);
                }
                sw.WriteLine("---------------------------------------------------------------");
                try
                {
                    int resultCode = -1;
                    byte[] dataByte = new byte[800];
                    int dataLength = 0;
                    int compress = 0;
                    uint timeout = 1000;
                    for (int j = 0; j < 191; j++)
                    {
                        resultCode = THTA830DLL.HUNGetHTACardData(htaHandle, htaId, dataByte, ref dataLength, j, compress, timeout);
                        if (resultCode == 0)
                        {
                            string sCardData = "";
                            for (int i = 0; i < dataLength; i++)
                                sCardData = sCardData + string.Format("{0,-2}", dataByte[i].ToString("X2")) + " ";
                            sw.WriteLine("Card NO " + Convert.ToString(j) + " :" + sCardData);
                        }
                        else
                        {
                            sw.WriteLine("Card NO " + Convert.ToString(j) + " :" + "Read CardData:error!return:" + resultCode.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    sw.WriteLine("---------------------------------------------------------------");
                    sw.Close();
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string sFileName = saveFileDialog1.FileName;
                StreamWriter sw;
                string filename = saveFileDialog1.FileName;
                if (System.IO.File.Exists(filename))
                {
                    sw = new StreamWriter(filename, true, Encoding.Unicode);
                }
                else
                {
                    sw = new StreamWriter(filename, false, Encoding.Unicode);
                }
                sw.WriteLine("---------------------------------------------------------------");
                try
                {
                    int resultCode = -1;
                    int compress = 0;
                    //log Module NO.0-319
                    byte[] dataByte = new byte[800];
                    int dataLength = 0;
                    uint timeout = 1000;
                    for (int j = 0; j < 319; j++)
                    {
                        resultCode = THTA830DLL.HUNGetHTALogData(htaHandle, htaId, dataByte, ref dataLength, j, compress, timeout);
                        if (resultCode == 0)
                        {
                            string sLogData = "";
                            for (int i = 0; i < dataLength; i++)
                                sLogData = sLogData + string.Format("{0,-2}", dataByte[i].ToString("X2")) + " ";
                            sw.WriteLine("Log NO " + Convert.ToString(j) + " :" + sLogData);
                        }
                        else
                        {
                            sw.WriteLine("Log NO " + Convert.ToString(j) + " :" + "Read LogData: error! return:" + resultCode.ToString());
                        }
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
                finally
                {
                    sw.WriteLine("---------------------------------------------------------------");
                    sw.Close();
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            int readLen = 16;
            byte[] readBuf = new byte[readLen];

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850GetFPInfo(htaHandle, readBuf, ref readLen, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
            if (resultCode == 0)
            {
                HTA860Return_TextBox.Text += "Data Buffer: ";
                for (int i = 0; i < readLen; i++)
                    HTA860Return_TextBox.Text += (readBuf[i].ToString("X2") + " ");
                HTA860Return_TextBox.Text += "\r\n";

                HTA860Return_TextBox.Text += String.Format("Version: {0}{1}.{2}{3}\r\n", Convert.ToChar(readBuf[0]), Convert.ToChar(readBuf[1]), Convert.ToChar(readBuf[2]), Convert.ToChar(readBuf[3]));

                int FP_Count = ((readBuf[5] << 8) + readBuf[4]);
                HTA860Return_TextBox.Text += String.Format("Maximum FP Count: {0}\r\n", FP_Count);

                int moduleID = ((readBuf[7] << 8) + readBuf[6]);
                HTA860Return_TextBox.Text += String.Format("Module ID: {0}\r\n", moduleID);
            }
        }

        private void SetNodeID_Btn_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            resultCode = THTA830DLL.SetGcuID(Convert.ToInt32(HTA860NodeID_TextBox.Text));
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            byte[] readFP1 = new byte[386];
            byte[] readFP2 = new byte[386];

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850QueryMasterFP(htaHandle, readFP1, readFP2, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
            if (resultCode == 0)
            {
                HTA860FP1_TextBox.Text = "";
                for (int i = 0; i < 386; i++)
                    HTA860FP1_TextBox.Text += (readFP1[i].ToString("X2") + " ");

                HTA860FP2_TextBox.Text = "";
                for (int i = 0; i < 386; i++)
                    HTA860FP2_TextBox.Text += (readFP2[i].ToString("X2") + " ");
            }
        }

        private void HTA860SetMasterFP_Btn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(HTA860FP1_TextBox.Text) ||
                String.IsNullOrEmpty(HTA860FP2_TextBox.Text))
                return;

            byte[] writeFP1 = new byte[386];
            byte[] writeFP2 = new byte[386];
            string[] wFP1 = HTA860FP1_TextBox.Text.Split(' ');
            string[] wFP2 = HTA860FP2_TextBox.Text.Split(' ');
            for (int i = 0; i < 386; i++)
            {
                writeFP1[i] = Convert.ToByte(wFP1[i], 16);
                writeFP2[i] = Convert.ToByte(wFP2[i], 16);
            }

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850UpdateMasterFP(htaHandle, writeFP1, writeFP2, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
        }

        private void button14_Click_2(object sender, EventArgs e)
        {
            int readLen = 160;
            byte[] readBuf = new byte[readLen];

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850ReadTableEx(htaHandle, 1, readBuf, ref readLen, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
            HTA860Return_TextBox.Text += String.Format("Length: {0}\r\n", readLen.ToString());

            if (resultCode == 0)
            {
                HTA860Value_TextBox.Text = "";
                for (int i = 0; i < readLen; i++)
                    HTA860Value_TextBox.Text += (readBuf[i].ToString("X2") + " ");

                for (int j = 0; j < 32; j++)
                {
                    int index = j * 5;
                    HTA860Return_TextBox.Text += String.Format("Set {0}: ", (j + 1).ToString("00"));
                    HTA860Return_TextBox.Text += String.Format("Start Time: {0}:{1} , ", readBuf[index].ToString("X2"), readBuf[index + 1].ToString("X2"));
                    HTA860Return_TextBox.Text += String.Format("End Time: {0}:{1} , ", readBuf[index + 2].ToString("X2"), readBuf[index + 3].ToString("X2"));
                    HTA860Return_TextBox.Text += String.Format("Week: {0}\r\n", readBuf[index + 4].ToString("X2"));
                }
            }
        }

        private void HTA860SetRingTime_Btn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(HTA860Value_TextBox.Text))
                return;

            int writeLen = 160;
            byte[] writeBuf = new byte[writeLen];

            string[] wValue = HTA860Value_TextBox.Text.Split(' ');
            for (int i = 0; i < writeLen; i++)
                writeBuf[i] = Convert.ToByte(wValue[i], 16);

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850WriteTableEx(htaHandle, 1, writeBuf, writeLen, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
        }

        private void HTA860GetWorkTime_Btn_Click(object sender, EventArgs e)
        {
            int readLen = 128;
            byte[] readBuf = new byte[readLen];

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850ReadTableEx(htaHandle, 2, readBuf, ref readLen, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
            HTA860Return_TextBox.Text += String.Format("Length: {0}\r\n", readLen.ToString());

            if (resultCode == 0)
            {
                HTA860Value_TextBox.Text = "";
                for (int i = 0; i < readLen; i++)
                    HTA860Value_TextBox.Text += (readBuf[i].ToString("X2") + " ");

                for (int j = 0; j < 32; j++)
                {
                    int index = j * 4;
                    HTA860Return_TextBox.Text += String.Format("Set {0}: ", (j + 1).ToString("00"));
                    HTA860Return_TextBox.Text += String.Format("Time: {0}:{1} , ", readBuf[index].ToString("X2"), readBuf[index + 1].ToString("X2"));
                    HTA860Return_TextBox.Text += String.Format("Class: {0}\r\n", readBuf[index + 2].ToString("X"));
                }
            }
        }

        private void HTA860SetWorkTime_Btn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(HTA860Value_TextBox.Text))
                return;

            int writeLen = 128;
            byte[] writeBuf = new byte[writeLen];

            string[] wValue = HTA860Value_TextBox.Text.Split(' ');
            for (int i = 0; i < writeLen; i++)
                writeBuf[i] = Convert.ToByte(wValue[i], 16);

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850WriteTableEx(htaHandle, 2, writeBuf, writeLen, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
        }

        private void HTA860GetMessage_Btn_Click(object sender, EventArgs e)
        {
            int readLen = 256;
            byte[] readBuf = new byte[readLen];

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850ReadTableEx(htaHandle, 3, readBuf, ref readLen, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
            HTA860Return_TextBox.Text += String.Format("Length: {0}\r\n", readLen.ToString());

            if (resultCode == 0)
            {
                HTA860Value_TextBox.Text = "";
                for (int i = 0; i < readLen; i++)
                    HTA860Value_TextBox.Text += (readBuf[i].ToString("X2") + " ");

                for (int j = 0; j < 16; j++)
                {
                    int index = j * 16;
                    HTA860Return_TextBox.Text += String.Format("Set {0}: ", (j + 1).ToString("00"));
                    HTA860Return_TextBox.Text += String.Format("Message: ", readBuf[index + 2].ToString("X"));

                    for (int k = 0; k < 16; k++)
                    {
                        if (readBuf[index + k] == 0x00 || readBuf[index + k] == 0xFF)
                            break;

                        if (readBuf[index + k] == 0x20)
                            HTA860Return_TextBox.Text += " ";
                        else
                            HTA860Return_TextBox.Text += Convert.ToChar(readBuf[index + k]);
                    }

                    HTA860Return_TextBox.Text += String.Format("\r\n", readBuf[index + 2].ToString("X"));
                }
            }

        }

        private void HTA860SetMessage_Btn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(HTA860Value_TextBox.Text))
                return;

            int writeLen = 256;
            byte[] writeBuf = new byte[writeLen];

            string[] wValue = HTA860Value_TextBox.Text.Split(' ');
            for (int i = 0; i < writeLen; i++)
                writeBuf[i] = Convert.ToByte(wValue[i], 16);

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850WriteTableEx(htaHandle, 3, writeBuf, writeLen, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
        }

        private void btnRemoteOpen_Click(object sender, EventArgs e)
        {

            int resultCode = -1;
            resultCode = THTA830DLL.hsHTA850RemoteOpen(htaHandle, 0, ref resultCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());

        }

        //public static extern int HTA850WriteSRAM(int hComm, int iaddress, byte[] cTableData, int iTableLen, ref int iReturnCode, uint iTimeout);
        //public static extern int HTA850ReadSRAM(int hComm, int iAddress, byte[] cTableData, ref int iTableLen, ref int iReturnCode, uint iTimeout);

        // public static extern int HTA850InsertMultiUserRecord(int hComm, int CardLen, int MsgLen, int iRecord, byte[] stRecord, ref int iReturnCode, uint iTimeOut);
        private void HTA860AddCard_Btn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(HTA860CardNO_TBox.Text))
                return;

            byte[] tmpCardBuf = new byte[34];
            // Default is 0x00 (from [0] to [15])
            for (int i = 0; i < 15; i++)
                tmpCardBuf[i] = 0x00;
            // Convert Card ID to Byte Array
            for (int i = 0; i < HTA860CardNO_TBox.Text.Length && i < 16; i++)
                tmpCardBuf[i] = Convert.ToByte(HTA860CardNO_TBox.Text[i]);

            // Default is 0x00 (from [16] to [17])
            tmpCardBuf[16] = 0x00;
            tmpCardBuf[17] = 0x00;

            // Default to be 0x20 (from [18] to [33])
            for (int i = 18; i < 34; i++)
                tmpCardBuf[i] = 0x20;
            // Convert Message String to Byte Array
            for (int i = 0; i < HTA860CardMsg_TBox.Text.Length && i < 16; i++)
                tmpCardBuf[18 + i] = Convert.ToByte(HTA860CardMsg_TBox.Text[i]);

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.hsHTA850InsertMultiUserFingerPrinter2(htaHandle, 16, 16, 1, tmpCardBuf, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
        }

        // public static extern int HTA850DeleteUserRecord(int hComm, int CardLen, byte[] cCardNo, ref int iReturnCode, uint iTimeOut);
        private void HTA860DeleteCard_Btn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(HTA860CardNO_TBox.Text))
                return;

            byte[] tmpCardBuf = new byte[16];
            // Default is 0x00 (from [0] to [15])
            for (int i = 0; i < 15; i++)
                tmpCardBuf[i] = 0x00;
            // Convert Card ID to Byte Array
            for (int i = 0; i < HTA860CardNO_TBox.Text.Length && i < 16; i++)
                tmpCardBuf[i] = Convert.ToByte(HTA860CardNO_TBox.Text[i]);

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850DeleteUserRecord(htaHandle, HTA860CardNO_TBox.Text.Length, tmpCardBuf, ref returnCode, 10000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
        }

        // public static extern int HTA850PollingData(int hComm, int iPrevRecord, byte[] stRecord, ref int iRecord, ref int iReturnCode, uint iTimeOut);
        private int _LastRecord = 0;
        private void HTA860StartPolling_Btn_Click(object sender, EventArgs e)
        {
            int tmpPrev = _LastRecord;
            //tmpPrev = 0;
            _LastRecord = 0;
            byte[] tmpBuffer = new byte[2048];

            int returnCode = 0;
            int resultCode = -1;
            resultCode = THTA830DLL.HTA850PollingData(htaHandle, tmpPrev, tmpBuffer, ref _LastRecord, ref returnCode, 10000);
            if (returnCode == 0 && _LastRecord > 0)
            {
                HTA860Return_TextBox.Text = String.Format("Record Count: {0}\r\n", _LastRecord.ToString());

                for (int i = 0; i < _LastRecord; i++)
                {
                    int tmpIndex = (i * 41);

                    StringBuilder tmpSB = new StringBuilder();
                    tmpSB.Append("Time:");
                    for (int j = 0; j < 20; j++)
                    {
                        if (tmpBuffer[tmpIndex + j] == 0x00 || tmpBuffer[tmpIndex + j] == 0xFF)
                            continue;
                        tmpSB.Append(Convert.ToChar(tmpBuffer[tmpIndex + j]));
                    }

                    tmpSB.Append(", Reader:").Append(tmpBuffer[tmpIndex + 20].ToString("X2"));
                    tmpSB.Append(", Input Type:").Append(tmpBuffer[tmpIndex + 21].ToString("X2"));
                    tmpSB.Append(", Section:").Append(tmpBuffer[tmpIndex + 22].ToString("X2"));
                    tmpSB.Append(", Class:").Append(tmpBuffer[tmpIndex + 23].ToString("X2"));
                    tmpSB.Append(", Event Code:").Append(tmpBuffer[tmpIndex + 24].ToString("X2"));

                    tmpSB.Append(", Card:");
                    for (int j = 25; j < 41; j++)
                    {
                        if (tmpBuffer[tmpIndex + j] == 0x00 || tmpBuffer[tmpIndex + j] == 0xFF)
                            continue;
                        tmpSB.Append(Convert.ToChar(tmpBuffer[tmpIndex + j]));
                    }

                    tmpSB.Append("\r\n");
                    HTA860Return_TextBox.Text += tmpSB.ToString();
                }
            }
            else
            {
                HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", resultCode.ToString());
                HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", returnCode.ToString("X4"));
                // The HTA-860 has no data
                if (resultCode == 1010)
                    HTA860Return_TextBox.Text += String.Format("The HTA-860: " + address + " has no data!\r\n");
            }
        }

        private void HTA860ClearDisplay_Btn_Click(object sender, EventArgs e)
        {
            HTA860Return_TextBox.Clear();
        }

        private void button14_Click_3(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(HTA860FP1_TextBox.Text) ||
               String.IsNullOrEmpty(HTA860FP2_TextBox.Text))
                return;

            byte[] writeFP1 = new byte[386];
            byte[] writeFP2 = new byte[386];
            string[] wFP1 = HTA860FP1_TextBox.Text.Split(' ');
            string[] wFP2 = HTA860FP2_TextBox.Text.Split(' ');
            for (int i = 0; i < 386; i++)
            {
                writeFP1[i] = Convert.ToByte(wFP1[i], 16);
                writeFP2[i] = Convert.ToByte(wFP2[i], 16);
            }

            struct_FingerPrinterFormat2 sfp = new struct_FingerPrinterFormat2();
            sfp.Card = new byte[16];
            sfp.DisplayMsg = new byte[16];
            sfp.FingerPrinter1 = new byte[386];
            sfp.FingerPrinter2 = new byte[386];
            sfp.Stay = new byte[2];


            byte[] bcardno = System.Text.Encoding.Default.GetBytes(HTA860CardNO_TBox.Text);
            string ss16 = "                ";
            string smsg = HTA860CardMsg_TBox.Text;
            if (smsg.Length < 16)
            {
                smsg = smsg + ss16.Substring(0, 16 - smsg.Length);
            }
            byte[] bmessage = System.Text.Encoding.Default.GetBytes(smsg);
            byte[] baStrufp = new byte[16 + 2 + 16 + 386 * 2];
            int resultCode = 0;
            //card no fill 0x00 when len<16
            //message fill 0x20 when len<16
            Array.Clear(baStrufp, 0, baStrufp.Length);
            if (bcardno.Length > 16)
            {
                Array.ConstrainedCopy(bcardno, 0, baStrufp, 0, 16);
            }
            else
            {
                Array.ConstrainedCopy(bcardno, 0, baStrufp, 0, bcardno.Length);
            }
            baStrufp[16] = 0;
            baStrufp[17] = 1;
            Array.ConstrainedCopy(bmessage, 0, sfp.DisplayMsg, 0, 16);
            Array.ConstrainedCopy(bcardno, 0, sfp.Card, 0, bcardno.Length);
            Array.ConstrainedCopy(writeFP1, 0, sfp.FingerPrinter1, 0, 386);
            Array.ConstrainedCopy(writeFP2, 0, sfp.FingerPrinter2, 0, 386);
            sfp.Stay[0] = 0;
            sfp.Stay[1] = 1;
            sfp.DisplayMsg[15] = 0;
            Array.ConstrainedCopy(bmessage, 0, baStrufp, 18, 16);
            Array.ConstrainedCopy(writeFP1, 0, baStrufp, 34, 386);
            Array.ConstrainedCopy(writeFP2, 0, baStrufp, 420, 386);



            int returnCode = THTA830DLL.hsHTA850InsertMultiUserFingerPrinter2(htaHandle, 16, 16, 1, getBytes(sfp), ref resultCode, 30000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", returnCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", resultCode.ToString("X4"));
        }

        private void button15_Click(object sender, EventArgs e)
        {
            byte[] fp1 = new byte[386];
            byte[] fp2 = new byte[386];
            int iCardFormatLen = 0;
            int iReturnCode = 0;
            byte[] querycard38 = new byte[38];
            int iRtn = THTA830DLL.hsHTA850QueryUserFingerPrinter2(htaHandle, HTA860CardNO_TBox.Text.Length, HTA860CardNO_TBox.Text, fp1, fp2, ref iCardFormatLen, ref iReturnCode, 30000);
            HTA860Return_TextBox.Text = String.Format("Result Code: {0}\r\n", iReturnCode.ToString());
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", iReturnCode.ToString("X4"));
            if (iReturnCode == 0)
            {
                HTA860FP1_TextBox.Text = "";
                for (int i = 0; i < 386; i++)
                    HTA860FP1_TextBox.Text += (fp1[i].ToString("X2") + " ");

                HTA860FP2_TextBox.Text = "";
                for (int i = 0; i < 386; i++)
                    HTA860FP2_TextBox.Text += (fp2[i].ToString("X2") + " ");
            }
        }

        byte[] getBytes(struct_FingerPrinterFormat2 str)
        {
            int size = Marshal.SizeOf(str);
            byte[] arr = new byte[size];
            IntPtr ptr = Marshal.AllocHGlobal(size);

            Marshal.StructureToPtr(str, ptr, true);
            Marshal.Copy(ptr, arr, 0, size);
            Marshal.FreeHGlobal(ptr);

            return arr;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            // Send:
            //  Data=Address(2 byte, Lo Hi)+DataLen(2 byte, Lo Hi)

            byte[] senddata = new byte[10];
            byte[] receivedata = new byte[512];
            int iReceiveLen = 0;
            int iReturnCode = 0;
            senddata[0] = 0;
            senddata[1] = 0;
            senddata[2] = 20;
            senddata[3] = 0;
            int iRtn = THTA830DLL.hsHTA850Set(htaHandle, 1, senddata, 0, receivedata, ref iReceiveLen, ref iReturnCode, 30000);
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", iReturnCode.ToString("X4"));
            if (iRtn == 0)
            {

                HTA860Value_TextBox.Text = "";
                for (int i = 0; i < iReceiveLen; i++)
                    HTA860Return_TextBox.Text += (receivedata[i].ToString("X2") + " ");


            }
            else
            {
                HTA860Value_TextBox.Text = "Return Code:" + iRtn.ToString();
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            byte bInitFlag = new byte();
            int iReturnCode = 0;

            int iRtn = THTA830DLL.hsHTA850Initial(htaHandle, bInitFlag, ref iReturnCode, 30000);
            HTA860Return_TextBox.Text += String.Format("Return Code: 0x{0}\r\n", iReturnCode.ToString("X4"));
            if (iRtn == 0)
            {
                HTA860Value_TextBox.Text = "Success!" ;
            }
            else
            {
                HTA860Value_TextBox.Text = "Fail! Return Code:" + iRtn.ToString();
            }
        }

    }
}