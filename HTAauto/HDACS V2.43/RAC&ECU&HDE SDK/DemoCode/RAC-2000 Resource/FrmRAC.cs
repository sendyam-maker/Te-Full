using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace DEMO
{
    public partial class FrmRAC2000 : Form
    {
        public UInt32 RAC2000Handle;
        public string Address;
        public int Port;
        public int RAC2000Id;

        private int prevRecord;

        public FrmRAC2000()
        {
            InitializeComponent();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardNO.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            if (chbCompress.Checked)
            {
                resultCode = TRAC2000DLL.AddZCard(RAC2000Id, tbCardNO.Text, tbCardNO.Text.Length, tbPWD.Text, tbPWD.Text.Length, Convert.ToChar(0), RAC2000Handle, timeout);
            }
            else
            {
                resultCode = TRAC2000DLL.AddCard(RAC2000Id, tbCardNO.Text, tbCardNO.Text.Length, tbPWD.Text, tbPWD.Text.Length, 0, Convert.ToChar(0), RAC2000Handle, timeout);
            }
            if (chbCompress.Checked)
            {
                if (resultCode == 0)
                {
                    MessageBox.Show("AddZCard OK!");
                }
                else
                {
                    MessageBox.Show("AddZCard error!error code :" + Convert.ToInt32(resultCode));
                }
            }
            else
            {
                if (resultCode == 0)
                {
                    MessageBox.Show("AddCard OK!");
                }
                else
                {
                    MessageBox.Show("AddCard error!error code :" + Convert.ToInt32(resultCode));
                }
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardNO.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            resultCode = TRAC2000DLL.AddVisitorCard(RAC2000Id, tbCardNO.Text, 0x01, tbCardNO.Text.Length, "20010101", "000000", "20151231", "235959", 0xff, 2000, 0x00, 1, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("AddVisitorCard OK!");
            }
            else
            {
                MessageBox.Show("AddVisitorCard error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardNO.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            if (chbCompress.Checked)
            {
                resultCode = TRAC2000DLL.DelZCard(RAC2000Id, tbCardNO.Text, tbCardNO.Text.Length, RAC2000Handle, timeout);
            }
            else
            {
                resultCode = TRAC2000DLL.DelCard(RAC2000Id, tbCardNO.Text, tbCardNO.Text.Length, RAC2000Handle, timeout);
            }
            if (chbCompress.Checked)
            {
                if (resultCode == 0)
                {
                    MessageBox.Show("DelZCard OK!");
                }
                else
                {
                    MessageBox.Show("DelZCard error!error code :" + Convert.ToInt32(resultCode));
                }
            }
            else
            {
                if (resultCode == 0)
                {
                    MessageBox.Show("DelCard OK!");
                }
                else
                {
                    MessageBox.Show("DelCard error!error code :" + Convert.ToInt32(resultCode));
                }
            }
        }


        private void button41_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardNO.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            resultCode = TRAC2000DLL.DelVisitorCard(RAC2000Id, 0x00, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("DelVisitorCard OK!");
            }
            else
            {
                MessageBox.Show("DelVisitorCard error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 5000;
            resultCode = TRAC2000DLL.RelayAction(RAC2000Id, Convert.ToChar(15), Convert.ToChar(15), RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("RelayAction ON OK!");
            }
            else
            {
                MessageBox.Show("RelayAction ON error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 5000;
            resultCode = TRAC2000DLL.RelayAction(RAC2000Id, Convert.ToChar(0), Convert.ToChar(15), RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("RelayAction OFF OK!");
            }
            else
            {
                MessageBox.Show("RelayAction OFF error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            for (int i = 0; i < clbRelayList.Items.Count; i++)
            {
                if (clbRelayList.CheckedItems.Contains(clbRelayList.Items[i]))
                {
                    resultCode = TRAC2000DLL.RelayAction(RAC2000Id, Convert.ToChar(1 << i), Convert.ToChar(1 << i), RAC2000Handle, timeout);
                    if (resultCode == 0)
                    {
                        MessageBox.Show("RelayAction ON" + Convert.ToString(i) + "Relay OK!");
                    }
                    else
                    {
                        MessageBox.Show("RelayAction ON" + Convert.ToString(i) + "Relay error!error code :" + Convert.ToInt32(resultCode));
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            for (int i = 0; i < clbRelayList.Items.Count; i++)
            {
                if (clbRelayList.CheckedItems.Contains(clbRelayList.Items[i]))
                {
                    resultCode = TRAC2000DLL.RelayAction(RAC2000Id, Convert.ToChar(0), Convert.ToChar(1 << i), RAC2000Handle, timeout);
                    if (resultCode == 0)
                    {
                        MessageBox.Show("RelayAction  OFF" + Convert.ToString(i) + "Relay OK!\n");
                    }
                    else
                    {
                        MessageBox.Show("RelayAction OFF" + Convert.ToString(i) + "Relay error!error code :" + Convert.ToInt32(resultCode));
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            rtbDisplay.Clear();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                button5.Enabled = false;
                int resultCode = -1;
                int timeout = 10000;
                byte[] dataBuffer = new byte[16640];
                int recordLen = 0;
                int flag = chbPollCompress.Checked ? 1 : 0;
                if (chbReceive.Checked)
                    flag = chbPollCompress.Checked ? 0x11 : 0x10;

                resultCode = TRAC2000DLL.Polling(RAC2000Id, prevRecord, dataBuffer, ref recordLen, RAC2000Handle, timeout, flag);
                prevRecord = recordLen;
                if (resultCode == 0)
                {
                    SEventStruct eventRec;
                    for (int j = 0; j < recordLen; j++)
                    {
                        eventRec = (SEventStruct)TCommon.BytesToStuct(dataBuffer, typeof(SEventStruct), Marshal.SizeOf(typeof(SEventStruct)) * j);
                        rtbDisplay.AppendText("Event Code:" + TCommon.ByteArrayToString(eventRec.cEventCode) + " ; ");
                        rtbDisplay.AppendText("Date Time:" + TCommon.ByteArrayToString(eventRec.cDateTime) + " ; ");
                        rtbDisplay.AppendText("Card Number:" + TCommon.ByteArrayToString(eventRec.cCard) + " ; ");
                        rtbDisplay.AppendText("Device ID:" + TCommon.ByteArrayToString(eventRec.cDeviceID) + " ; ");
                        rtbDisplay.AppendText("Reader ID:" + TCommon.ByteArrayToString(eventRec.cReaderID));
                        rtbDisplay.AppendText("\n");
                    }
                    rtbDisplay.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Polling ok\n");
                }
                else if (resultCode == 1003)
                {
                    rtbDisplay.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Request equipment is overtime\n");
                }
                else if (resultCode == 1004)
                {
                    rtbDisplay.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! The handle of value in equipment is false!\n");
                }
                else if (resultCode == 1005)
                {
                    rtbDisplay.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Transmitting package to equipment  is error!\n");
                }
                else if (resultCode == 1006)
                {
                    rtbDisplay.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Respond package CRC to equipment error!\n");
                }
                else
                {
                    rtbDisplay.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Polling error!error code :" + Convert.ToString(resultCode) + "\n");
                }

            }
            finally
            {
                button5.Enabled = true;
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dateBuffer = new byte[10];
            byte[] timeBuffer = new byte[10];
            resultCode = TRAC2000DLL.GetDateTime(RAC2000Id, dateBuffer, timeBuffer, RAC2000Handle, timeout);
            string dateString = TCommon.ByteArrayToString(dateBuffer);
            string timeString = TCommon.ByteArrayToString(timeBuffer);
            DateTime date = Convert.ToDateTime(dateString.Substring(0, 4) + "/" + dateString.Substring(4, 2) + "/" + dateString.Substring(6, 2));
            DateTime time = Convert.ToDateTime(timeString.Substring(0, 2) + ":" + timeString.Substring(2, 2) + ":" + timeString.Substring(4, 2));

            if (resultCode == 0)
            {
                //dtpDate.Value = date;
                //dtpTime.Value = time;
                //WeekDay: sunday: 7, Monday:1, Tuesday:2 ....
                string sshowdatetime = "Date:" + dateString.Substring(0, 8) + ",WeekDay:" + dateString.Substring(8, 1) + ",Time:" + timeString.Substring(0, 6);
                MessageBox.Show("GetDateTime OK! Return Value:" + sshowdatetime);
            }
            else
            {
                MessageBox.Show("GetDateTime error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DateTime date;
            DateTime time;
            //if (cbSystemTime.Checked)
            //{
            date = DateTime.Now;
            time = DateTime.Now;
            //}
            //else
            //{
            //    date = dtpDate.Value;
            //    time = dtpTime.Value;
            //}

            //sunday == 7
            //monday == 1
            string sweek = ((int)date.DayOfWeek).ToString();
            if (sweek == "0") { sweek = "7"; }
            string dateString = date.ToString("yyyyMMdd") + sweek; //string.Format("{0:yyyyMMdd}", date);
            string timeString = string.Format("{0:HHmmss}", date);
            int resultCode = -1;
            int timeout = 2000;
            resultCode = TRAC2000DLL.SetDateTime(RAC2000Id, dateString, timeString, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("SetDateTime OK!");
            }
            else
            {
                MessageBox.Show("SetDateTime error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            int dataLen = 112;
            resultCode = TRAC2000DLL.GetEEData(RAC2000Id, dataBuffer, ref dataLen, 0, 112, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                rtbReturn.Clear();
                for (int i = 0; i < dataLen; i++)
                    rtbReturn.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                MessageBox.Show("GetEEData OK!");
            }
            else
            {
                MessageBox.Show("GetEEData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            dataBuffer[0] = Convert.ToByte(Convert.ToChar(1));
            dataBuffer[1] = Convert.ToByte(Convert.ToChar(1));
            resultCode = TRAC2000DLL.SetEEData(RAC2000Id, dataBuffer, 52, 2, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("SetEEData OK!");
            }
            else
            {
                MessageBox.Show("SetEEData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            dataBuffer[0] = Convert.ToByte(Convert.ToChar(4));
            dataBuffer[1] = Convert.ToByte(Convert.ToChar(4));
            resultCode = TRAC2000DLL.SetEEData(RAC2000Id, dataBuffer, 52, 2, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("Restore EEPROM OK!");
            }
            else
            {
                MessageBox.Show("Restore EEPROM error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            int dataLen = 112;
            resultCode = TRAC2000DLL.GetRAMData(RAC2000Id, dataBuffer, ref dataLen, 0, 112, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                rtbReturn.Clear();
                for (int i = 0; i < dataLen; i++)
                    rtbReturn.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                MessageBox.Show("GetRAMData OK!");
            }
            else
            {
                MessageBox.Show("GetRAMData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            dataBuffer[0] = Convert.ToByte(Convert.ToChar(1));
            dataBuffer[1] = Convert.ToByte(Convert.ToChar(1));
            resultCode = TRAC2000DLL.SetRAMData(RAC2000Id, dataBuffer, 52, 2, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("SetRAMData OK!");
            }
            else
            {
                MessageBox.Show("SetRAMData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            dataBuffer[0] = Convert.ToByte(Convert.ToChar(0));
            dataBuffer[1] = Convert.ToByte(Convert.ToChar(0));
            resultCode = TRAC2000DLL.SetRAMData(RAC2000Id, dataBuffer, 52, 2, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("Restore RAMData OK!");
            }
            else
            {
                MessageBox.Show("Restore RAMData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            dataBuffer[0] = Convert.ToByte(Convert.ToChar(1));
            resultCode = TRAC2000DLL.SetEEData(RAC2000Id, dataBuffer, 92, 1, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("SetCompressCard OK!");
            }
            else
            {
                MessageBox.Show("SetCompressCard error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            dataBuffer[0] = Convert.ToByte(Convert.ToChar(0));
            resultCode = TRAC2000DLL.SetEEData(RAC2000Id, dataBuffer, 92, 1, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("SetCompressCard OK!");
            }
            else
            {
                MessageBox.Show("SetCompressCard error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                button21.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                int dataLen = 0;
                // Directly write data in buffer
                // Set the Relay 1 ON of the Device 1
                byte[] dataBuffer = new byte[255];
                dataBuffer[0] = Convert.ToByte(25);
                dataBuffer[1] = Convert.ToByte(0);
                dataBuffer[2] = Convert.ToByte(1);
                dataBuffer[3] = Convert.ToByte(0);
                dataBuffer[4] = Convert.ToByte(0);
                dataBuffer[5] = Convert.ToByte(0);
                dataBuffer[6] = Convert.ToByte(1);
                dataBuffer[7] = Convert.ToByte(252);
                dataBuffer[8] = Convert.ToByte(3);
                dataBuffer[9] = Convert.ToByte(2);
                dataBuffer[10] = Convert.ToByte(1);
                dataBuffer[11] = Convert.ToByte(1);
                dataBuffer[12] = Convert.ToByte(188);
                dataBuffer[13] = Convert.ToByte(55);
                dataBuffer[14] = Convert.ToByte(3);
                TRAC2000DLL.ClearBuffer(RAC2000Handle);
                resultCode = TRAC2000DLL.WriteData(dataBuffer, 15, ref dataLen, RAC2000Handle);

                for (int i = 0; i < dataBuffer.Length; i++)
                {
                    dataBuffer[i] = Convert.ToByte(0);
                }
                resultCode = TRAC2000DLL.ReadData(dataBuffer, ref dataLen, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    rtbReturn.Clear();
                    for (int i = 0; i < dataLen; i++)
                        rtbReturn.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                    MessageBox.Show("WriteData OK!");
                }
                else
                {
                    MessageBox.Show("WriteData error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button21.Enabled = true;
            }
        }


        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                button20.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                int dataLen = 0;
                byte[] dataBuffer = new byte[31];
                resultCode = TRAC2000DLL.GetVersion(RAC2000Id, dataBuffer, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    string versionString = "";
                    for (int i = 0; i < dataBuffer.Length; i++)
                    {
                        if (dataBuffer[i] == 0)
                            break;
                        versionString += Convert.ToChar(dataBuffer[i]);
                    }
                    MessageBox.Show(versionString);
                }
                else
                {
                    MessageBox.Show("GetVersion error!error code :" + Convert.ToString(resultCode));
                }

            }
            finally
            {
                button20.Enabled = true;
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            try
            {
                button36.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                int iSensor = 0x00;
                resultCode = TRAC2000DLL.GetSensor(RAC2000Id, ref iSensor, RAC2000Handle, timeout);
                if (resultCode == 0)
                {

                    string strSensor = iSensor.ToString("X");
                    MessageBox.Show("Read Sensor :" + strSensor);
                }
                else
                {
                    MessageBox.Show("GetSensor error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button36.Enabled = true;
            }
        }


        private void btnDisableAuto_Click(object sender, EventArgs e)
        {
            //Diable Auto Send Mode
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[28];
            resultCode = TRAC2000DLL.hacSetLanMode(RAC2000Id, 0, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                rtbDisplay.Text = string.Format("Disable OK:{0}", TCommon.ByteArrayToString(dataBuffer));
            }
            else
            {
                rtbDisplay.Text = string.Format("Disable Failed!error code :{0}", resultCode);
            }

        }

        private void btnEnableAuto_Click(object sender, EventArgs e)
        {
            //Enable Auto Send Mode
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[28];
            resultCode = TRAC2000DLL.hacSetLanMode(RAC2000Id, 1, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                rtbDisplay.Text = string.Format("Enable OK:{0}", TCommon.ByteArrayToString(dataBuffer));
            }
            else
            {
                rtbDisplay.Text = string.Format("Enable Failed!error code :{0}", resultCode);
            }

        }


        private void button23_Click(object sender, EventArgs e)
        {
            try
            {
                button23.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                byte[] dataBuffer = new byte[512];
                int dataLen = 0;
                resultCode = TRAC2000DLL.GetFlashData(RAC2000Id, dataBuffer, ref dataLen, 0x0069, 10, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    rtbReturn.Clear();
                    for (int i = 0; i < dataLen; i++)
                        rtbReturn2.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                    MessageBox.Show("GetFlashData OK!");
                }
                else
                {
                    MessageBox.Show("GetFlashData error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button23.Enabled = true;
            }

        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                button22.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                byte[] dataBuffer = new byte[255];
                dataBuffer[0] = Convert.ToByte(Convert.ToChar(1));
                resultCode = TRAC2000DLL.SetFlashData(RAC2000Id, dataBuffer, 92, 1, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("SetFlashData OK!");
                }
                else
                {
                    MessageBox.Show("SetFlashData error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button22.Enabled = true;
            }
        }


        // Read ParaData
        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                button19.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                byte[] dataBuffer = new byte[255];
                int dataLen = 0;
                resultCode = TRAC2000DLL.GetParaData(RAC2000Id, dataBuffer, ref dataLen, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    rtbReturn.Clear();
                    for (int i = 0; i < dataLen; i++)
                        rtbReturn2.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                    MessageBox.Show("GetParaData OK!");
                }
                else
                {
                    MessageBox.Show("GetParaData error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button19.Enabled = true;
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            try
            {
                button26.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                byte[] dataBuffer = new byte[255];
                int dataLen = 0;
                resultCode = TRAC2000DLL.GetSysParaData(RAC2000Id, dataBuffer, ref dataLen, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    rtbReturn.Clear();
                    for (int i = 0; i < dataLen; i++)
                        rtbReturn2.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                    MessageBox.Show("GetSysParaData OK!");
                }
                else
                {
                    MessageBox.Show("GetSysParaData error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button26.Enabled = true;
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            try
            {
                button25.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                byte[] dataBuffer = new byte[255];
                dataBuffer[0] = Convert.ToByte(Convert.ToChar(1));
                dataBuffer[1] = Convert.ToByte(Convert.ToChar(1));
                int iFalshLen = 2;
                resultCode = TRAC2000DLL.SetSysParaData(RAC2000Id, dataBuffer, iFalshLen, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("SetSysParaData OK!");
                }
                else
                {
                    MessageBox.Show("SetSysParaData error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button25.Enabled = true;
            }
        }

        //Read Mifare
        private void button24_Click(object sender, EventArgs e)
        {
            try
            {
                button24.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                int iKeyType = 0, iBlock = 0, iStartDigit = 0, iDigitLength = 0, iCompact = 0;

                resultCode = TRAC2000DLL.GetMifare(RAC2000Id, ref iKeyType, ref iBlock, ref iStartDigit,
                    ref iDigitLength, ref iCompact, RAC2000Handle, timeout);

                if (resultCode == 0)
                {
                    rtbReturn2.Text = string.Format("Get MifareCard Parameter OK:(KeyType:{0},iBlock:{1},iStartDigit:{2},iDigitLength:{3},iCompact:{4})",
                        iKeyType, iBlock, iStartDigit, iDigitLength, iCompact);
                    switch (iKeyType)
                    {
                        case 0: rbSerial.Checked = true;
                            break;
                        case 1: rbKeyA.Checked = true;
                            break;
                        case 2: rbKeyB.Checked = true;
                            break;
                    }
                    tbBlock.Text = iBlock.ToString();
                    tbStartBit.Text = iStartDigit.ToString();
                    tbDigitLength.Text = iDigitLength.ToString();
                    cbCompress.Checked = (iCompact == 1);
                }
                else
                {
                    rtbReturn2.Text = string.Format("Get MifareCard Parameter error!error code :{0}", resultCode);
                }
            }
            finally
            {
                button24.Enabled = true;
            }
        }

        //Write Mifare
        private void button27_Click(object sender, EventArgs e)
        {
            try
            {
                button27.Enabled = false;

                //Set Read Card Para
                int resultCode = -1;
                int timeout = 2000;
                int iKeyType = 0, iBlock = 0, iStartDigit = 0, iDigitLength = 0, iCompact = 0;
                //Key Type
                if (rbKeyA.Checked)
                    iKeyType = 1;
                else if (rbKeyB.Checked)
                    iKeyType = 2;
                iBlock = Convert.ToInt32(tbBlock.Text);
                iStartDigit = Convert.ToInt32(tbStartBit.Text);
                iDigitLength = Convert.ToInt32(tbDigitLength.Text);
                if (cbCompress.Checked)
                    iCompact = 1;
                if (tbKey.Text.Length != 12)
                {
                    MessageBox.Show("Key Length Invalid!");
                    return;
                }
                byte[] bKey = new byte[7];
                string sKey = tbKey.Text;
                for (int i = 0; i < 6; i++)
                {
                    bKey[i] = (byte)int.Parse(sKey.Substring(i * 2, 2), System.Globalization.NumberStyles.HexNumber);
                }

                resultCode = TRAC2000DLL.SetMifare(RAC2000Id, iKeyType, iBlock, iStartDigit,
                    iDigitLength, iCompact, bKey, RAC2000Handle, timeout);


                if (resultCode == 0)
                {
                    rtbReturn2.Text = string.Format("Set MifareCard Parameter OK:(KeyType:{0},iBlock:{1},iStartDigit:{2},iDigitLength:{3},iCompact:{4})",
                        iKeyType, iBlock, iStartDigit, iDigitLength, iCompact);
                }
                else
                {
                    rtbReturn2.Text = string.Format("Set MifareCard Parameter error!error code :{0}", resultCode);
                }
            }
            finally
            {
                button27.Enabled = true;
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardNO.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }

            resultCode = TRAC2000DLL.AddCardEX(RAC2000Id, tbCardNO.Text, tbCardNO.Text.Length, "", 0, tbPersonName.Text, tbPersonName.Text.Length, 0, Convert.ToChar(0), RAC2000Handle, timeout);

            if (resultCode == 0)
            {
                MessageBox.Show("AddCard OK!");
            }
            else
            {
                MessageBox.Show("AddCard error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void btnhacFingerPrinterQueryUser_Click(object sender, EventArgs e)
        {
            byte[] f1 = new byte[386];
            byte[] f2 = new byte[386];
            int iLen = 0;
            int iRtn = 0;

            if (TRAC2000DLL.hacFingerPrinterQueryUser(RAC2000Id, RAC2000Handle, txt960CardNo.Text.Length, txt960CardNo.Text, f1, f2, ref iLen, ref iRtn, 6000) == 0)
            { MessageBox.Show(tbCardNO.Text + " success!"); }
            else
            { MessageBox.Show(tbCardNO.Text + " failed!"); }
        }

        private void btnhacAddCardFingerPrint_Click(object sender, EventArgs e)
        {
            byte[] f1 = new byte[386];
            byte[] f2 = new byte[386];
            int iLen = 0;
            int iRtn = 0;
            //Get txt960CardNo.Text FingertPrinter
            if (TRAC2000DLL.hacFingerPrinterQueryUser(RAC2000Id, RAC2000Handle, txt960CardNo.Text.Length, txt960CardNo.Text, f1, f2, ref iLen, ref iRtn, 6000) == 0)
            {
                MessageBox.Show("Retrive Card:" + txt960CardNo.Text + " success!");
                //Based on system parameters to determine the 13th byte
                //There are 2 modes .. Name Display mode / No name Display mode

                //Add FingertPrinter cardNum is 12345
                //Name Display mode (HelloWorld);
                string sNewCard = "0000011111";
                int iretcode = TRAC2000DLL.hacAddCardFingerPrintEx(RAC2000Id, sNewCard, sNewCard.Length, txt960Passwrd.Text, txt960Passwrd.Text.Length, txt960PsersonName.Text, txt960PsersonName.Text.Length, 0, 14, f1, f2, RAC2000Handle, 15000); //Name Displa
                if (iretcode == 0)
                {
                    MessageBox.Show("Insert Card:" + sNewCard + " success!");
                }
                else
                {
                    MessageBox.Show("Insert Card Fail! Return Code:" + iretcode);
                }

                //Add FingertPrinter cardNum is 9876
                //No name Display mode
                //TRAC2000DLL.hacAddCardFingerPrint(RAC2000Id, "9876", 4, "111", 3, 0, 14, f1, f2, RAC2000Handle, 15000);//No name Display

            }
            else
            { MessageBox.Show(txt960CardNo.Text + " failed!"); }
        }

        private void btnhsECUReadIO_Click(object sender, EventArgs e)
        {
            int cSensor = 0;
            int cRelay = 0;
            int iRtn = 0;
            if (TRAC2000DLL.hsECUReadIO(RAC2000Handle, RAC2000Id, ref cSensor, ref cRelay, ref iRtn, 3000) == 0)
            {
                //Bit 0 － Sensor 1的現狀 , 1→Close    0→Open
                //Bit 1 － Sensor 2的現狀 , 1→Close    0→Open
                rtbECU680.Text = "";
                if ((cSensor & 0x01) == 1)
                { rtbECU680.Text += "Sensor1 Close" + "\n"; }
                else
                { rtbECU680.Text += "Sensor1 Open" + "\n"; }
                if ((cSensor & 0x02) == 2)
                { rtbECU680.Text += "Sensor2 Close" + "\n"; }
                else
                { rtbECU680.Text += "Sensor2 Open" + "\n"; }

                //Bit 0 － Relay 1 , 1→動作    0→未動作
                //Bit 1 － Relay 2 , 1→動作    0→未動作

                if ((cRelay & 0x01) == 1)
                { rtbECU680.Text += "Relay1 Action" + "\n"; }
                else
                { rtbECU680.Text += "Relay1 not Action" + "\n"; }
                if ((cRelay & 0x02) == 2)
                { rtbECU680.Text += "Relay2 Action" + "\n"; }
                else
                { rtbECU680.Text += "Relay2 not Action" + "\n"; }
            }
        }

        private void btnhsECUReadParamenter_Click(object sender, EventArgs e)
        {
            //iReturn =hsECUReadParamenter(ghComm,iNodeID,cReceiveBuff,&iReadLen,&iReturnCode,1000);
            byte[] cParaData = new byte[100];
            int iRaraLen = 0;
            int iRtn = 0;
            int resultCode = -1;
            resultCode = TRAC2000DLL.hsECUReadParamenter(RAC2000Handle, RAC2000Id, cParaData, ref iRaraLen, ref iRtn, 3000);
            if (resultCode == 0)
            {
                rtbECU680.Clear();
                for (int i = 0; i < iRaraLen; i++)
                    rtbECU680.AppendText(string.Format("{0,-2}", cParaData[i].ToString("X2")) + " ");
                MessageBox.Show("GetParaData OK!");


            }
            else
            {
                MessageBox.Show("GetParaData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void btnhsECUAddCard_Click(object sender, EventArgs e)
        {
            int resultCode = TRAC2000DLL.hsECUAddCard(RAC2000Handle, RAC2000Id, textBox4.Text, 10, 1, 1, 2010, 12, 31, 23, 59, 1000);
            if (resultCode == 0)
            {
                MessageBox.Show("AddCard OK!");
            }
            else
            {
                MessageBox.Show("AddCard error!error code :" + Convert.ToInt32(resultCode));
            }

        }

        private void btnhsECUPolling_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 10000;
            byte[] dataBuffer = new byte[16640];
            int recordLen = 0;
            int iRtn = 0;
            rtbECU680.Text = "";

            //resultCode = TRAC2000DLL.Polling(RAC2000Id, prevRecord, dataBuffer, ref recordLen, RAC2000Handle, timeout, flag);
            resultCode = TRAC2000DLL.hsECUPolling(RAC2000Handle, RAC2000Id, prevRecord, dataBuffer, ref recordLen, ref  iRtn, timeout);
            prevRecord = recordLen;
            if (resultCode == 0)
            {
                SEventStruct eventRec;
                for (int j = 0; j < recordLen; j++)
                {
                    eventRec = (SEventStruct)TCommon.BytesToStuct(dataBuffer, typeof(SEventStruct), Marshal.SizeOf(typeof(SEventStruct)) * j);
                    rtbECU680.AppendText("Event Code:" + TCommon.ByteArrayToString(eventRec.cEventCode) + " ; ");
                    rtbECU680.AppendText("Date Time:" + TCommon.ByteArrayToString(eventRec.cDateTime) + " ; ");
                    rtbECU680.AppendText("Card Number:" + TCommon.ByteArrayToString(eventRec.cCard) + " ; ");
                    rtbECU680.AppendText("Device ID:" + TCommon.ByteArrayToString(eventRec.cDeviceID) + " ; ");
                    rtbECU680.AppendText("\n");
                }
                rtbECU680.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Polling ok\n");
            }
            else if (resultCode == 1003)
            {
                rtbECU680.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Request equipment is overtime\n");
            }
            else if (resultCode == 1004)
            {
                rtbECU680.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! The handle of value in equipment is false!\n");
            }
            else if (resultCode == 1005)
            {
                rtbECU680.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Transmitting package to equipment  is error!\n");
            }
            else if (resultCode == 1006)
            {
                rtbECU680.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Respond package CRC to equipment error!\n");
            }
            else
            {
                rtbECU680.AppendText("polling id:" + Convert.ToString(RAC2000Id) + "! Polling error!error code :" + Convert.ToString(resultCode) + "\n");
            }

        }

        private void btnhsECUReadPower_Click(object sender, EventArgs e)
        {
            //
            byte cPower = 0;
            byte[] cCard = new byte[10];
            int iRtn = 0;
            if (TRAC2000DLL.hsECUReadPower(RAC2000Handle, RAC2000Id, ref cPower, cCard, ref iRtn, 3000) == 0)
            {
                rtbECU680.AppendText("Card Num:" + Encoding.Default.GetString(cCard));
                rtbECU680.AppendText("Power:" + cPower.ToString() + "\n");
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            string filename = AppDomain.CurrentDomain.BaseDirectory + @"\dump.txt";
            int iRtn = 0;
            iRtn = TRAC2000DLL.hacDumpLegalCard(RAC2000Id, RAC2000Handle, filename, ref iRtn, 5000);
            //iRtn = TRAC2000DLL.hsHTA850DumpFile(RAC2000Handle, filename, ref iRtn, 5000);
            if (iRtn == 0)
            { rtbReturn2.AppendText("hacDumpLegalCard is oK"); }
        }

        private void btnhac34AddCard_Click(object sender, EventArgs e)
        {
            //Add Card
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardID.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }

            resultCode = TRAC2000DLL.hac34AddCard(RAC2000Id, tbCardID.Text, tbCardID.Text.Length, string.Empty, 0, 0, Convert.ToChar(0), RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("AddCard OK!({0})", tbCardID.Text);
            }
            else
            {
                tbExecResult.Text = string.Format("AddCard error!error code :{0}", resultCode);
            }
        }

        private void btnhac34AddCardLoop_Click(object sender, EventArgs e)
        {
            //Add Card Loop
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardID.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            string cardID = tbCardID.Text;
            for (int i = 0; i < 10; i++)
            {
                cardID = "0000000000" + Convert.ToString(Convert.ToInt64(tbCardID.Text) + i);
                cardID = cardID.Substring(cardID.Length - 10);
                resultCode = TRAC2000DLL.hac34AddCard(RAC2000Id, cardID, cardID.Length, string.Empty, 0, 0, Convert.ToChar(0), RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    tbExecResult.Text = string.Format("AddCard OK!({0})\r\n", cardID) + tbExecResult.Text;
                }
                else
                {
                    tbExecResult.Text = string.Format("AddCard ({0}) error!error code :{1}\r\n", cardID, resultCode) + tbExecResult.Text;
                }
                Application.DoEvents();
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardNO.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }

            resultCode = TRAC2000DLL.AddCardEX(RAC2000Id, tbCardNO.Text, tbCardNO.Text.Length, tbPWD.Text, tbPWD.Text.Length, tbPersonName.Text, tbPersonName.Text.Length, 0, Convert.ToChar(0), RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                MessageBox.Show("AddCard OK!");
            }
            else
            {
                MessageBox.Show("AddCard error!error code :" + Convert.ToInt32(resultCode));
            }

        }

        private void button31_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;

            byte[] bsend = new byte[255];
            byte[] breceive = new byte[255];
            int ireceivelen = 0;
            /*
             * 
               Byte 1 : 01H  (fix)
               Byte 2 : Relay action mode
                        00H  force Relay Off
                        01H~FEH  Relay action 1~254 seconds
                        FFH  force Relay On

             */
            bsend[0] = 01;
            bsend[1] = Convert.ToByte(txtRelayOnSec.Text);
            resultCode = TRAC2000DLL.hacHWRWCommandCCH(2, RAC2000Id, 40, bsend, 2, breceive, ref ireceivelen, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                MessageBox.Show("Open Relay OK!");
            }
            else
            {
                MessageBox.Show("Open Relay error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            try
            {
                button32.Enabled = false;

                //Set Read Card Para
                int resultCode = -1;
                int timeout = 2000;
                int iKeyType = 0, iBlock = 0, iStartDigit = 0, iDigitLength = 0, iCompact = 0;
                //Key Type
                if (rbKeyA.Checked)
                    iKeyType = 1;
                else if (rbKeyB.Checked)
                    iKeyType = 2;
                iBlock = Convert.ToInt32(tbBlock.Text);
                iStartDigit = Convert.ToInt32(tbStartBit.Text);
                iDigitLength = Convert.ToInt32(tbDigitLength.Text);
                if (cbCompress.Checked)
                    iCompact = 1;
                if (tbKey.Text.Length != 12)
                {
                    MessageBox.Show("Key Length Invalid!");
                    return;
                }
                byte[] bKey = new byte[7];
                string sKey = tbKey.Text;
                for (int i = 0; i < 6; i++)
                {
                    bKey[i] = (byte)int.Parse(sKey.Substring(i * 2, 2), System.Globalization.NumberStyles.HexNumber);
                }

                resultCode = TRAC2000DLL.SetMifare(RAC2000Id, iKeyType, iBlock, iStartDigit,
                    iDigitLength, iCompact, bKey, RAC2000Handle, timeout);


                if (resultCode == 0)
                {
                    rtbReturn2.Text = string.Format("Set MifareCard Parameter OK:(KeyType:{0},iBlock:{1},iStartDigit:{2},iDigitLength:{3},iCompact:{4})",
                        iKeyType, iBlock, iStartDigit, iDigitLength, iCompact);
                }
                else
                {
                    rtbReturn2.Text = string.Format("Set MifareCard Parameter error!error code :{0}", resultCode);
                }
            }
            finally
            {
                button32.Enabled = true;
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            try
            {
                button33.Enabled = false;

                //Set Read Card Para
                int resultCode = -1;
                int timeout = 2000;
                int iKeyType = 0, iBlock = 0, iStartDigit = 0, iDigitLength = 0, iCompact = 0;
                //Key Type
                if (rbKeyA.Checked)
                    iKeyType = 1;
                else if (rbKeyB.Checked)
                    iKeyType = 2;
                iBlock = Convert.ToInt32(tbBlock.Text);
                iStartDigit = Convert.ToInt32(tbStartBit.Text);
                iDigitLength = Convert.ToInt32(tbDigitLength.Text);
                if (cbCompress.Checked)
                    iCompact = 1;
                if (tbKey.Text.Length != 12)
                {
                    MessageBox.Show("Key Length Invalid!");
                    return;
                }
                byte[] bKey = new byte[7];
                string sKey = tbKey.Text;
                for (int i = 0; i < 6; i++)
                {
                    bKey[i] = (byte)int.Parse(sKey.Substring(i * 2, 2), System.Globalization.NumberStyles.HexNumber);
                }

                resultCode = TRAC2000DLL.SetMifare(RAC2000Id, iKeyType, iBlock, iStartDigit,
                    iDigitLength, iCompact, bKey, RAC2000Handle, timeout);


                if (resultCode == 0)
                {
                    rtbReturn2.Text = string.Format("Set MifareCard Parameter OK:(KeyType:{0},iBlock:{1},iStartDigit:{2},iDigitLength:{3},iCompact:{4})",
                        iKeyType, iBlock, iStartDigit, iDigitLength, iCompact);
                }
                else
                {
                    rtbReturn2.Text = string.Format("Set MifareCard Parameter error!error code :{0}", resultCode);
                }
            }
            finally
            {
                button33.Enabled = true;
            }
        }

        private void btnhac34GetDateTime_Click(object sender, EventArgs e)
        {
            //Get Time
            int resultCode = -1;
            int timeout = 2000;
            byte[] dateBuffer = new byte[10];
            byte[] timeBuffer = new byte[10];
            resultCode = TRAC2000DLL.hac34GetDateTime(RAC2000Id, dateBuffer, timeBuffer, RAC2000Handle, timeout);
            string dateString = TCommon.ByteArrayToString(dateBuffer);
            string timeString = TCommon.ByteArrayToString(timeBuffer);
            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("GetDateTime OK! Date:{0} Time:{1}", dateString.Replace('\0', ' '), timeString.Replace('\0', ' '));
            }
            else
            {
                tbExecResult.Text = string.Format("GetDateTime error!error code :{0}", resultCode);
            }
        }

        private void btnhac34SetDateTime_Click(object sender, EventArgs e)
        {
            //Sync Time
            string sWeek = "7";// ((int)DateTime.Now.DayOfWeek).ToString();
            switch (DateTime.Now.DayOfWeek)
            {

                case DayOfWeek.Monday: sWeek = "1"; break;
                case DayOfWeek.Tuesday: sWeek = "2"; break;
                case DayOfWeek.Wednesday: sWeek = "3"; break;
                case DayOfWeek.Thursday: sWeek = "4"; break;
                case DayOfWeek.Friday: sWeek = "5"; break;
                case DayOfWeek.Saturday: sWeek = "6"; break;
                case DayOfWeek.Sunday: sWeek = "7"; break;
            }
            string dateString = DateTime.Now.ToString("yyyyMMdd") + sWeek;
            string timeString = DateTime.Now.ToString("HHmmss");
            int resultCode = -1;
            int timeout = 2000;
            resultCode = TRAC2000DLL.hac34SetDateTime(RAC2000Id, dateString, timeString, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("SetDateTime (Date:{0} Time:{1}) OK!", dateString, timeString);
            }
            else
            {
                tbExecResult.Text = string.Format("SetDateTime error!error code :{0}", resultCode);
            }
        }

        private void btnhac34Polling_Click(object sender, EventArgs e)
        {
            //Polling
            btnhac34Polling.Enabled = false;
            try
            {
                int resultCode = -1;
                int timeout = 10000;
                byte[] dataBuffer = new byte[16640];
                int recordLen = 0;
                int flag = cbCompress.Checked ? 1 : 0;
                resultCode = TRAC2000DLL.hac34Polling(RAC2000Id, prevRecord, dataBuffer, ref recordLen, RAC2000Handle, timeout, flag);
                prevRecord += recordLen;
                if (resultCode == 0)
                {
                    SEventStruct eventRec;
                    for (int j = 0; j < recordLen; j++)
                    {
                        eventRec = (SEventStruct)TCommon.BytesToStuct(dataBuffer, typeof(SEventStruct), Marshal.SizeOf(typeof(SEventStruct)) * j);
                        string NewEvent = "Event Code:" + TCommon.ByteArrayToString(eventRec.cEventCode) + " ; " +
                            "Date Time:" + TCommon.ByteArrayToString(eventRec.cDateTime) + " ; " +
                            "Card Number:" + TCommon.ByteArrayToString(eventRec.cCard) + " ; " +
                            "Device ID:" + TCommon.ByteArrayToString(eventRec.cDeviceID) + " ; " +
                            "Reader ID:" + TCommon.ByteArrayToString(eventRec.cReaderID);
                        tbExecResult.Text = NewEvent.Trim().Replace('\0', ' ') + "\r\n" + tbExecResult.Text;
                    }

                    tbExecResult.Text = "polling id:" + Convert.ToString(RAC2000Id) + "! Polling ok\r\n" + tbExecResult.Text;
                }
                else if (resultCode == 1003)
                {
                    tbExecResult.Text = "polling id:" + Convert.ToString(RAC2000Id) + "! Request equipment is overtime\r\n" + tbExecResult.Text;
                }
                else if (resultCode == 1004)
                {
                    tbExecResult.Text = "polling id:" + Convert.ToString(RAC2000Id) + "! The handle of value in equipment is false!\r\n" + tbExecResult.Text;
                }
                else if (resultCode == 1005)
                {
                    tbExecResult.Text = "polling id:" + Convert.ToString(RAC2000Id) + "! Transmitting package to equipment  is error!\r\n" + tbExecResult.Text;
                }
                else if (resultCode == 1006)
                {
                    tbExecResult.Text = "polling id:" + Convert.ToString(RAC2000Id) + "! Respond package CRC to equipment error!\r\n" + tbExecResult.Text;
                }
                else
                {
                    tbExecResult.Text = "polling id:" + Convert.ToString(RAC2000Id) + "! Polling error!error code :" + Convert.ToString(resultCode) + "\r\n" + tbExecResult.Text;
                }

            }
            finally
            {
                btnhac34Polling.Enabled = true;
            }
        }

        private void btnhac34SetEEDataEILL_Click(object sender, EventArgs e)
        {
            //Enable Illegal Card
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[280];
            dataBuffer[0] = Convert.ToByte(Convert.ToChar(1));
            resultCode = TRAC2000DLL.hac34SetEEData(RAC2000Id, dataBuffer, 1, 0x46, 1, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                tbExecResult.Text = "Enable Illegal Card:OK!";
            }
            else
            {
                tbExecResult.Text = string.Format("Enable Illegal Card:Error:{0}", resultCode);
            }
        }

        private void btnhac34DelCard_Click(object sender, EventArgs e)
        {
            //Delete Card
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardID.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }

            resultCode = TRAC2000DLL.hac34DelCard(RAC2000Id, tbCardID.Text, tbCardID.Text.Length, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("Delete Card OK!({0})", tbCardID.Text);
            }
            else
            {
                tbExecResult.Text = string.Format("Delete Card error!error code :{0}", resultCode);
            }
        }

        private void btnhac34DelAllCard_Click(object sender, EventArgs e)
        {
            //Delete All Card
            int resultCode = -1;
            int timeout = 20000;
            resultCode = TRAC2000DLL.hac34DelAllCard(RAC2000Id, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("Delete All Card OK!");
            }
            else
            {
                tbExecResult.Text = string.Format("Delete All Card error!error code :{0}", resultCode);
            }
        }

        private void btnhac34SetEEDataDILL_Click(object sender, EventArgs e)
        {
            //Disable Illegal Card
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[280];
            resultCode = TRAC2000DLL.hac34SetEEData(RAC2000Id, dataBuffer, 1, 0x46, 1, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                tbExecResult.Text = "Disable Illegal Card:OK!";
            }
            else
            {
                tbExecResult.Text = string.Format("Disable Illegal Card:Error:{0}", resultCode);
            }
        }

        private void btnhac34DelCardLoop_Click(object sender, EventArgs e)
        {
            //Delete Card Loop
            int resultCode = -1;
            int timeout = 2000;
            if (tbCardID.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            string cardID = tbCardID.Text;
            for (int i = 0; i < 10; i++)
            {
                cardID = "0000000000" + Convert.ToString(Convert.ToInt64(tbCardID.Text) + i);
                cardID = cardID.Substring(cardID.Length - 10);
                resultCode = TRAC2000DLL.hac34DelCard(RAC2000Id, cardID, cardID.Length, RAC2000Handle, timeout);
                if (resultCode == 0)
                {
                    tbExecResult.Text = string.Format("Delete Card OK!({0})\r\n", cardID) + tbExecResult.Text;
                }
                else
                {
                    tbExecResult.Text = string.Format("Delete Card ({0}) error!error code :{1}\r\n", cardID, resultCode) + tbExecResult.Text;
                }
                Application.DoEvents();
            }
        }

        private void hac34SetEEDataCard_Click(object sender, EventArgs e)
        {
            //Door:Card
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[280];
            resultCode = TRAC2000DLL.hac34SetEEData(RAC2000Id, dataBuffer, 1, 0x45, 1, RAC2000Handle, timeout);

            if (resultCode == 0)
            {
                tbExecResult.Text = "Set Door:Card:OK!";
            }
            else
            {
                tbExecResult.Text = string.Format("Set Door:Card:Error:{0}", resultCode);
            }
        }

        private void hac34SetEEDataCardKeyboard_Click(object sender, EventArgs e)
        {
            //Door:Card+Keyboard
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[280];
            dataBuffer[0] = 1;

            resultCode = TRAC2000DLL.hac34SetEEData(RAC2000Id, dataBuffer, 1, 0x45, 1, RAC2000Handle, timeout);

            if (resultCode == 0)
            {
                tbExecResult.Text = "Set Door:Card+Keyboard:OK!";
            }
            else
            {
                tbExecResult.Text = string.Format("Set Door:Card+Keyboard:Error:{0}", resultCode);
            }
        }

        private void btnhac34GetReadCardParameter_Click(object sender, EventArgs e)
        {
            //Get Read Card Para
            int resultCode = -1;
            int timeout = 2000;
            int iKeyType = 0, iBlock = 0, iStartDigit = 0, iDigitLength = 0, iCompact = 0;

            resultCode = TRAC2000DLL.hac34GetReadCardParameter(RAC2000Id, ref iKeyType, ref iBlock, ref iStartDigit,
                ref iDigitLength, ref iCompact, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("Get ReadCard Parameter OK:(KeyType:{0},iBlock:{1},iStartDigit:{2},iDigitLength:{3},iCompact:{4})",
                    iKeyType, iBlock, iStartDigit, iDigitLength, iCompact);
                switch (iKeyType)
                {
                    case 0: rbSerial.Checked = true;
                        break;
                    case 1: rbKeyA.Checked = true;
                        break;
                    case 2: rbKeyB.Checked = true;
                        break;
                }
                tbBlock.Text = iBlock.ToString();
                tbStartBit.Text = iStartDigit.ToString();
                tbDigitLength.Text = iDigitLength.ToString();
                cbCompress.Checked = (iCompact == 1);
            }
            else
            {
                tbExecResult.Text = string.Format("Get ReadCard Parameter error!error code :{0}", resultCode);
            }
        }

        private void btnhac34SetReadCardParameter_Click(object sender, EventArgs e)
        {
            //Set Read Card Para
            int resultCode = -1;
            int timeout = 2000;
            int iKeyType = 0, iBlock = 0, iStartDigit = 0, iDigitLength = 0, iCompact = 0;
            //Key Type
            if (rbKeyA.Checked)
                iKeyType = 1;
            else if (rbKeyB.Checked)
                iKeyType = 2;
            iBlock = Convert.ToInt32(tbBlock.Text);
            iStartDigit = Convert.ToInt32(tbStartBit.Text);
            iDigitLength = Convert.ToInt32(tbDigitLength.Text);
            if (cbCompress.Checked)
                iCompact = 1;
            if (tbKey.Text.Length != 12)
            {
                MessageBox.Show("Key Length Invalid!");
                return;
            }
            byte[] bKey = new byte[7];
            string sKey = tbKey.Text;
            for (int i = 0; i < 6; i++)
            {
                bKey[i] = (byte)int.Parse(sKey.Substring(i * 2, 2), System.Globalization.NumberStyles.HexNumber);
            }

            resultCode = TRAC2000DLL.hac34SetReadCardParameter(RAC2000Id, iKeyType, iBlock, iStartDigit,
                iDigitLength, iCompact, bKey, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("Set ReadCard Parameter OK:(KeyType:{0},iBlock:{1},iStartDigit:{2},iDigitLength:{3},iCompact:{4})",
                    iKeyType, iBlock, iStartDigit, iDigitLength, iCompact);
            }
            else
            {
                tbExecResult.Text = string.Format("Set ReadCard Parameter error!error code :{0}", resultCode);
            }
        }

        private void btnhac34GetVersion_Click(object sender, EventArgs e)
        {
            //Get Version
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[28];
            resultCode = TRAC2000DLL.hac34GetVersion(RAC2000Id, dataBuffer, RAC2000Handle, timeout);


            if (resultCode == 0)
            {
                tbExecResult.Text = string.Format("Get Version OK:{0}", TCommon.ByteArrayToString(dataBuffer));
            }
            else
            {
                tbExecResult.Text = string.Format("Get Version  error!error code :{0}", resultCode);
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (txt960CardNo.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }

            resultCode = TRAC2000DLL.AddCardEX(RAC2000Id, txt960CardNo.Text, txt960CardNo.Text.Length, lb960Password.Text, lb960Password.Text.Length, txt960PsersonName.Text, txt960PsersonName.Text.Length, 0, Convert.ToChar(0), RAC2000Handle, timeout);

            if (resultCode == 0)
            {
                MessageBox.Show("AddCard " + txt960CardNo.Text + " OK!");
            }
            else
            {
                MessageBox.Show("AddCard error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (txt960CardNo.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            resultCode = TRAC2000DLL.DelCard(RAC2000Id, txt960CardNo.Text, txt960CardNo.Text.Length, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("DelCard OK!");
            }
            else
            {
                MessageBox.Show("DelCard error!error code :" + Convert.ToInt32(resultCode));
            }

        }




        private void button37_Click(object sender, EventArgs e)
        {
            int timeout = 2000;
            if (RAC2000Handle > 0)
            {
                byte[] bOutCardData = new byte[1024];
                byte[] bInData = new byte[1024];
                int iretlen = 0;
                int iret = TRAC2000DLL.hacHWRWCommandCCH(2, 1, 4, bOutCardData, 0, bInData, ref iretlen, RAC2000Handle, timeout);
                if (iret == 0)
                {
                    clDecodeDeviceVersionInfo cVersion = new clDecodeDeviceVersionInfo();
                    byte[] bData = new byte[1024];
                    int iDataLength = bInData[9];
                    Array.Copy(bInData, 10, bData, 0, iDataLength);

                    cVersion.doDecode("RAC-960", "RW04", bData);
                    string versionString = "Model:" + cVersion.Devicemodel + "\r\n" +
                         ",MainVersion:" + Convert.ToInt16(cVersion.Version_main).ToString("0") + "." + Convert.ToInt16(cVersion.Version_second).ToString("00") + "," +
                         (cVersion.Version_beta == "00" ? "" : " " + cVersion.Version_beta) +
                         ",ROM Date:" + cVersion.Version_date + "\r\n" +
                      ",legal card:" + cVersion.Devicecurrentlegalcards + ",swip card:" + cVersion.Devicecurrentswipcards;
                    MessageBox.Show("Read Version ok!\r\n" + versionString);
                }
                else
                {
                    MessageBox.Show("Read Version fail!");
                }
            }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            int timeout = 2000;
            if (RAC2000Handle > 0)
            {
                byte[] bOutCardData = new byte[1024];
                byte[] bInData = new byte[1024];
                int iretlen = 0;
                int iret = TRAC2000DLL.hacHWRWCommandCCH(0, RAC2000Id, 30, bOutCardData, 0, bInData, ref iretlen, RAC2000Handle, timeout);
                if (iret == 0)
                {
                    if (bInData[0x0b] == 0x01)
                        ckHID.Checked = true;
                    else
                        ckHID.Checked = false;
                    MessageBox.Show("Get OK!");
                }
                else
                {
                    MessageBox.Show("Get Fail!");
                }
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            int timeout = 2000;
            if (RAC2000Handle > 0)
            {
                byte[] bOutCardData = new byte[1024];
                byte[] bInData = new byte[1024];
                int iretlen = 0;
                int iHid = 0;
                if (ckHID.Checked)
                    iHid = 1;
                else
                    iHid = 0;
                bOutCardData[0] = 0x01; //declare send data 1 byte
                bOutCardData[1] = Convert.ToByte(iHid);
                //The device will be rebooted
                int iret = TRAC2000DLL.hacHWRWCommandCCH(1, RAC2000Id, 30, bOutCardData, 2, bInData, ref iretlen, RAC2000Handle, timeout);
                if (iret == 0)
                {
                    MessageBox.Show("Set OK!");
                }
                else
                {
                    MessageBox.Show("Set Fail!");
                }
            }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            if (cbxWiegand.SelectedIndex < 0)
                return;
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            int dataLen = 112;
            resultCode = TRAC2000DLL.GetEEData(RAC2000Id, dataBuffer, ref dataLen, 0x1A00 + cbxWiegand.SelectedIndex * 10, 10, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                string shexdata = "";
                for (int i = 0; i < dataLen; i++)
                    shexdata += string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ";
                txtWiegandData.Text = shexdata;
                MessageBox.Show("GetEEData OK!");
            }
            else
            {
                MessageBox.Show("GetEEData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = TCommon.HexStrToBytes(txtWiegandData.Text);
            if (dataBuffer.Length != 10)
                return;

            resultCode = TRAC2000DLL.SetEEData(RAC2000Id, dataBuffer, 0x1A00 + cbxWiegand.SelectedIndex * 10, 10, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("SetEEData OK!");
            }
            else
            {
                MessageBox.Show("SetEEData error!error code :" + Convert.ToString(resultCode));
            }
        }

        public static int BCD2Int32(byte inByte)
        {
            return ((((inByte & 0xF0) >> 4) * 10) + (inByte & 0x0F));
        }
        public static byte Int32ToBCD(int inInt32)
        {
            // 轉出 0-99 以內的有效值
            int tempInt = inInt32 % 100;
            int units = tempInt % 10;
            int tens = (tempInt - units) / 10;
            return (byte)((tens << 4) + units);
        }

        private void button44_Click_1(object sender, EventArgs e)
        {




            int resultCode = -1;
            int timeout = 2000;
            byte[] dataBuffer = new byte[255];
            int dataLen = 0;
            resultCode = TRAC2000DLL.GetEEData(RAC2000Id, dataBuffer, ref dataLen, 0xA0, 11, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                byte[] b = new byte[11];
                Array.Copy(dataBuffer, 0, b, 0, 11);
                comboBox1.SelectedIndex = (b[0] > 3 ? 3 : b[0]);
                ucDateOfYear1.UI_Month = BCD2Int32(b[1]);
                ucDateOfYear1.UI_Day = BCD2Int32(b[2]);
                ucTimeOfDay1.UI_Hour = BCD2Int32(b[3]);
                ucTimeOfDay1.UI_Minute = BCD2Int32(b[4]);
                ucDateOfYear2.UI_Month = BCD2Int32(b[5]);
                ucDateOfYear2.UI_Day = BCD2Int32(b[6]);
                ucTimeOfDay2.UI_Hour = BCD2Int32(b[7]);
                ucTimeOfDay2.UI_Minute = BCD2Int32(b[8]);
                ucTimeOfDay3.UI_Hour = BCD2Int32(b[9]);
                ucTimeOfDay3.UI_Minute = BCD2Int32(b[10]);
                MessageBox.Show("GetEEData OK!");
            }
            else
            {
                MessageBox.Show("GetEEData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button45_Click_1(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            byte[] b = new byte[11];
            b[0] += (byte)(comboBox1.SelectedIndex);
            b[1] = Int32ToBCD(ucDateOfYear1.UI_Month);
            b[2] = Int32ToBCD(ucDateOfYear1.UI_Day);
            b[3] = Int32ToBCD(ucTimeOfDay1.UI_Hour);
            b[4] = Int32ToBCD(ucTimeOfDay1.UI_Minute);
            b[5] = Int32ToBCD(ucDateOfYear2.UI_Month);
            b[6] = Int32ToBCD(ucDateOfYear2.UI_Day);
            b[7] = Int32ToBCD(ucTimeOfDay2.UI_Hour);
            b[8] = Int32ToBCD(ucTimeOfDay2.UI_Minute);
            b[9] = Int32ToBCD(ucTimeOfDay3.UI_Hour);
            b[10] = Int32ToBCD(ucTimeOfDay3.UI_Minute);
            resultCode = TRAC2000DLL.SetEEData(RAC2000Id, b, 0xA0, 11, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("SetEEData OK!");
            }
            else
            {
                MessageBox.Show("SetEEData error!error code :" + Convert.ToString(resultCode));
            }
        }

        private void button46_Click(object sender, EventArgs e)
        {
            FrmDump960 fmDump960 = new FrmDump960(RAC2000Id, RAC2000Handle);
            fmDump960.ShowDialog();
        }

        private void button50_Click(object sender, EventArgs e)
        {
            byte[] f1 = new byte[386];
            byte[] f2 = new byte[386];
            int iLen = 0;
            int iRtn = 0;

            if (TRAC2000DLL.hacFingerPrinterQueryUser(RAC2000Id, RAC2000Handle, txt820CardNo.Text.Length, txt820CardNo.Text, f1, f2, ref iLen, ref iRtn, 6000) == 0)
            { MessageBox.Show(txt820CardNo.Text + " success!"); }
            else
            { MessageBox.Show(txt820CardNo.Text + " failed!"); }
        }

        private void button49_Click(object sender, EventArgs e)
        {
            byte[] f1 = new byte[386];
            byte[] f2 = new byte[386];
            int iLen = 0;
            int iRtn = 0;
            //Get txt960CardNo.Text FingertPrinter
            if (TRAC2000DLL.hacFingerPrinterQueryUser(RAC2000Id, RAC2000Handle, txt820CardNo.Text.Length, txt820CardNo.Text, f1, f2, ref iLen, ref iRtn, 6000) == 0)
            {
                MessageBox.Show("Retrive Card:" + txt960CardNo.Text + " success!");
                //Based on system parameters to determine the 13th byte
                //There are 2 modes .. Name Display mode / No name Display mode

                //Add FingertPrinter cardNum is 12345
                //Name Display mode (HelloWorld);
                string sNewCard = "0000011111";
                int iretcode = TRAC2000DLL.hacAddCardFingerPrintEx(RAC2000Id, sNewCard, sNewCard.Length, "", 0, "", 0, 0, 14, f1, f2, RAC2000Handle, 15000); //Name Displa
                if (iretcode == 0)
                {
                    MessageBox.Show("Insert Card:" + sNewCard + " success!");
                }
                else
                {
                    MessageBox.Show("Insert Card Fail! Return Code:" + iretcode);
                }

                //Add FingertPrinter cardNum is 9876
                //No name Display mode
                //TRAC2000DLL.hacAddCardFingerPrint(RAC2000Id, "9876", 4, "111", 3, 0, 14, f1, f2, RAC2000Handle, 15000);//No name Display

            }
            else
            { MessageBox.Show(txt960CardNo.Text + " failed!"); }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            int resultCode = -1;
            int timeout = 2000;
            if (txt960CardNo.Text.Trim() == "")
            {
                MessageBox.Show("The increase card no. can not acts empty,please re-input");
                return;
            }
            resultCode = TRAC2000DLL.DelCard(RAC2000Id, txt820CardNo.Text, txt820CardNo.Text.Length, RAC2000Handle, timeout);
            if (resultCode == 0)
            {
                MessageBox.Show("DelCard OK!");
            }
            else
            {
                MessageBox.Show("DelCard error!error code :" + Convert.ToInt32(resultCode));
            }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            byte[] f1 = new byte[386];
            byte[] f2 = new byte[386];
            int iLen = 0;
            int iRtn = 0;
            //Get txt960CardNo.Text FingertPrinter
            
             
                //Based on system parameters to determine the 13th byte
                //There are 2 modes .. Name Display mode / No name Display mode

                //Add FingertPrinter cardNum is 12345
                //Name Display mode (HelloWorld);
                string sNewCard = txt820CardNo.Text;
                int iretcode = TRAC2000DLL.hacAddCardFingerPrintEx(RAC2000Id, sNewCard, sNewCard.Length, "", 0, "", 0, 0, 12, f1, f2, RAC2000Handle, 15000); //Name Displa
                if (iretcode == 0)
                {
                    MessageBox.Show("Insert Card:" + sNewCard + " success!");
                }
                else
                {
                    MessageBox.Show("Insert Card Fail! Return Code:" + iretcode);
                }

                //Add FingertPrinter cardNum is 9876
                //No name Display mode
                //TRAC2000DLL.hacAddCardFingerPrint(RAC2000Id, "9876", 4, "111", 3, 0, 14, f1, f2, RAC2000Handle, 15000);//No name Display

           
        }

        private void rtbReturn2_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.Text = rtbReturn2.Text;
        }

    }

}