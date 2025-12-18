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
    public partial class FrmRAC2000EL : Form
    {
        public UInt32 RAC2000Handle;
        public string Address;
        public int Port;
        public int RAC2000Id;

        private int prevRecord;

        public FrmRAC2000EL()
        {
            InitializeComponent();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                button9.Enabled = false;
                int resultCode = -1;
                int timeout = 2000;
                byte[] dateBuffer = new byte[10];
                byte[] timeBuffer = new byte[8];
                int ReturnCode = 0;
                resultCode = TRAC2000ELDLL.GetDateTime(RAC2000Handle, RAC2000Id, dateBuffer, timeBuffer, ref ReturnCode, timeout);
                string dateString = TCommon.ByteArrayToString(dateBuffer);
                string timeString = TCommon.ByteArrayToString(timeBuffer);
                DateTime date = Convert.ToDateTime(dateString);
                DateTime time = Convert.ToDateTime(timeString);
                //DateTime date = Convert.ToDateTime(dateString.Substring(0, 4) + "/" + dateString.Substring(4, 2) + "/" + dateString.Substring(6, 2));
                //DateTime time = Convert.ToDateTime(timeString.Substring(0, 2) + ":" + timeString.Substring(2, 2) + ":" + timeString.Substring(4, 2));
                if (resultCode == 0)
                {
                    dtpDate.Value = date;
                    dtpTime.Value = time;
                    MessageBox.Show("GetDateTime OK!" + dateString + ":" + timeString);
                }
                else
                {
                    MessageBox.Show("GetDateTime error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button9.Enabled = true;
            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                button9.Enabled = true;
                DateTime date;
                DateTime time;
                int ReturnCode = 0;
                date = DateTime.Now;
                time = DateTime.Now;

                string dateString = string.Format("{0:yyyyMMdd}", date);
                string timeString = string.Format("{0:HHmmss}", date);
                int resultCode = -1;
                int timeout = 2000;
                resultCode = TRAC2000ELDLL.SetDateTime(RAC2000Handle, RAC2000Id, dateString, timeString, ref ReturnCode, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("SetDateTime OK!");
                }
                else
                {
                    MessageBox.Show("SetDateTime error!error code :" + Convert.ToString(resultCode));
                }
            }
            finally
            {
                button9.Enabled = true;
            }
        }

        //write table
        private void button11_Click(object sender, EventArgs e)
        {
            button11.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                dataBuffer[0] = 0x00;
                dataBuffer[1] = 0x00;
                dataBuffer[2] = 0xF0;
                dataBuffer[3] = 0x00;
                dataBuffer[4] = 0X07;
                dataBuffer[5] = 0X31;
                dataBuffer[6] = 0X31;
                dataBuffer[7] = 0X31;
                dataBuffer[8] = 0X32;
                dataBuffer[9] = 0X40;
                dataBuffer[10] = 0X31;

                resultCode = TRAC2000ELDLL.WriteTable(RAC2000Handle, RAC2000Id, dataBuffer, dataLength, ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Write Table:OK!");
                }
                else
                {
                    MessageBox.Show("Write Table error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button11.Enabled = true;
            }
        }

        //read table
        private void button12_Click(object sender, EventArgs e)
        {
            button12.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                dataBuffer[0] = 0x00;
                dataBuffer[1] = 0x00;
                dataBuffer[2] = 0xF0;
                dataBuffer[3] = 0x00;
                dataBuffer[4] = 0X07;
                int iReadLen = 10;
                resultCode = TRAC2000ELDLL.ReadTable(RAC2000Handle, RAC2000Id, dataBuffer, ref iReadLen, ref result, timeout);
                iReadLen = 10;
                if (resultCode == 0)
                {
                    rtbTable.Clear();
                    rtbTable.AppendText("Read Table :");
                    for (int i = 0; i < iReadLen; i++)
                        rtbTable.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                    MessageBox.Show("Read Table:OK!");
                }
                else
                {
                    MessageBox.Show("Read Table error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button12.Enabled = true;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            button13.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                int iReadLen = 0;
                resultCode = TRAC2000ELDLL.hsELReadDeviceInfo(RAC2000Handle, RAC2000Id, dataBuffer, ref iReadLen, ref result, timeout);
                if (resultCode == 0)
                {
                    rtbReturn.Clear();
                    string content = "";
                    for (int i = 0; i < iReadLen; i++)
                    {
                        content += string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ";
                    }
                    rtbReturn.AppendText("Get Info :" + content + "\n");
                    MessageBox.Show("Get Info OK");
                }
                else
                {
                    MessageBox.Show("Get Info error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button13.Enabled = true;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            button14.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                int iReadLen = 0;
                resultCode = TRAC2000ELDLL.GetELInfo(RAC2000Handle, RAC2000Id, dataBuffer, ref iReadLen, ref result, timeout);
                if (resultCode == 0)
                {
                    rtbReturn.Clear();
                    string content = "";
                    for (int i = 0; i < iReadLen; i++)
                    {
                        content += string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ";
                    }
                    rtbReturn.AppendText("GetRac2000ELInfo :" + content + "\n");
                    MessageBox.Show("Get Info OK");
                }
                else
                {
                    MessageBox.Show("Get Info error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button14.Enabled = true;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            button15.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                resultCode = TRAC2000ELDLL.Initialize(RAC2000Handle, RAC2000Id, "31", ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Initial RAC2000EL:OK!");
                }
                else
                {
                    MessageBox.Show("Initial RAC2000EL:Error! :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button15.Enabled = true;
            }

        }

        private void button16_Click(object sender, EventArgs e)
        {
            button16.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                resultCode = TRAC2000ELDLL.hsELReleaseAlarm(RAC2000Handle, RAC2000Id, "31", ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Release Alarm RAC2000EL:OK!");
                }
                else
                {
                    MessageBox.Show("Release Alarm RAC2000EL:Error! :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button16.Enabled = true;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            button18.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                dataBuffer[0] = 0x00;
                dataBuffer[1] = 0x00;
                dataBuffer[2] = 0x13;
                dataBuffer[3] = 0x00;
                int iReadLen = 4;
                resultCode = TRAC2000ELDLL.hsELReadParameter(RAC2000Handle, RAC2000Id, dataBuffer, ref iReadLen, ref result, timeout);
                if (resultCode == 0)
                {
                    rtbReturn.Clear();
                    rtbReturn.AppendText("Read Parameter :");
                    for (int i = 0; i < iReadLen; i++)
                        rtbReturn.AppendText(string.Format("{0,-2}", dataBuffer[i].ToString("X2")) + " ");
                    MessageBox.Show("Read Table:OK!");
                }
                else
                {
                    MessageBox.Show("Read Parameter error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button18.Enabled = true;
            }
        }

        //write Parameter
        private void button19_Click(object sender, EventArgs e)
        {
            button19.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                dataBuffer[0x0] = 0x00;
                dataBuffer[0x1] = 0x00;
                dataBuffer[0x2] = 0x13;
                dataBuffer[0x3] = 0x00;

                dataBuffer[0x4] = 0x00;
                dataBuffer[0x5] = 0x00;
                dataBuffer[0x6] = 0xE8;
                dataBuffer[0x7] = 0x03;
                dataBuffer[0x8] = 0x00;
                dataBuffer[0x9] = 0x30;
                dataBuffer[0xa] = 0x30;
                dataBuffer[0xb] = 0x30;
                dataBuffer[0xc] = 0x30;
                dataBuffer[0xd] = 0x00;
                dataBuffer[0xe] = 0x00;
                dataBuffer[0xf] = 0x00;
                dataBuffer[0x10] = 0x00;
                dataBuffer[0x11] = 0x00;
                dataBuffer[0x12] = 0x00;
                dataBuffer[0x13] = 0x00;
                dataBuffer[0x14] = 0x00;

                dataBuffer[0x15] = 0x00;
                dataBuffer[0x16] = 0x00;
                int iWrittenLen = 23;

                resultCode = TRAC2000ELDLL.hsELWriteParameter(RAC2000Handle, RAC2000Id, dataBuffer, iWrittenLen, ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Write Parameter:OK!");
                }
                else
                {
                    MessageBox.Show("Write Parameter error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button19.Enabled = true;
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            button20.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                int iTotal = 8;

                dataBuffer[0] = 0xff;
                dataBuffer[1] = 0xfe;
                dataBuffer[2] = 0xff;
                dataBuffer[3] = 0xff;
                dataBuffer[4] = 0xff;
                dataBuffer[5] = 0xff;
                dataBuffer[6] = 0xff;
                dataBuffer[7] = 0xff;

                int iReadLen = 0;
                resultCode = TRAC2000ELDLL.hsELPublicFloor(RAC2000Handle, RAC2000Id, iTotal, dataBuffer, ref iReadLen, ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Set Public Floor:OK!");
                }
                else
                {
                    MessageBox.Show("Set Public Floor error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button20.Enabled = true;
            }
        }

        /// <summary>
        /// convert struct to byte array
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="structType"></param>
        private byte[] SCardFormatToBytes(object structObj)
        {
            //得到结构体的大小
            int size = Marshal.SizeOf(structObj);
            //创建byte数组
            byte[] bytes = new byte[size];
            //分配结构体大小的内存空间
            IntPtr structPtr = Marshal.AllocHGlobal(size);
            //将结构体拷到分配好的内存空间
            Marshal.StructureToPtr(structObj, structPtr, false);
            //从内存空间拷到byte数组
            Marshal.Copy(structPtr, bytes, 0, size);
            //释放内存空间
            Marshal.FreeHGlobal(structPtr);
            //返回byte数组
            return bytes;
        }

        //暂时添加不成功

        private void button17_Click(object sender, EventArgs e)
        {
            button20.Enabled = false;
            try
            {
                if (tbCardNO.Text.Trim() == "")
                {
                    MessageBox.Show("The increase card no. can not acts empty,please re-input");
                    return;
                }
                if (tbCardNO.Text.Trim().Length < 10)
                {
                    MessageBox.Show("Length should not number less than 10");
                    return;
                }

                CardStruct stCard = new CardStruct();
                int resultCode = -1;
                int result = -1;
                int timeout = 10000;
                int iTotal = 0x01;

                byte[] cardListByte = new byte[Marshal.SizeOf(typeof(CardStruct)) * (3 - 1)];
                stCard.iType = 0x00;
                stCard.iTime = 255;
                stCard.iHoliday = 0x00;
                stCard.cCardNo = new char[20];
                Convert.ToString(tbCardNO.Text.ToString().Trim()).ToCharArray().CopyTo(stCard.cCardNo, 0);
                stCard.cActiveFloor = new char[8];
                Convert.ToString(0xff).ToCharArray().CopyTo(stCard.cActiveFloor, 0); ;
                byte[] cardByte = SCardFormatToBytes(stCard);
                for (int j = 0; j < Marshal.SizeOf(typeof(CardStruct)); j++)
                {
                    cardListByte[Marshal.SizeOf(typeof(CardStruct)) + j] = cardByte[j];
                }

                resultCode = TRAC2000ELDLL.hsELAddAuthorization(RAC2000Handle, RAC2000Id, iTotal, cardByte, ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Insert Card:OK!");
                }
                else
                {
                    MessageBox.Show("Insert Card error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }

            }
            finally
            {
                button20.Enabled = true;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            button8.Enabled = false;
            try
            {
                if (tbCardNO.Text.Trim() == "")
                {
                    MessageBox.Show("The increase card no. can not acts empty,please re-input");
                    return;
                }

                int resultCode = -1;
                int result = -1;
                int timeout = 5000;
                resultCode = TRAC2000ELDLL.hsELDeleteAuthorization(RAC2000Handle, RAC2000Id, tbCardNO.Text.ToString().Trim(), ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Delete Card:OK!");
                }
                else
                {
                    if (result == 0x06)
                    {
                        MessageBox.Show("Delete Card:Not Exist!");
                    }
                    else
                    {
                        MessageBox.Show("Delete Card error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                    }
                }
            }
            finally
            {
                button8.Enabled = true;
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            button40.Enabled = false;
            try
            {
                int resultCode = -1;
                int result = 0;
                int timeout = 5000;

                resultCode = TRAC2000ELDLL.hsELDeleteAllAuthorization(RAC2000Handle, RAC2000Id, ref result, timeout);
                if (resultCode == 0)
                {
                    MessageBox.Show("Delete All Card:OK!");
                }
                else
                {
                    MessageBox.Show("Delete All Card error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                }
            }
            finally
            {
                button40.Enabled = true;
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            button41.Enabled = false;
            try
            {
                if (tbCardNO.Text.Trim() == "")
                {
                    MessageBox.Show("The increase card no. can not acts empty,please re-input");
                    return;
                }
                int resultCode = -1;
                int result = -1;
                int timeout = 1000;
                byte[] dataBuffer;
                int dataLength = 240;
                dataBuffer = new byte[dataLength];
                int iReadLen = 0;
                resultCode = TRAC2000ELDLL.hsELQueryAuthorization(RAC2000Handle, RAC2000Id, tbCardNO.Text.ToString().Trim(), dataBuffer, ref iReadLen, ref result, timeout);
                if (resultCode == 0)
                {
                    string sBuffer = "";
                    //check 7 byte card
                    int cardlen7 = 0;
                    if (dataBuffer[1] > 128)
                    {
                        cardlen7 = dataBuffer[1] - 128;
                    }

                    for (int i = 1; i < dataBuffer.Length; i++)
                    {
                        if (dataBuffer[i] == 0)
                            break;
                        if (cardlen7 > 0 & i > 1)
                        {
                            sBuffer += dataBuffer[i].ToString("X2");
                        }
                        else
                        {
                            //7Byte card, ignore first byte 
                            if (cardlen7 == 0)
                            {
                                sBuffer += Convert.ToChar(dataBuffer[i]);
                            }
                        }
                    }
                    if (sBuffer == tbCardNO.Text.ToString().Trim())
                    {
                        MessageBox.Show("Query Card Found:OK!");
                    }
                    else
                    {
                        MessageBox.Show("Query Card Not Found:OK!");
                    }
                }
                else
                {
                    if (result == 0x06)
                    {
                        MessageBox.Show("Query Card:Not Exist!");
                    }
                    else
                    {
                        MessageBox.Show("Query Card:Error!error code :" + Convert.ToString(resultCode) + "!return code:" + result);
                    }
                }
            }
            finally
            {
                button41.Enabled = true;
            }

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
                int iPrevRecord = 0;
                int result = 0;

                resultCode = TRAC2000ELDLL.Polling(RAC2000Handle, RAC2000Id, prevRecord, dataBuffer, ref recordLen, ref result, timeout);
                prevRecord = recordLen;
                if (resultCode == 0)
                {
                    SEventStruct eventRec;
                    for (int j = 0; j < recordLen; j++)
                    {
                        eventRec = (SEventStruct)TCommon.BytesToStuct(dataBuffer, typeof(SEventStruct), Marshal.SizeOf(typeof(SEventStruct)) * j);
                        rtbDisplay.AppendText("Event Code:" + TCommon.ByteArrayToString(eventRec.cEventCode) + ":");
                        rtbDisplay.AppendText("Date Time:" + TCommon.ByteArrayToString(eventRec.cDateTime) + " :");
                        rtbDisplay.AppendText("Card Number:" + TCommon.ByteArrayToString(eventRec.cCard) + ":");
                        rtbDisplay.AppendText("Device ID:" + TCommon.ByteArrayToString(eventRec.cDeviceID) + ":");
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

        private void button6_Click(object sender, EventArgs e)
        {
            rtbDisplay.Clear();
        }

        private void button12_Click_1(object sender, EventArgs e)
        {

        }




    }
}