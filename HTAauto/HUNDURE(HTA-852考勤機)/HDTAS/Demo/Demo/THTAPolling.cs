using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Demo
{
    public delegate void OnHintEvent(string sHint);
    public class THTAPolling
    {
        public string address;
        public int port;
        public int HTAID;
        public int compress;
        public OnHintEvent Onhint;
        int HTAHandle;
        int iprevLen;
        bool StatusFlag;

        public THTAPolling(string address, int port, int id)
        {
            this.address = address;
            this.port = port;
            this.HTAID = id;
            StatusFlag = false;
        }
        public bool connectHTA()
        {
            bool blResult = false;
            int resultCode;
            resultCode = THTA830DLL.HUNHTAOpenSocket(ref HTAHandle, address, port);
            string resultStr = "";
            if (resultCode == 0)
            {
                resultStr = "connect HTA(" + address + ":" + Convert.ToString(port) + ":" + Convert.ToString(HTAID) + ") success!";
                blResult = true;
            }
            else
            {
                resultStr = "connect HTA(" + address + ":" + Convert.ToString(port) + ":" + Convert.ToString(HTAID) + ") error! return :" + Convert.ToString(resultCode);
            }
            if (Onhint != null)
            {
                Onhint(resultStr);
            }
            return blResult;
        }

        private bool DisConnectHTA()
        {
            int resultCode = -1;
            resultCode = THTA830DLL.HUNHTACloseSocket(HTAHandle);
            if (resultCode == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void PollingHTA()
        {
            try
            {
                while (StatusFlag)
                {
                    System.Threading.Thread.Sleep(1000);
                    int resultCode = -1;
                    uint timeout = 1000;
                    int dataLength = 0;
                    // EventFormat[] EventData = new EventFormat[255];
                    byte[] EventData = new byte[255 * 60];
                    string show = "";
                    resultCode = THTA830DLL.HUNHTAPolling(HTAHandle, HTAID, iprevLen, EventData, ref dataLength, compress, timeout);
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
                            show = "HTA DATA(" + address + ":" + Convert.ToString(port) + ":" + Convert.ToString(HTAID) + ")--" + "No.:" + Convert.ToString(i + 1) + " " + "Class Code: " + tmpData.ClassCode + "  " +
                                "legal Code: " + tmpData.IllegalCode + "  " +
                                "Date time: " + tmpData.sDateTime + "  " +
                                "Card NO: " + tmpData.sCard + "  " +
                                "Device ID: " + tmpData.sDeviceID;
                            if (Onhint != null)
                            {
                                lock(this)
                                {
                                    Onhint(show);
                                }
                            }
                        }
                    }
                    //the HTA has no data
                    else if (HTAHandle != 0 && resultCode == 1010)
                    {
                        if (Onhint != null)
                        {
                            lock (this)
                            {
                                Onhint("the HTA:(" + address + ":" + Convert.ToString(port) + ":" + Convert.ToString(HTAID) + ") has no data!");
                            }
                        }
                    }
                    else
                    {
                        if (Onhint != null)
                        {
                            lock (this)
                            {
                                Onhint("Polling HTA(" + address + ":" + Convert.ToString(port) + ":" + Convert.ToString(HTAID) + ") error!return:" + resultCode.ToString());
                            }
                        } 
                    }
                }
                DisConnectHTA();
            }
            catch (Exception ex)
            {
                if (Onhint != null)
                {
                    Onhint("Polling HTA(" + address + ":" + Convert.ToString(port) + ":" + Convert.ToString(HTAID) + ") error!exception:" + ex.Message);
                } 
            }
        }

        public void start()
        {
            StatusFlag = true;
            Thread thread = new Thread(new ThreadStart(PollingHTA));
            thread.Start();    
        }

        public void stop()
        {
            StatusFlag = false;
        }
    }
}
