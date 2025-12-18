using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;




namespace DEMO
{
    public class clDecodeDeviceVersionInfo
    {
        private string _devicemodel;
        
        public string Devicemodel
        {
            get { return _devicemodel; }
            set { _devicemodel = value; }
        }
        private string _version_main;
        
        public string Version_main
        {
            get { return _version_main; }
            set { _version_main = value; }
        }
        private string _version_second;
        
        public string Version_second
        {
            get { return _version_second; }
            set { _version_second = value; }
        }
        private string _version_beta;
        
        public string Version_beta
        {
            get { return _version_beta; }
            set { _version_beta = value; }
        }
        private string _version_date;
        
        public string Version_date
        {
            get { return _version_date; }
            set { _version_date = value; }
        }
        private int _devicemaxlegalcard;
        
        public int Devicemaxlegalcard
        {
            get { return _devicemaxlegalcard; }
            set { _devicemaxlegalcard = value; }
        }
        private int _devicemaxswipcard;
        
        public int Devicemaxswipcard
        {
            get { return _devicemaxswipcard; }
            set { _devicemaxswipcard = value; }
        }
        private int _devicecurrentlegalcards;
        
        public int Devicecurrentlegalcards
        {
            get { return _devicecurrentlegalcards; }
            set { _devicecurrentlegalcards = value; }
        }
        private int _devicecurrentswipcards;
        
        public int Devicecurrentswipcards
        {
            get { return _devicecurrentswipcards; }
            set { _devicecurrentswipcards = value; }
        }


        public void doDecode(string sDeviceModel, string sReadWriteCommand, byte[] bVersionData)
        {
            string sComBine = sDeviceModel + "." + sReadWriteCommand;
            switch (sComBine)
            {
                case "RAC-940.RW01":
                    doDeocde940_RW01(bVersionData);
                    break;
                case "RAC-940.R05":
                    doDeocde_R05(bVersionData);
                    break;
                case "RAC-960.RW04":  //version
                    doDeocde_RW04(bVersionData);
                    break;
                case "PXR-FP.R100":
                    doDeocde_FPPXR100(bVersionData);
                    break ;
                default:
                    break;
            }

        }

        private void doDeocde940_RW01(byte[] bVersionData)
        {
            if (bVersionData.Length >= 24)
            {
                if (bVersionData[0] == 0x34 & bVersionData[1] == 0x45)
                {
                    _devicemodel = "RAC-940PE";
                }
                if (bVersionData[0] == 0x34 & bVersionData[1] == 0x4D)
                {
                    _devicemodel = "RAC-940PM";
                }
                int iaddr = 2;
                _version_main = bVersionData[iaddr].ToString("X2");
                iaddr = 3;
                _version_second = bVersionData[iaddr].ToString("X2");
                iaddr = 4;
                _version_beta = bVersionData[iaddr].ToString("X2");
                iaddr = 5;
                _version_date = "20" + bVersionData[iaddr].ToString("X2") + "/" + bVersionData[iaddr + 1].ToString("X2") + "/" + bVersionData[iaddr + 2].ToString("X2");
                iaddr = 8;
                _devicemaxlegalcard = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
                iaddr = 12;
                _devicemaxswipcard = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
                iaddr = 16;
                _devicecurrentlegalcards = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
                iaddr = 20;
                _devicecurrentswipcards = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
            }
        }

        private void doDeocde_R05(byte[] bVersionData)
        {
            if (bVersionData.Length >= 7)
            {
                if (bVersionData[0] == 0x34 & bVersionData[1] == 0x45)
                {
                    _devicemodel = "RAC-940PE";
                }
                if (bVersionData[0] == 0x34 & bVersionData[1] == 0x4D)
                {
                    _devicemodel = "RAC-940PM";
                }
                int iaddr = 2;
                _version_main = bVersionData[iaddr].ToString("X2");
                iaddr = 3;
                _version_second = bVersionData[iaddr].ToString("X2");

                iaddr = 4;
                _version_date = "20" + bVersionData[iaddr].ToString("X2") + "/" + bVersionData[iaddr + 1].ToString("X2") + "/" + bVersionData[iaddr + 2].ToString("X2");

            }
        }

        private void doDeocde_RW04(byte[] bVersionData)
        {
            string sDeviceIdentify =
@"{0xA1,RAC-960PE/RAC-960PM/RAC-960PME};
{0xA2,RAC-960PEF/RAC-960PMF/RAC-960F};
{0xA3,RAC-960PMD};
{0xA6,RAC-960PCRF};
{0xB1,HTA-860PE/HTA-860PEPM};
{0xB2,HTA-860PEF/HTA-860PMF/HTA-860F};
{0xC1,HDE-960PE/HDE-960PM};
{0xC3,HDE-970PE-R/HDE-970PM-R};
{0xD1,HTA-856PE/HTA-856PM};
{0xE1,RAC-970PE/RAC-970PM};
{0xE2,RAC-970PEF/RAC-970PMF};
{0xF1,HTA-870PE/HTA-870PM};
{0xF2,HTA-870PEF/HTA-870PMF};
{0x00,HTA-850PE/HTA-850PM};
{0x03,HTA-852PEF/HTA-852PMF};
{0x04,RAC-852PxFV};
{0x06,RAC-820PxFV};
{0x11,PXR-96EFSK};
{0x12,PXR-96MFSK};
{0x13,PXR-96FSK};
{0x14,PXR-96CRFSK};
{0x21,PXR-96EFSKL};
{0x22,PXR-96MFSKL};
{0x23,PXR-96FSKL};
{0x31,PXR-97EFSK};
{0x32,PXR-97MFSK};
{0x33,PXR-97FSK};
{0x41,PXR-97EFSKL};
{0x42,PXR-97MFSKL};
{0x43,PXR-97FSKL};
{0x60,RAC-820PxF}";
            //V01 add start
            sDeviceIdentify += ";{0x62,RAC-810PMF}";
            //V01 add end
            string sretcode = "";
            MatchCollection matches = Regex.Matches(sDeviceIdentify, "{0x" + bVersionData[0].ToString("X2") + ".+?}", RegexOptions.IgnoreCase);
            if (matches.Count > 0)
            {
                string[] sField = matches[0].Value.Substring(1, matches[0].Length - 2).Split(',');
                _devicemodel = sField[1];
            }
            if (bVersionData.Length >= 16)
            {
                
                int iaddr = 1;
                _version_beta = "";
                _version_main = bVersionData[iaddr].ToString("X2");
                iaddr = 2;
                _version_second = bVersionData[iaddr].ToString("X2");

                iaddr = 4;
                _version_date = "20" + bVersionData[iaddr].ToString("X2") + "/" + bVersionData[iaddr + 1].ToString("X2") + "/" + bVersionData[iaddr + 2].ToString("X2");
                iaddr = 7;
                _devicecurrentlegalcards = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
                iaddr = 11;
                _devicecurrentswipcards = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
            }
        }


        private void doDeocde_FPPXR100(byte[] bVersionData)
        {
            string sDeviceIdentify =
@"{0x1C,PXR-82MFS};
{0x7D,PXR-96EFSK};
{0x7E,PXR-96MFSK};
{0x7F,PXR-96FSK};
{0x81,PXR-82EFS/FS};
{0x86,PXR-96CRFSK};
{0xB0,PXR-852 EFVSK};
{0xB1,PXR-852 MFVSK};
{0xB2,PXR-852 CRFVSK};
{0x54,PXR-96EFSKL};
{0x55,PXR-96MFSKL};
{0x56,PXR-96FSKL};
{0x57,PXR-82MFSL};
{0x58,PXR-82EFSL/FSL};
{0xA6,PXR-82MFSL};
{0xAD,PXR-82MFVS};
{0xAE,PXR-82CFVS};
{0xAF,PXR-82CRFVS}";
            string sretcode = "";
            MatchCollection matches = Regex.Matches(sDeviceIdentify, "{0x" + bVersionData[0].ToString("X2") + ".+?}", RegexOptions.IgnoreCase);
            if (matches.Count > 0)
            {
                string[] sField = matches[0].Value.Substring(1, matches[0].Length - 2).Split(',');
                _devicemodel = sField[1];
            }
            if (bVersionData.Length >= 16)
            {
                
                int iaddr = 1;
                _version_main = bVersionData[iaddr].ToString("X2");
                iaddr = 2;
                _version_second = bVersionData[iaddr].ToString("X2");

                iaddr = 3;
                _version_date = "20" + bVersionData[iaddr].ToString("X2") + "/" + bVersionData[iaddr + 1].ToString("X2") + "/" + bVersionData[iaddr + 2].ToString("X2");
                iaddr = 7;
                _devicecurrentlegalcards = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
                iaddr = 11;
                _devicecurrentswipcards = bVersionData[iaddr] + bVersionData[iaddr + 1] * 256 + bVersionData[iaddr + 2] * 65536 + bVersionData[iaddr + 3] * 16777216;
            }
        }
    }
}
