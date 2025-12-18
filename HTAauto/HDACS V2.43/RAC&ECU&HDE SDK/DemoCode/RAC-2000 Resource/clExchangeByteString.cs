using System;
using System.Collections.Generic;
using System.Text;

namespace DEMO
{
    public static class clExchangeByteString
    {
        /// <summary>
        /// 直接將 byte 內的字元轉成 string,如果遇到 byte值為00, 表示字串結束
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public static string ByteToString(byte[] buffer)
        {
            string sBuffer = "";
            for (int i = 0; i < buffer.Length; i++)
            {
                if (buffer[i] == 0)
                    return sBuffer;
                sBuffer += Convert.ToChar(buffer[i]);

            }
            return sBuffer;
        }

        public static string Bytes2HexStr(byte[] bData, bool haveSpace)
        {
            string strData = string.Empty;
            if (bData == null)
                return strData;
            for (int i = 0; i < bData.Length; i++)
                if (haveSpace)
                    strData += bData[i].ToString("X2") + " ";
                else
                    strData += bData[i].ToString("X2");
            return strData;
        }

        public static string Bytes2HexStr(byte[] bData)
        {
            string strData = string.Empty;
            for (int i = 0; i < bData.Length; i++)
            {
                strData += bData[i].ToString("X2");
                if (i < (bData.Length - 1))
                    strData += " ";
            }
            return strData;
        }
        /// <summary>
        /// 將byte陣列轉為二進位字串
        /// </summary>
        /// <param name="bData"></param>
        /// <returns></returns>
        public static string Bytes2BinStr(byte[] bData)
        {
            string strData = string.Empty;
            Array.ForEach(bData, delegate(byte b) { strData += b.ToString("X2"); });
            string s16 = strData;
            strData = "";
            for (int ii = 0; ii < s16.Length; ii += 2)
            {
                string strData1 = Convert.ToString(Convert.ToInt32(s16.Substring(ii, 2), 16), 2);
                strData = strData + strData1.PadLeft(8, '0');
            }
            //strData = Convert.ToString(Convert.ToInt32(strData, 16), 2);
            //strData = strData.PadLeft(bData.Length * 8, '0');
            return strData;
        }

        public static void FillMemory(byte[] myArray, int mySize, byte myUnit)
        {
            for (int i = 0; i < mySize; i++)
            {
                myArray[i] = myUnit;
            }
        }

        public static void FillMemory(byte[] myArray, int inx, int mySize, byte myUnit)
        {
            for (int i = 0; i < mySize; i++)
            {
                myArray[inx + i] = myUnit;
            }
        }

        public static byte[] HexStrToBytes(string hexString)
        {
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }

        public static int BCD2DEC(int x)
        {
            return (((((x) & 0xF0) >> 4) * 10) + ((x) & 0x0F));
        }
        public static int DEC2BCD(int x)
        {
            return ((((x) / 10) << 4) | ((x) % 10));
        }
        public static int Bin2Dec(int binVal)
        {
            int value = 0;
            try
            {
                string s1 = Convert.ToString(binVal, 16);
                value = Convert.ToInt32(s1);
                return value;
            }
            catch
            {
                return 0;
            }
        }
        public static string Bin2Dec(int binVal, int len)
        {
            string value = null;
            //value = string.Format("D2", (Convert.ToInt16(Convert.ToString(binVal, 16))));
            try
            {
                value = Convert.ToInt16(Convert.ToString(binVal, 16)).ToString("D2");
                return value;
            }
            catch
            {
                return "00";
            }
        }
    }
}
