using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace DEMO
{

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    struct RTC_STRUCT_ret_time
    {


        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Years;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Months;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Days;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Date; //date of week
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Hours;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Minutes;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Seconds;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] reserve;
    }

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    struct SEventStruct
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 5)]
        public byte[] cEventCode;      // Event ID
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 20)]
        public byte[] cDateTime;       // Datetime
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 20)]
        public byte[] cCard;           // Card NO.
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10)]
        public byte[] cDeviceID;       // Device ID  From 1 start count
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10)]
        public byte[] cReaderID;       // Reader ID
    }
    class TCommon
    {

        public static object BytesToStuct(byte[] bytes, Type type)
        {

            int size = Marshal.SizeOf(type);
            if (size > bytes.Length)
            {
                //返回空

                return null;
            }

            IntPtr structPtr = Marshal.AllocHGlobal(size);
            Marshal.Copy(bytes, 0, structPtr, size);

            object obj = Marshal.PtrToStructure(structPtr, type);

            Marshal.FreeHGlobal(structPtr);


            return obj;
        }

        public static object BytesToStuct(byte[] bytes, Type type, int startPosition)
        {
            int size = Marshal.SizeOf(type);
            byte[] buffer = new byte[size];
            for (int i = 0; i < size; i++)
            {
                buffer[i] = bytes[startPosition + i];
            }
            object obj = BytesToStuct(buffer, type);
            return obj;
        }

        /// <summary>
        /// convert byte array to struct array
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="type"></param>
        /// <param name="amount"></param>
        /// <returns></returns>
        public static object[] BufferToStuct(byte[] bytes, Type type, int startPos, int amount)
        {
            int size = Marshal.SizeOf(type);
            byte[] buffer = new byte[size];

            //account total of type
            int objTotal;
            if (amount != 0)
            {
                if (amount < ((bytes.Length - startPos) / size))
                {
                    objTotal = amount;
                }
                else
                {
                    objTotal = (bytes.Length - startPos) / size;
                }
            }
            else
            {
                objTotal = (bytes.Length - startPos) / size;
            }
            object[] objList = new object[objTotal];

            //full data
            for (int i = 0; i < objTotal; i++)
            {
                for (int j = 0; j < size; j++)
                {
                    buffer[j] = bytes[startPos + i * size + j];
                }
                object obj = BytesToStuct(buffer, type);
                objList[i] = obj;
            }
            return objList;
        }


        public static byte[] StructToBytes(object structObj)
        {

            int size = Marshal.SizeOf(structObj);
            byte[] bytes = new byte[size];
            IntPtr structPtr = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(structObj, structPtr, false);
            Marshal.Copy(structPtr, bytes, 0, size);
            Marshal.FreeHGlobal(structPtr);
            return bytes;
        }

        public static string ByteArrayToString(byte[] buffer)
        {
            string sBuffer = "";
            for (int i = 0; i < buffer.Length; i++)
            {
                sBuffer += Convert.ToChar(buffer[i]);
            }
            return sBuffer;
        }

        public static string ByteArrayToString(byte[] buffer, int length)
        {
            string sBuffer = "";
            for (int i = 0; i < length; i++)
            {
                sBuffer += Convert.ToChar(buffer[i]);
            }
            return sBuffer;
        }

        public static string ByteArrayToHexString(byte[] ba, int length)
        {
            if (length == 0)
                length = ba.Length;
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            for (int i = 0; i < length; i++)
            {

                hex.AppendFormat("{0:x2}", ba[i]);
            }
            return hex.ToString();
        }

        public static byte[] HexStrToBytes(string hexString)
        {
            if (hexString.IndexOf(" ") > 0)
                hexString = hexString.Replace(" ", "");
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }
    }
}
