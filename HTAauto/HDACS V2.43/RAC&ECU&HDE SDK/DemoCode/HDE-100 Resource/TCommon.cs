using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace DEMO
{
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    struct SEventStruct
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 12)]
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
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct CardStruct
    {
        public int iType;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 20)]
        public char [] cCardNo;
        public int iHoliday;
        public int iTime;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
        public char [] cActiveFloor;

    }

    class TCommon
    {
        /// <summary>
        /// convert byte array to struct
        /// </summary>
        /// <param name="bytes">byte数组</param>
        /// <param name="type">结构体类型</param>
        /// <returns>转换后的结构体</returns>
        public static object BytesToStuct(byte[] bytes, Type type)
        {
            //得到结构体的大小
            int size = Marshal.SizeOf(type);
            //byte数组长度小于结构体的大小
            if (size > bytes.Length)
            {
                //返回空

                return null;
            }
            //分配结构体大小的内存空间
            IntPtr structPtr = Marshal.AllocHGlobal(size);
            //将byte数组拷到分配好的内存空间
            Marshal.Copy(bytes, 0, structPtr, size);
            //将内存空间转换为目标结构体

            object obj = Marshal.PtrToStructure(structPtr, type);
            //释放内存空间
            Marshal.FreeHGlobal(structPtr);
            //返回结构体

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

        /// <summary>
        /// convert struct to byte array
        /// </summary>
        /// <param name="structObj">要转换的结构体</param>
        /// <returns>转换后的byte数组</returns>
        public static byte[] StructToBytes(object structObj)
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

        public static string ByteArrayToString(byte[] buffer)
        {
            string sBuffer = "";
            for (int i = 0; i < buffer.Length; i++)
            {
                sBuffer += Convert.ToChar(buffer[i]);
                if (buffer[i] == 0)
                    return sBuffer;
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
    }
}
