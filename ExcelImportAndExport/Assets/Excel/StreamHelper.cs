using System;
using System.IO;
using System.Text;



    public class StreamHelper : MemoryStream
    {
        public StreamHelper()
        {

        }

        public StreamHelper(byte[] buffer) : base(buffer)
        {

        }

        #region Short
        /// <summary>
        /// 从流中读取一个short数据
        /// </summary>
        /// <returns></returns>
        public short ReadShort()
        {
            byte[] arr = new byte[2];
            base.Read(arr, 0, 2);
            return BitConverter.ToInt16(arr, 0);
        }

        /// <summary>
        /// 把一个short数据写入流
        /// </summary>
        /// <param name="value"></param>
        public void WriteShort(short value)
        {
            byte[] arr = BitConverter.GetBytes(value);
            base.Write(arr, 0, arr.Length);
        }
        #endregion

        #region UShort
        /// <summary>
        /// 从流中读取一个ushort数据
        /// </summary>
        /// <returns></returns>
        public ushort ReadUShort()
        {
            byte[] arr = new byte[2];
            base.Read(arr, 0, 2);
            return BitConverter.ToUInt16(arr, 0);
        }

        /// <summary>
        /// 把一个ushort数据写入流
        /// </summary>
        /// <param name="value"></param>
        public void WriteUShort(ushort value)
        {
            byte[] arr = BitConverter.GetBytes(value);
            base.Write(arr, 0, arr.Length);
        }
        #endregion

        #region Int
        /// <summary>
        /// 从流中读取一个int数据
        /// </summary>
        /// <returns></returns>
        public int ReadInt()
        {
            byte[] arr = new byte[4];
            base.Read(arr, 0, 4);
            return BitConverter.ToInt32(arr, 0);
        }

        /// <summary>
        /// 把一个int数据写入流
        /// </summary>
        /// <param name="value"></param>
        public void WriteInt(int value)
        {
            byte[] arr = BitConverter.GetBytes(value);
            base.Write(arr, 0, arr.Length);
        }
        #endregion


        #region UTF8String
        /// <summary>
        /// 从流中读取一个sting数组
        /// </summary>
        /// <returns></returns>
        public string ReadUTF8String()
        {
            ushort len = this.ReadUShort();
            byte[] arr = new byte[len];
            base.Read(arr, 0, len);
            return Encoding.UTF8.GetString(arr);
        }

        /// <summary>
        /// 把一个string数据写入流
        /// </summary>
        /// <param name="str"></param>
        public void WriteUTF8String(string str)
        {
            byte[] arr = Encoding.UTF8.GetBytes(str);
            if (arr.Length > 65535)
            {
                throw new InvalidCastException("字符串超出范围");
            }
            WriteUShort((ushort)arr.Length);
            base.Write(arr, 0, arr.Length);
        }
        #endregion
    }