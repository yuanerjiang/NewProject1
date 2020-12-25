///////////////////////////////////////////////////////////////////////////////
///
/// ScriptGenerator.cs
/// 
/// (c)2013 Kim, Hyoun Woo
///
///////////////////////////////////////////////////////////////////////////////
using System;
using UnityEngine;
using UnityEditor;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;

using Object = UnityEngine.Object;
using System.Text;
using static ExcelImportWindow;

public class WriteClass
    {
        /// <summary>
        /// 创建实体
        /// </summary>
        public static void CreateEntity(string filePath, string fileName, List<MemberFieldData>dataList)
        {
            if (dataList == null) return;

            if (!Directory.Exists( filePath))
            {
                Directory.CreateDirectory(filePath);
            }

            StringBuilder sbr = new StringBuilder();
            sbr.Append("\r\n");
            sbr.AppendFormat("//创建时间：{0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sbr.Append("//备    注：此代码为工具生成 请勿手工修改\r\n");
            sbr.Append("//===================================================\r\n");
            sbr.Append("using System.Collections;\r\n");
            sbr.Append("\r\n");
            sbr.Append("/// <summary>\r\n");
            sbr.AppendFormat("/// {0}实体\r\n", fileName);
            sbr.Append("/// </summary>\r\n");
            sbr.AppendFormat("public  class {0}Entity\r\n", fileName);
            sbr.Append("{\r\n");

            for (int i = 0; i < dataList.Count; i++)
            {
                sbr.Append("    /// <summary>\r\n");
                sbr.AppendFormat("    /// {0}\r\n", dataList[i].Type);
                sbr.Append("    /// </summary>\r\n");
                sbr.AppendFormat("    public {0} {1} {{ get; set; }}\r\n", dataList[i].Type,  string.IsNullOrEmpty(dataList[i].FieldName) ? dataList[i].Name: dataList[i].FieldName);
                sbr.Append("\r\n");
            }

            sbr.Append("}\r\n");


            using (FileStream fs = new FileStream(string.Format("{0}/{1}.cs", filePath, fileName), FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.Write(sbr.ToString());
                }
            }
        }


    /// <summary>
    /// 创建数据文件
    /// </summary>
    /// <param name="path"></param>
    /// <param name="fileName"></param>
    /// <param name="buffer"></param>
    public static void CreateData(string path,string fileName,byte[]buffer)
    {
        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }


        using (FileStream fs = new FileStream(string.Format("{0}/{1}.data", path, fileName), FileMode.Create))
        {
            fs.Write(buffer, 0, buffer.Length);


            fs.Close();
           
        }
    }

 }


