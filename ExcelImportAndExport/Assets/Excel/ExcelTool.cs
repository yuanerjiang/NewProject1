using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using UnityEngine;

public class ExcelTool 
{

    /// <summary>
    /// 导入excel的数据的总行列
    /// </summary>
     int row, column;


    /// <summary>
    /// 获取excel表头
    /// </summary>
    public  string[] titles = null;


    /// <summary>
    /// 数据流
    /// </summary>
    public byte[] buffer = null;



    #region 导出excel
    /// <summary>
    /// 导出excel表到指定路径(要求数据类都为字段)
    /// </summary>
    /// <typeparam name="T">数据类型</typeparam>
    /// <param name="path">保存路径</param>
    /// <param name="t">数据类</param>
    /// <param name="rowName">excel表头内容</param>
    public  void CreateXLSX<T>(string path, List<T> t, string[] rowName = null) where T : class
    {
        //需要写入多少行数据
        int count = t.Count;

        //创建文件
        FileInfo info = new FileInfo(path);

        if (info.Exists)
        {
            info.Delete();
            info = new FileInfo(path);
        }

        //写入内容
        using (ExcelPackage package = new ExcelPackage(info))
        {
            //先创建一个工作表
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");

            if (rowName != null && rowName.Length > 0)
            {
                //写入表头
                for (int i = 1; i < rowName.Length + 1; i++)
                {
                    sheet.Cells[1, i].Value = rowName[i - 1];
                }

                //写入具体内容  
                for (int k = 2; k < count + 2; k++)
                {
                    //获取到数据类内的所有字段
                    Type type = t[k - 2].GetType();
                    FieldInfo[] infos = type.GetFields();

                    for (int i = 1; i < infos.Length + 1; i++)
                    {
                        sheet.Cells[k, i].Value = infos[i - 1].GetValue(t[k - 2]).ToString();
                    }
                }
            }
            package.Save();
        }

    }


    /// <summary>
    /// 导出excel表到指定路径
    /// </summary>
    /// <typeparam name="T">数据内容类</typeparam>
    /// <typeparam name="K">表头类</typeparam>
    /// <param name="path"></param>
    /// <param name="t"></param>
    /// <param name="k"></param>
    public  void CreateXLSX<T, K>(string path, List<T> t, K k) where T : class where K : class
    {

        //需要写入多少行数据
        int count = t.Count;
        //创建文件
        FileInfo info = new FileInfo(path);

        if (info.Exists)
        {
            info.Delete();
            info = new FileInfo(path);
        }

        //写入内容
        using (ExcelPackage package = new ExcelPackage(info))
        {
            //先创建一个工作表
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");

            FieldInfo[] fields = k.GetType().GetFields();

            //写入表头
            for (int i = 1; i < fields.Length + 1; i++)
            {
                sheet.Cells[1, i].Value = fields[i - 1];
            }

            //写入具体内容  
            for (int j = 2; j < count + 2; j++)
            {
                //获取到数据类内的所有字段
                Type type = t[j - 2].GetType();
                FieldInfo[] infos = type.GetFields();

                for (int i = 1; i < infos.Length + 1; i++)
                {
                    sheet.Cells[j, i].Value = infos[i - 1].GetValue(t[j - 2]).ToString();
                }
            }
            package.Save();
        }
    }

    #endregion



    /// <summary>
    /// 获取导入的excel
    /// </summary>
    /// <param name="path"></param>
    /// <param name="sheetName"></param>
    public  void  ImportExcel(string path)
    {
        using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            using (ExcelPackage package = new ExcelPackage(fileStream))
            {
                ExcelWorksheets workSheets = package.Workbook.Worksheets;

                ExcelWorksheet workSheet = workSheets[1];

                ExcelAddressBase address = workSheet.Dimension;

                row = address.Rows;
                column = address.Columns;


                List<string> titleList = new List<string>();

                titles = new string[column];


                //读取excel的第一行
                for (int i = 0; i < column; i++)
                {
                    if (workSheet.Cells[1, i + 1].Value == null)
                    {
                        continue;
                    }
                    titleList.Add(titles[i]);
                }

                titles = titleList.ToArray();

                for (int i = 0; i < titles.Length; i++)
                {
                    titles[i] = workSheet.Cells[1, i + 1].Value.ToString();
                }


                //读取excel中的数据
                buffer = GetBytes(workSheet.Cells);

                Debug.Log(buffer.Length);

            }
        }
    }



    /// <summary>
    /// 把excel中的数据转换成字节保存
    /// </summary>
    /// <param name="cells"></param>
    /// <returns></returns>
    private byte[] GetBytes(ExcelRange cells)
    {
        using (StreamHelper stream = new StreamHelper())
        {
            stream.WriteInt(row-1);
            stream.WriteInt(titles.Length);

            for (int i = 0; i < row-1; i++)
            {
                for (int k = 0; k < titles.Length; k++)
                {
                    if (cells[i+2,k+1].Value == null)
                    {
                        stream.WriteUTF8String("null");
                        continue;
                    }

                    stream.WriteUTF8String(cells[i+2, k+1].Value.ToString());
                }
            }

            return stream.ToArray();
        }
    }

}
