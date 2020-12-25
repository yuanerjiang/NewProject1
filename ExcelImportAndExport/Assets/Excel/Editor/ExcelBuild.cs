using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using UnityEditor;
using UnityEngine;

public class ExcelBuild : Editor
{

    [MenuItem("Tools/Excel/Example/CreateItemAsset")]
    public static void CreateItemAsset()
    {
        string dataPath = Application.dataPath + "/Data/Space.data";


        if (File.Exists(dataPath))
        {
            byte[] buffer = File.ReadAllBytes(dataPath);

            using (StreamHelper stream = new StreamHelper(buffer))
            {
                int row = stream.ReadInt();
                int column = stream.ReadInt();

                string[,] data = new string[row, column];


                int index = 0;
                for (int i = 0; i < row; i++)
                {
                    SpaceEntity space = new SpaceEntity();


                    for (int k =0; k < column; k++)
                    {
                        data[i, k] = stream.ReadUTF8String();
                    }
                }

                
                for (int i = 0; i < row; i++)
                {
                    SpaceEntity space = new SpaceEntity();
                    space.id = data[i, index++];
                    space.name = data[i, index++];
                    space.kongjian = data[i, index++];

                    index = 0;

                    Debug.Log(space.id + "=====" + space.name + "======" + space.kongjian);
                }
            }
        }
      

    }

}
