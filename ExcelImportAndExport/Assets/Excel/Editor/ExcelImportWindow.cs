using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using SFB;
using System;
using System.IO;
using UnityQuickSheet;

public class ExcelImportWindow : EditorWindow
{

    [MenuItem("Tools/Excel/ImportExcel")]
    public static void GetWindow()
    {
        ExcelImportWindow window = EditorWindow.GetWindow<ExcelImportWindow>("导入excel");

        window.Show();
    }


    /// <summary>
    /// 所选文件名字
    /// </summary>
    string fileArray = "";


    /// <summary>
    /// 所选的文件名数组
    /// </summary>
    string[] fileNames = null;


    /// <summary>
    /// excel工具类
    /// </summary>
    ExcelTool tools;

    /// <summary>
    /// excel表头数据配置类
    /// </summary>
    private List<ColumnHeader> ColumnHeaderList = new List<ColumnHeader>();


    /// <summary>
    /// 生成的数据类名
    /// </summary>
    private string className = string.Empty;

    private void OnGUI()
    {
        EditorGUILayout.Space(10);
        if (GUILayout.Button("选择路径",GUILayout.Width(200)))
        {
            fileNames = StandaloneFileBrowser.OpenFilePanel("", "", "*", false);

            for (int i = 0; i < fileNames.Length; i++)
            {
                fileArray += fileNames[i] + "\n";
            }
        }

        EditorGUILayout.Space(10);

        EditorGUILayout.BeginVertical();

        EditorGUILayout.TextField("选择的excel文件路径为:");

        fileArray =  EditorGUILayout.TextArea(fileArray, GUILayout.MaxHeight(100));

        EditorGUILayout.EndVertical();


        EditorGUILayout.Space(5);
        EditorGUILayout.BeginHorizontal();

        if (GUILayout.Button("清空导入",GUILayout.Width(100)))
        {
            ClearImport();
        }
        EditorGUILayout.Space();

        if (GUILayout.Button("导入数据",GUILayout.Width(100)))
        {
            Import();
        }
        EditorGUILayout.EndHorizontal();

        EditorGUILayout.Space(5);

        DrawExcelTitle();

    }

    private void DrawExcelTitle()
    {
        if (fileArray == null || fileArray.Length == 0)
        {
            return;
        }

        if (tools==null||tools.titles==null)
        {
            return;
        }

        GUILayout.Label("数据类型设置:");
        const int MEMBER_WIDTH = 110;

        using (new GUILayout.HorizontalScope(EditorStyles.toolbar))
        {
            GUILayout.Label("成员变量", GUILayout.Width(MEMBER_WIDTH));
            GUILayout.FlexibleSpace();
            string[] names = { "变量名", "类型" };
            int[] widths = { 110, 80 };
            for (int i = 0; i < names.Length; i++)
            {
                GUILayout.Label(new GUIContent(names[i]), GUILayout.Width(widths[i]));
            }
        }

        using (new EditorGUILayout.VerticalScope("box"))
        {
            foreach (ColumnHeader header in ColumnHeaderList)
            {
                GUILayout.BeginHorizontal();

                // show member field with label, read-only
                EditorGUILayout.LabelField(header.name, GUILayout.Width(MEMBER_WIDTH));
                GUILayout.FlexibleSpace();

                header.fieldName = EditorGUILayout.TextField(header.fieldName, GUILayout.Width(110));

                // specify type with enum-popup
                header.type = (CellType)EditorGUILayout.EnumPopup(header.type, GUILayout.Width(80));
                GUILayout.Space(30);

                GUILayout.EndHorizontal();

            }

            GUILayout.BeginHorizontal();

            EditorGUILayout.LabelField("生成的数据类名:", GUILayout.Width(120));
            className = EditorGUILayout.TextField(className, GUILayout.Width(120));

            GUILayout.EndHorizontal();
        }


    }


    void DrawGenerateBtn()
    {
        EditorGUILayout.Space(5);
        GUILayout.BeginHorizontal();

        if (GUILayout.Button("生成数据类", GUILayout.Width(200)))
        {
            if (string.IsNullOrEmpty(className))
            {
                className = "DataSheet";
            }
            CreateDataClassScript(className);

            if (tools.buffer == null)
            {
                EditorUtility.DisplayDialog("Error", "数据为空", "OK");
            }

            CreateDataFile(className, tools.buffer);

            AssetDatabase.Refresh();
        }
        GUILayout.EndHorizontal();
    }


    /// <summary>
    /// 清空导入
    /// </summary>
    void ClearImport()
    {
        fileArray = "";
        fileNames = null;
        tools = null;
        ColumnHeaderList.Clear();
        className = string.Empty;
    }


    /// <summary>
    /// 导入EXCEL数据并生成对应数据文件
    /// </summary>
    void Import()
    {
        if (string.IsNullOrEmpty(fileArray))
        {
            return;
        }

        string[] fileItems = fileArray.TrimEnd().Split('\n');

        for (int i = 0; i < fileItems.Length; i++)
        {
            if (!fileItems[i].EndsWith(".xlsx"))
            {
                continue;
            }

            tools= new ExcelTool();
            tools.ImportExcel(fileItems[i]);
        }

        for (int i = 0; i < tools.titles.Length; i++)
        {
            ColumnHeader header = new ColumnHeader();
            header.name = tools.titles[i];
            header.type = CellType.Undefined;

            ColumnHeaderList.Add(header);
        }
    }


    /// <summary>
    /// 生成数据类
    /// </summary>
    /// <param name="fileName"></param>
     void CreateDataClassScript(string fileName)
    {
        string filePath = Application.dataPath + "/Entity";

        List<MemberFieldData> fieldList = new List<MemberFieldData>();


        foreach (ColumnHeader header in ColumnHeaderList)
        {
            MemberFieldData member = new MemberFieldData();
            member.Name = header.name;
            member.type = header.type;
            member.FieldName = header.fieldName;

            fieldList.Add(member);
        }

        WriteClass.CreateEntity(filePath, fileName, fieldList);
    }


    /// <summary>
    /// 生成数据文件
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="buffer"></param>
    void CreateDataFile(string fileName,byte[]buffer)
    {
        string filePath = Application.dataPath + "/Data";

        WriteClass.CreateData(filePath, fileName, buffer);
    }


    [System.Serializable]
    public class ColumnHeader
    {
        public CellType type;
        public string name;
        public string fieldName;
        public bool isEnable;
    }

    public enum CellType
    {
        Undefined,
        String,
        Short,
        Int,
        Long,
        Float,
        Double,
        Enum,
        Bool,
    }

    public class MemberFieldData
    {
        public CellType type = CellType.Undefined;
        private string name;
        private string fileName;

        public string FieldName
        {
            get { return fileName; }
            set { fileName = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Type
        {
            get
            {
                switch (type)
                {
                    case CellType.String:
                        return "string";
                    case CellType.Short:
                        return "short";
                    case CellType.Int:
                        return "int";
                    case CellType.Long:
                        return "long";
                    case CellType.Float:
                        return "float";
                    case CellType.Double:
                        return "double";
                    case CellType.Enum:
                        return "enum";
                    case CellType.Bool:
                        return "bool";
                    default:
                        return "string";
                }
            }
        }

        public MemberFieldData()
        {
            name = "";
            type = CellType.Undefined;
        }

    }

}
