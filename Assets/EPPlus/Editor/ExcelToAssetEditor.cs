using UnityEngine;
using UnityEditor;
using OfficeOpenXml;
using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections.Generic;

public class ExcelToAssetEditor
{
    static string testExcelPath;
    static string testAssetPath;

    [MenuItem("Tools/EPPlus/ExcelToAssets")]
    public static void ExcelConvert()
    {
        GetPaths(ref testExcelPath, ref testAssetPath);

        ExcelToAsset(testExcelPath, testAssetPath);
    }

    [MenuItem("Tools/EPPlus/ExcelToJson")]
    public static void ExcelConvertToJson()
    {
        GetPaths(ref testExcelPath, ref testAssetPath);

        ExcelToJson(testExcelPath, testAssetPath);
    }

    static void GetPaths(ref string excelPath, ref string assetPath)
    {
        testExcelPath = EditorUtility.OpenFilePanelWithFilters("Search Excel", Application.dataPath + "/EPPlus/TestExcels", new string[] { "Excel", "*" });
        if (string.IsNullOrEmpty(testExcelPath))
        {
            EditorUtility.DisplayDialog("Error", "Excel can't empty", "See");
            return;
        }

        testAssetPath = "Assets/EPPlus/Data";
        if (!Directory.Exists(testAssetPath))
            Directory.CreateDirectory(testAssetPath);
    }

    #region excel to asset by reflection

    static void ExcelToAsset(string excelPath, string assetPath)
    {
       using(FileStream excelStream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
       {
           List<List<object>> excelData = new List<List<object>>();
           using(ExcelPackage excelPackage = new ExcelPackage(excelStream))
           {
                List<string> typeLis = new List<string>();
                Type dataType;
                ExcelWorksheets workSheets = excelPackage.Workbook.Worksheets;

                // index based on 1 default
                for (int i = 1, len = workSheets.Count; i <= len; i++)
                {
                    ExcelWorksheet sheet = workSheets[i];
                    int max_row = sheet.Dimension.Rows;
                    int max_col = sheet.Dimension.Columns;

                    typeLis.Clear();
                    for(int col = 1; col <= max_col; col++)
                        typeLis.Add(sheet.GetValue<string>(2, col));

                    dataType = GetType(sheet.Name);
                    FieldInfo[] fieldInfo = dataType.GetFields();

                    List<object> tileLis = new List<object>();
                    object obj;
                    for(int row = 3; row <= max_row; row++)
                    {
                        var dataObj = dataType.Assembly.CreateInstance(sheet.Name);
                        for (int col = 1; col <= max_col; col++)
                        {
                            switch (typeLis[col - 1])
                            {
                                case "int" :
                                    obj = sheet.GetValue<int>(row, col);
                                break;
                                case "string":
                                default:
                                    obj = sheet.GetValue<string>(row, col);
                                break;
                            }
                            
                            fieldInfo[col - 1].SetValue(dataObj, obj);
                        }
                        tileLis.Add(dataObj);
                     }
                     excelData.Add(tileLis);
                }
            }

            SettingToAsset(excelData, assetPath);
        }
    }

    static Type GetType(string typeName)
    {
        Type type = null;
        Assembly curExecuteAssembly = Assembly.GetExecutingAssembly();
        AssemblyName[] refAssembly = curExecuteAssembly.GetReferencedAssemblies();
        foreach (var assemblyName in refAssembly)
        {
            var assembly = Assembly.Load(assemblyName);
            if(assemblyName != null)
            {
                type = assembly.GetType(typeName);
                if (type != null)
                    break;
            }
        }

        //typeof(SurfaceTile).Assembly.GetType()

        return type;
    }

    static void SettingToAsset(List<List<object>> data, string assetPath)
    {
        MapTileConfig mapConfig = ScriptableObject.CreateInstance<MapTileConfig>();

        for (int i = 0; i < data[0].Count; i++)
        {
            mapConfig.BasicTileData.Add((SurfaceTile)data[0][i]);
        }
        for (int i = 0; i < data[1].Count; i++)
        {
            mapConfig.BuildingData.Add((BuildingTile)data[1][i]);
        }

        AssetDatabase.CreateAsset(mapConfig, testAssetPath + "/MapTileConfig.asset");
        AssetDatabase.SaveAssets();
    }
    #endregion

    #region excel to  json
    static void ExcelToJson(string excelPath, string assetPath)
    {
        using (FileStream excelStream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
        {
            List<string> excelData = new List<string>();
            using (ExcelPackage excelPackage = new ExcelPackage(excelStream))
            {
                List<string> typeLis = new List<string>();
                List<string> nameLis = new List<string>();
                StringBuilder jsonData = new StringBuilder();
                ExcelWorksheets workSheets = excelPackage.Workbook.Worksheets;

                jsonData.Append("{");

                // index based on 1 default
                for (int i = 1, len = workSheets.Count; i <= len; i++)
                {
                    ExcelWorksheet sheet = workSheets[i];
                    int max_row = sheet.Dimension.Rows;
                    int max_col = sheet.Dimension.Columns;

                    nameLis.Clear();
                    typeLis.Clear();
                    for (int col = 1; col <= max_col; col++)
                        nameLis.Add(sheet.GetValue<string>(1, col));
                    for (int col = 1; col <= max_col; col++)
                        typeLis.Add(sheet.GetValue<string>(2, col));

                    jsonData.Append("\"");
                    jsonData.Append(sheet.Name);
                    jsonData.Append("\"");
                    jsonData.Append(":");
                    jsonData.Append("[");
                    for (int row = 3; row <= max_row; row++)
                    {
                        jsonData.Append("{");
                        for (int col = 1; col <= max_col; col++)
                        {
                            jsonData.Append("\"");
                            jsonData.Append(nameLis[col - 1]);
                            jsonData.Append("\"");
                            jsonData.Append(":");
                            switch (typeLis[col - 1])
                            {
                                case "int":
                                    jsonData.Append(sheet.GetValue<int>(row, col));
                                    break;
                                case "string":
                                default:
                                    jsonData.Append("\"");
                                    jsonData.Append(sheet.GetValue<string>(row, col));
                                    jsonData.Append("\"");
                                    break;
                            }
                            if(col != max_col)
                                jsonData.Append(",");
                        }
                        jsonData.Append("}");
                        if (row < max_row)
                            jsonData.Append(",");
                    }
                    jsonData.Append("]");
                    if (i < len)
                        jsonData.Append(",");
                }
                jsonData.Append("}");

                SaveToJson(jsonData.ToString(), assetPath);
            }
        }
    }

    static void SaveToJson(string jsonData, string assetPath)
    {
        FileStream dataStream = File.Open(assetPath + "/MapTileConfig.txt", FileMode.Create, FileAccess.Write);
        StreamWriter writer = new StreamWriter(dataStream);
        writer.Write(jsonData);

        writer.Close();
        dataStream.Close();

        AssetDatabase.Refresh();

        MapTileConfigForJson config = JsonUtility.FromJson<MapTileConfigForJson>(jsonData);
    }
    #endregion

}
