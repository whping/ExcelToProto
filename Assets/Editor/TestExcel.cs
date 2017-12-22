using UnityEngine;
using System.Collections;
using UnityEditor;
using System.IO;
using ExcelDataReader;
using System.Text;

public class TestExcel
{
	[MenuItem("Excel/Test")]
	public static void ReadExcel()
	{
		string path = Application.dataPath + "/movement.xlsx";
		using (var stream = File.Open(path, FileMode.Open, FileAccess.Read)) 
		{
			// Auto-detect format, supports:
			//  - Binary Excel files (2.0-2003 format; *.xls)
			//  - OpenXml Excel files (2007 format; *.xlsx)
			using (var reader = ExcelReaderFactory.CreateReader(stream)) 
			{

				// Choose one of either 1 or 2:

				// 1. Use the reader methods
				do {
					while (reader.Read()) {
						// reader.GetDouble(0);
						StringBuilder sb = new StringBuilder();
						for (int i = 0; i < reader.FieldCount; ++i)
						{
							sb.Append(reader.GetValue(i));
							sb.Append(",");
						}
						Debug.Log(sb.ToString());
					}
					return;
				} while (reader.NextResult());

				// 2. Use the AsDataSet extension method
				//var result = reader.AsDataSet();
				//foreach (var table in result.Tables) {
					
				//}

				// The result of each spreadsheet is in result.Tables
			}
		}
	}
}
