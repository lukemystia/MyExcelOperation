using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace MyExcelOperation
{
	/// <summary>
	/// xls ファイルを xlsx ファイルへ変換する
	/// </summary>
	public static class XlsToXlsx
	{
		/// <summary>
		/// xls ファイルを xlsx ファイルへ変換する
		/// </summary>
		/// <param name="xlsFilePath">xls ファイルパス</param>
		/// <param name="xlsxFilePath">変換した xlsx ファイルを出力するパス</param>
		public static void Convert(string xlsFilePath, string xlsxFilePath)
		{
			try
			{
				if (!File.Exists(xlsFilePath))
				{
					throw new FileNotFoundException($"There is no file called {xlsFilePath}.");
				}

				if (File.Exists(xlsxFilePath))
				{
					throw new FileNotFoundException($"A file called {xlsxFilePath} already exists.");
				}

				var excelApp = new Application();
				var eWorkbook = excelApp.Workbooks.Open(xlsFilePath);

				eWorkbook.SaveAs(xlsxFilePath,
					 XlFileFormat.xlOpenXMLWorkbook, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing,
					 XlSaveAsAccessMode.xlNoChange, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing, Type.Missing);

				eWorkbook.Close(false, Type.Missing, Type.Missing);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
