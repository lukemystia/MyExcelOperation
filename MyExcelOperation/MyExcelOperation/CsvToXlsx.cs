using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;
using ClosedXML.Excel;
using System.IO;

namespace MyExcelOperation
{
	/// <summary>
	/// csv ファイルを xlsx ファイルへ変換する
	/// </summary>
	public static class CsvToXlsx
	{
		/// <summary>
		/// csv ファイルを xlsx ファイルへ変換する
		/// </summary>
		/// <param name="originCsvPath">csvファイルパス</param>
		/// <param name="outputXlsxPath">出力するxlsxファイルパス</param>
		public static void Output(string originCsvPath, string outputXlsxPath)
		{
			try
			{
				if (!File.Exists(originCsvPath))
				{
					throw new FileNotFoundException($"There is no file called {originCsvPath}.");
				}

				if (File.Exists(outputXlsxPath))
				{
					throw new FileNotFoundException($"A file called {outputXlsxPath} already exists.");
				}

				var datas = csvRead(originCsvPath);
				makeXlsx(outputXlsxPath, datas);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// csvファイル読み込み
		///	区切りをカンマに指定
		/// </summary>
		/// <param name="filePath"></param>
		private static List<List<string>> csvRead(string filePath)
		{
			var datas = new List<List<string>>();

			using (var parser = new TextFieldParser(filePath, Encoding.UTF8))
			{
				parser.TextFieldType = FieldType.Delimited;
				parser.SetDelimiters(",");

				parser.HasFieldsEnclosedInQuotes = true;
				parser.TrimWhiteSpace = false;

				while (!parser.EndOfData)
				{
					var temp = parser.ReadFields().ToList();

					datas.Add(temp);
				}
			}

			return datas;
		}

		/// <summary>
		/// xlsxファイル作成
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="csvData"></param>
		private static void makeXlsx(string filePath, List<List<string>> csvData)
		{
			using (var workbook = new XLWorkbook())
			using (var worksheet = workbook.Worksheets.Add("Sheet1"))
			{
				var workRow = 0;
				var workCell = worksheet.Cell(++workRow, 1);

				csvData.ForEach(lineData =>
				{
					lineData.ForEach(cellData =>
					{
						workCell.Value = cellData;
						workCell = workCell.CellRight();
					});

					workCell = worksheet.Cell(++workRow, 1);
				});

				// セルサイズ自動調整
				worksheet.ColumnsUsed().AdjustToContents();
				workbook.SaveAs(filePath);
			}
		}
	}
}
