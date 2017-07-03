using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace MyExcelOperation
{
	/// <summary>
	/// 操作関連
	/// </summary>
	public static class Operation
	{
		#region シート操作
		/// <summary>
		/// シート名取得
		/// </summary>
		/// <param name="filePath">ファイルフルパス</param>
		/// <returns></returns>
		public static IEnumerable<string> GetSheetNames(string filePath)
		{
			if (!File.Exists(filePath))
			{
				throw new FileNotFoundException($"There is no file [{filePath}].");
			}

			var names = new List<string>();

			using (var book = new XLWorkbook(filePath))
			{
				var sheetNum = getSheetNum(filePath);
				for (int i = 1; i <= sheetNum; i++)
				{
					var sheet = book.Worksheet(i);
					names.Add(sheet.Name);
				}
			}

			return names.Select(s => s).AsEnumerable();
		}

		/// <summary>
		/// シート数を返す
		/// </summary>
		/// <param name="filePath">ファイルフルパス</param>
		/// <returns></returns>
		private static int getSheetNum(string filePath)
		{
			var xlApp = new Application();
			try
			{
				var xlBooks = xlApp.Workbooks;
				try
				{
					var xlBook = xlBooks.Open(filePath);
					try
					{
						var xlSheets = xlBook.Worksheets;
						try
						{
							var num = xlSheets.Count;

							return num;
						}
						finally
						{
							Marshal.ReleaseComObject(xlSheets);
						}
					}
					finally
					{
						if (xlBook != null)
						{
							xlBook.Close(false);
						}
						Marshal.ReleaseComObject(xlBook);
					}
				}
				finally
				{
					Marshal.ReleaseComObject(xlBooks);
				}
			}
			finally
			{
				if (xlApp != null)
				{
					xlApp.Quit();
				}
				Marshal.ReleaseComObject(xlApp);
			}
		}

		/// <summary>
		/// 改ページプレビュー範囲設定
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		/// <param name="printAreas">設定する範囲</param>
		/// <param name="horizontalPageBreak">2 と設定すると 2と3行目の間にページ区切り</param>
		public static void SetSeparatePrintAreas(string filePath, string sheetName, string printAreas, int horizontalPageBreak = 0)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				sheet.PageSetup.PrintAreas.Add(printAreas);

				if (horizontalPageBreak != 0)
				{
					sheet.PageSetup.AddHorizontalPageBreak(horizontalPageBreak);
				}

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}

		/// <summary>
		/// 改ページプレビュー範囲設定
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		/// <param name="horizontalPageBreak">2 と設定すると 2と3行目の間にページ区切り</param>
		public static void SetSeparatePrintAreas(string filePath, string sheetName, int horizontalPageBreak = 0)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				if (horizontalPageBreak != 0)
				{
					sheet.PageSetup.AddHorizontalPageBreak(horizontalPageBreak);
				}

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}

		/// <summary>
		/// シートを作成
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		public static void MakeSheet(string filePath, string sheetName)
		{
			using (var book = new XLWorkbook(filePath))
			{
				var sheetExist = book.Worksheets
					.Where(x => x.Name == sheetName)
					.Select(x => x.Worksheet)
					.Any();

				if (sheetExist)
				{
					throw new ArgumentException($"\"{sheetName}\" sheet already exists.");
				}

				book.AddWorksheet(sheetName);
				book.Save();
			}
		}

		/// <summary>
		/// シートをコピー
		///	コピー先シートが存在しなければ作られるので安心
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName">コピーするシート名</param>
		/// <param name="copySheetName">コピー先のシート名</param>
		public static void CopySheet(string filePath, string sheetName, string copySheetName)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				sheet.CopyTo(copySheetName);

				book.Save();
			}
		}

		/// <summary>
		/// シートの削除
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		public static void DeleteSheet(string filePath, string sheetName)
		{
			using (var book = new XLWorkbook(filePath))
			{
				book.Worksheets.Delete(sheetName);
				book.Save();
			}
		}

		/// <summary>
		/// 別ブックのシートをコピー
		///	コピー先シートが存在しなければ作られるので安心
		/// </summary>
		/// <param name="filePath">コピー先のブックパス</param>
		/// <param name="sheetName">コピー後のシート名</param>
		/// <param name="copyPath">コピー元のブックパス</param>
		/// <param name="copySheetName">コピー元のシート名</param>
		public static void CopySheetBetweenBooks(string filePath, string sheetName, string copyPath, string copySheetName)
		{
			using (var book = new XLWorkbook(filePath))
			using (var copyBook = new XLWorkbook(copyPath))
			{
				copyBook.Worksheet(copySheetName).CopyTo(book, sheetName);
				book.Save();
			}
		}
		#endregion シート操作

		#region セル値取得
		/// <summary>
		/// 指定セルの値取得
		/// </summary>
		/// <typeparam name = "T" >戻り値の型</typeparam>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="cellAddress">セルアドレス文字列(A1等)</param>
		/// <returns></returns>
		public static T GetCellValues<T>(string filePath, string sheetName, string cellAddress)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var cell = sheet.Cell(cellAddress);
				return cell.GetValue<T>();
			}
		}

		/// <summary>
		/// 指定セルの値取得
		/// </summary>
		/// <typeparam name = "T" >戻り値の型</typeparam>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="row"></param>
		/// <param name="column">A=1とした番号</param>
		/// <returns></returns>
		public static T GetCellValues<T>(string filePath, string sheetName, int row, int column)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var cell = sheet.Cell(row, column);
				return cell.GetValue<T>();
			}
		}

		/// <summary>
		/// 指定セルの値取得
		/// </summary>
		/// <typeparam name = "T" >戻り値の型</typeparam>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="row">行数字</param>
		/// <param name="column">列文字(A等)</param>
		/// <returns></returns>
		public static T GetCellValues<T>(string filePath, string sheetName, int row, string column)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var cell = sheet.Cell(row, column);
				return cell.GetValue<T>();
			}
		}

		/// <summary>
		/// タグアドレス取得
		/// </summary>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="tagFirstStr">タグの1文字目</param>
		/// <returns>key:タグ名 value:アドレスのDictionary</returns>
		public static Dictionary<string, string> GetTagAddress(string filePath, string sheetName, string tagFirstStr)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var cellDic = sheet.CellsUsed()
					.Where(x => x.Value.ToString().Substring(0, 1) == tagFirstStr)
					.Select(x => new { x.Value, x.Address })
					.ToDictionary(x => x.Value.ToString(), x => x.Address.ToString());

				return cellDic;
			}
		}
		#endregion セル値取得

		#region セル書込み
		#region 単体
		/// <summary>
		/// 指定セルに書込み
		/// </summary>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="writeValue">書き込む文字列</param>
		/// <param name="cellAddress">セルアドレス文字列(A1等)</param>
		/// <returns></returns>
		public static void WriteCellValues(string filePath, string sheetName, string writeValue, string cellAddress)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var cell = sheet.Cell(cellAddress);
				cell.Value = writeValue;

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}

		/// <summary>
		/// 指定セルに書込み
		/// </summary>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="writeValue">書き込む文字列</param>
		/// <param name="row"></param>
		/// <param name="column">A=1とした番号</param>
		/// <returns></returns>
		public static void WriteCellValues(string filePath, string sheetName, string writeValue, int row, int column)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var cell = sheet.Cell(row, column);
				cell.Value = writeValue;

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}

		/// <summary>
		/// 指定セルに書込み
		/// </summary>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="writeValue">書き込む文字列</param>
		/// <param name="row">行数字</param>
		/// <param name="column">列文字(A等)</param>
		/// <returns></returns>
		public static void WriteCellValues(string filePath, string sheetName, string writeValue, int row, string column)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var cell = sheet.Cell(row, column);
				cell.Value = writeValue;

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}
		#endregion 単体

		#region 複数
		/// <summary>
		/// 複数セルに書込み
		/// </summary>
		/// <param name="filePath">ファイルフルパス</param>
		/// <param name="sheetName">対象シート名</param>
		/// <param name="writeDatas">書き込むデータ群</param>
		/// <returns></returns>
		public static void WriteCellValues(string filePath, string sheetName, List<WriteData> writeDatas)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				foreach (var data in writeDatas)
				{
					var cell = sheet.Cell(data.Address);
					cell.Value = data.Value;
				}

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}
		#endregion 複数

		/// <summary>
		/// 範囲指定セルに書き込み
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		/// <param name="writeValues">書き込む文字列群</param>
		/// <param name="startCellAddress">開始セル</param>
		/// <param name="endCellAddress">終了セル</param>
		public static void WriteRangeCell(string filePath, string sheetName, string[] writeValues, string startCellAddress, string endCellAddress)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				sheet.Cell(startCellAddress).InsertData(new[] { writeValues });

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}

		/// <summary>
		/// 範囲コピー
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		/// <param name="range">コピー範囲(A1:C3 のような書き方)</param>
		/// <param name="copyAddress">コピー位置(A4 のような書き方)</param>
		public static void RangeCopy(string filePath, string sheetName, string range, string copyAddress)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				var copyRange = sheet.Range(range);
				var startCell = sheet.Cell(copyAddress);

				copyRange.CopyTo(startCell);

				sheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				book.Save();
			}
		}
		#endregion セル書込み

		#region 行・列
		/// <summary>
		/// 行の削除
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		/// <param name="deleteRow">削除する行の番号</param>
		public static void RowDelete(string filePath, string sheetName, int deleteRow)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				sheet.Row(deleteRow).Delete();

				book.Save();
			}
		}

		/// <summary>
		/// 列の削除
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sheetName"></param>
		/// <param name="deleteColumn">削除する列の文字</param>
		public static void ColumnDelete(string filePath, string sheetName, string deleteColumn)
		{
			using (var book = new XLWorkbook(filePath))
			using (var sheet = book.Worksheet(sheetName))
			{
				sheet.Column(deleteColumn).Delete();

				book.Save();
			}
		}

		/// <summary>
		/// アドレス文字行数をずらしたアドレス文字を取得
		/// </summary>
		/// <param name="address"></param>
		/// <param name="addNum"></param>
		/// <returns></returns>
		public static string AddRowAddress(this string address, int addNum)
		{
			var lineNum = int.Parse(new Regex(@"[^0-9]").Replace(address, ""));
			var rowStr = address.Substring(0, address.Length - lineNum.ToString().Length);

			return $"{rowStr}{lineNum + addNum}";
		}

		/// <summary>
		/// アドレス文字列数をずらしたアドレス文字を取得
		/// </summary>
		/// <param name="address"></param>
		/// <param name="addNum"></param>
		/// <returns></returns>
		public static string AddColumnAddress(this string address, int addNum)
		{
			var lineNum = int.Parse(new Regex(@"[^0-9]").Replace(address, ""));
			var rowStr = address.Substring(0, address.Length - lineNum.ToString().Length);
			var rowNum = rowStr.ToInt() + addNum;

			return $"{rowNum.ToAlphabet()}{lineNum}";
		}

		/// <summary>
		/// 数値をExcelのカラム名的なアルファベット文字列へ変換します
		/// </summary>
		/// <param name="self"></param>
		/// <returns>
		/// Excelのカラム名的なアルファベット文字列
		/// (変換できない場合は、空文字を返します)
		/// </returns>
		private static string ToAlphabet(this int self)
		{
			if (self <= 0)
			{
				return "";
			}

			var n = self % 26;
			n = (n == 0) ? 26 : n;

			var s = ((char)(n + 64)).ToString();

			if (self == n)
			{
				return s;
			}

			return ((self - n) / 26).ToAlphabet() + s;
		}

		/// <summary>
		/// Excelのカラム名的なアルファベットを数値へ変換します
		/// </summary>
		/// <param name="self"></param>
		/// <returns>
		/// 数値
		/// (変換できない場合は、0を返します)
		/// </returns>
		private static int ToInt(this string self)
		{
			var result = 0;
			if (string.IsNullOrEmpty(self))
			{
				return result;
			}

			var chars = self.ToCharArray();
			var len = self.Length - 1;
			foreach (var c in chars)
			{
				var asc = (int)c - 64;
				if (asc < 1 || asc > 26)
				{
					return 0;
				}

				result += asc * (int)Math.Pow((double)26, (double)len--);
			}
			return result;
		}
		#endregion 行・列
	}

	/// <summary>
	/// 複数のcellに書き込みたい時に使用するデータモデル
	/// </summary>
	public class WriteData
	{
		/// <summary>
		/// 書込むアドレス
		/// </summary>
		public string Address { get; }

		/// <summary>
		/// 書き込むデータ
		/// </summary>
		public string Value { get; }


		/// <summary>
		/// コンストラクタ
		/// </summary>
		/// <param name="value">書き込むデータ</param>
		/// <param name="address">書込むアドレス</param>
		public WriteData(string value, string address)
		{
			this.Value = value;
			this.Address = address;
		}

		/// <summary>
		/// コンストラクタ
		/// </summary>
		/// <param name="value">書き込むデータ</param>
		/// <param name="columnAddress">書込む列アドレス文字</param>
		/// <param name="rowAddress">書込む行アドレス</param>
		public WriteData(string value, string columnAddress, int rowAddress)
		{
			this.Value = value;
			this.Address = $"{columnAddress}{rowAddress.ToString()}";
		}

		/// <summary>
		/// コンストラクタ
		/// </summary>
		/// <param name="value">書き込むデータ</param>
		/// <param name="columnAddress">書込む列アドレス</param>
		/// <param name="rowAddress">書込む行アドレス</param>
		public WriteData(string value, int columnAddress, int rowAddress)
		{
			this.Value = value;
			this.Address = $"{ToAlphabet(columnAddress)}{rowAddress.ToString()}";
		}

		/// <summary>
		/// 数値をExcelのカラム名的なアルファベット文字列へ変換します
		/// </summary>
		/// <param name="self"></param>
		/// <returns>
		/// Excelのカラム名的なアルファベット文字列
		/// (変換できない場合は、空文字を返します)
		/// </returns>
		private string ToAlphabet(int self)
		{
			if (self <= 0)
			{
				return "";
			}

			var n = self % 26;
			n = (n == 0) ? 26 : n;

			var s = ((char)(n + 64)).ToString();

			if (self == n)
			{
				return s;
			}

			return ToAlphabet(((self - n) / 26)) + s;
		}
	}
}
