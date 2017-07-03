using Microsoft.VisualStudio.TestTools.UnitTesting;
using MyExcelOperation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelOperation.Tests
{
	[TestClass()]
	public class CsvToXlsxTests
	{
		/// <summary>
		/// ファイル置き場
		/// </summary>
		private static string csvPath => @"C:\Users\hogeuser\Desktop\";


		[TestMethod()]
		public void CsvReadTest()
		{
			var target = new PrivateType(typeof(CsvToXlsx));
			var result = target.InvokeStatic("csvRead", $"{csvPath}Book1.csv") as List<List<string>>;

			Assert.IsNotNull(result);
		}

		[TestMethod()]
		public void MakeXlsxTest()
		{
			var filePath = @"C:\Users\hogeuser\Desktop\Book1.xlsx";

			var target = new PrivateType(typeof(CsvToXlsx));
			target.InvokeStatic("makeXlsx", filePath, target.InvokeStatic("csvRead", $"{csvPath}Book1.csv") as List<List<string>>);

			Assert.IsTrue(System.IO.File.Exists(filePath));
		}

		[TestMethod()]
		public void CSVファイルがないテスト()
		{
			var oriPath = $"{csvPath}Book123.csv";
			var outPath = $"{csvPath}Book1.xlsx";

			try
			{
				CsvToXlsx.Output(oriPath, outPath);
			}
			catch (Exception)
			{
				return;
			}

			Assert.Fail("例外が発生しませんでした");
		}

		[TestMethod()]
		public void xlsxファイルがあるテスト()
		{
			var oriPath = $"{csvPath}Book123.csv";
			var outPath = $"{csvPath}Book1.xlsx";

			try
			{
				CsvToXlsx.Output(oriPath, outPath);
			}
			catch (Exception)
			{
				return;
			}

			Assert.Fail("例外が発生しませんでした");
		}

		[TestMethod()]
		public void 正常テスト()
		{
			var oriPath = $"{csvPath}Book1.csv";
			var outPath = $"{csvPath}Book1.xlsx";

			try
			{
				CsvToXlsx.Output(oriPath, outPath);
			}
			catch (Exception ex)
			{
				Assert.Fail(ex.Message);
			}
		}
	}
}