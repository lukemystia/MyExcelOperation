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
	public class OperationTests
	{
		[TestMethod()]
		public void GetSheetNamesTest()
		{
			var names = Operation.GetSheetNames(@"C:\Users\hogeuser\Desktop\Book1.xlsx");
			Assert.AreEqual(names.First(), "Sheet1");
		}

		[TestMethod()]
		public void GetCellValuesTest()
		{
			var cellVal = Operation.GetCellValues<string>(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "C5");
			Assert.AreEqual(cellVal, "S1");

			var cellVal2 = Operation.GetCellValues<string>(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet2", 5, 3);
			Assert.AreEqual(cellVal2, "S2");

			var cellVal3 = Operation.GetCellValues<string>(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet3", 5, "C");
			Assert.AreEqual(cellVal3, "S3");
		}

		[TestMethod()]
		public void WriteCellValuesTest()
		{
			Operation.WriteCellValues(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "test1", "D6");
			var cellVal = Operation.GetCellValues<string>(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "D6");
			Assert.AreEqual(cellVal, "test1");

			Operation.WriteCellValues(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet2", "test2", 6, 4);
			var cellVal2 = Operation.GetCellValues<string>(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet2", 6, 4);
			Assert.AreEqual(cellVal2, "test2");

			Operation.WriteCellValues(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet3", "test3", 6, "D");
			var cellVal3 = Operation.GetCellValues<string>(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet3", 6, "D");
			Assert.AreEqual(cellVal3, "test3");
		}

		[TestMethod()]
		public void GetTagAddressTest()
		{
			var tags = Operation.GetTagAddress(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "<");

			Assert.AreEqual("B9", tags["<tagB9>"]);
		}

		[TestMethod()]
		public void GetTagAddressErrTest()
		{
			var tags = Operation.GetTagAddress(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "<");

			try
			{
				var aaa = tags["<non>"];
			}
			catch (Exception)
			{
				return;
			}

			Assert.Fail("例外が発生しませんでした");
		}

		[TestMethod()]
		public void WriteRangeCellTest()
		{
			var writeDatas = new List<string>() { "1", "2", "3", "4", "5" }.ToArray();
			Operation.WriteRangeCell(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", writeDatas, "A5", "B6");
		}

		[TestMethod()]
		public void WriteCellValuesTest1()
		{
			Operation.WriteCellValues(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "a5", "A5");
			Operation.WriteCellValues(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "c5", "C5");
			var cellVal = Operation.GetCellValues<string>(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "C5");
			Assert.AreEqual(cellVal, "c5");
		}

		[TestMethod()]
		public void RangeCopyTest()
		{
			Operation.RangeCopy(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "A5:L5", "A9");
		}

		[TestMethod()]
		public void SetSeparatePrintAreasTest()
		{
			Operation.SetSeparatePrintAreas(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "A4:L5");
			Operation.SetSeparatePrintAreas(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "A6:L7");
			Operation.SetSeparatePrintAreas(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", 3);
			Operation.SetSeparatePrintAreas(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", 5);
		}

		[TestMethod()]
		public void CopySheetTest()
		{
			Operation.CopySheet(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "1");
		}

		[TestMethod()]
		public void MakeSheetTest()
		{
			Operation.MakeSheet(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "SheetTest");
		}

		[TestMethod()]
		public void MakeSheetErrorTest()
		{
			try
			{
				Operation.MakeSheet(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1");
			}
			catch (Exception)
			{
				return;
			}

			Assert.Fail("例外が発生しませんでした");
		}

		[TestMethod()]
		public void CopySheetTest1()
		{
			try
			{
				Operation.CopySheet(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "SheetTest");
			}
			catch (Exception)
			{
				return;
			}

			Assert.Fail("例外が発生しませんでした");
		}

		[TestMethod()]
		public void DeleteSheetTest()
		{
			Operation.DeleteSheet(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "1");
		}

		[TestMethod()]
		public void RowDeleteTest()
		{
			Operation.RowDelete(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", 14);
		}

		[TestMethod()]
		public void ColumnDeleteTest()
		{
			Operation.ColumnDelete(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", "P");
		}

		[TestMethod()]
		public void WriteCellValuesTest2()
		{
			var writeData = new List<WriteData>()
			{
				new WriteData("A12", "A12"),
				new WriteData("A13", "A", 13),
				new WriteData("A14", 1, 14),
				new WriteData("AA", "AA15"),
				new WriteData("AB", "AB", 16),
				new WriteData("AC", 29, 17),
			};

			Operation.WriteCellValues(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "Sheet1", writeData);
		}

		[TestMethod()]
		public void AddColumnAddressTest()
		{
			var aaa = Operation.AddColumnAddress("Z1", 1);
			Assert.AreEqual(aaa, "AA1");
		}

		[TestMethod()]
		public void CopySheetBetweenBooksTest2()
		{
			Operation.CopySheetBetweenBooks(@"C:\Users\hogeuser\Desktop\Book1.xlsx", "NewCopy", @"C:\Users\hogeuser\Desktop\Book2.xlsx", "copyOri");
		}
	}
}