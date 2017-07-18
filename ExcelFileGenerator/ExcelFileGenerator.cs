using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelFileGenerator
{
	public class ExcelFileGenerator
	{
		public static void Main(string[] args)
		{
			CreateTable();
			CreateFile();
			Console.WriteLine();
		}

		public static System.Data.DataTable CreateTable()
		{
			List<string> namesList = new List<string> { "Pavel", "Petar", "Tony", "Bob", "Joe", "Tod", "Ivan", "Pablo", "Kevin", "Kro" };
			string name = string.Empty;
			byte age = 0;
			byte score = 0;
			byte averageScore = 0;
			string formula = "C2:C101";
			Random rnd = new Random();

			System.Data.DataTable table = new System.Data.DataTable();
			table.Columns.Add("Name", typeof(string));
			table.Columns.Add("age", typeof(byte));
			table.Columns.Add("score", typeof(byte));
			table.Columns.Add("averageScore", typeof(byte));
			table.Columns.Add("formula", typeof(string));

			//TODO make it work with 100k records
			for (int i = 1; i < 100; i++)
			{
				byte rndName = (byte)rnd.Next(namesList.Count);
				name = (string)namesList[rndName];
				age = (byte)rnd.Next(20, 81);
				score = (byte)rnd.Next(0, 101);
				averageScore = (byte)(age + score);
				table.Rows.Add(name, age, score, averageScore, formula);
			}
			return table;
		}

		public static void CreateFile()
		{
			Application excelApp = null;
			Workbook workbook = null;
			Worksheet sheet = null;
			Range range = null;

			try
			{
				string file = Path.Combine(Environment.CurrentDirectory, @"..\..\scores.xlsx");
				excelApp = new Application();
				workbook = excelApp.Workbooks.Add();
				sheet = (Worksheet)workbook.Sheets[1];
				sheet.Name = "scores";

			
				sheet.Cells[1, 1] = "Name";
				sheet.Cells[1, 2] = "Age";
				sheet.Cells[1, 3] = "Score";
				sheet.Cells[1, 4] = "AverageScore";
				sheet.Cells[1, 5] = "Formula";

				range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 5]];
				range.Interior.Color = XlRgbColor.rgbSkyBlue;
				range.Font.Bold = true;

				byte rowCounter = 2;
				foreach (DataRow datarow in CreateTable().Rows)
				{
					for (int i = 0; i < CreateTable().Columns.Count; i++)
					{
						sheet.Cells[rowCounter, i + 1] = datarow.ItemArray[i];
						if (rowCounter % 2 == 0)
						{
							range = sheet.Range[sheet.Cells[rowCounter, 1], sheet.Cells[rowCounter, 5]];
							range.Font.Color = XlRgbColor.rgbGreen;
						}
					}
					rowCounter += 1;
				}

				workbook.SaveAs(file);
				workbook.Close();
				excelApp.Quit();
			}
			catch (Exception exception)
			{
				Console.Write(exception);
			}
		}
	}
}
