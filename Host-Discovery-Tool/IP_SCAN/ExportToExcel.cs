using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace IP_SCAN
{
	class ExportToExcel
	{
		Excel.Application excelApp;
		Excel._Workbook wBook;
		Excel._Worksheet wSheet;
		Excel.Range wRange;



		public void NewExcel()
		{

			string time = DateTime.Now.ToString("yyyy_MM_dd_HH_mm");
			string pathFile = @"D:\C#\IP Scan data\test" + time;


			// 開啟一個新的應用程式
			excelApp = new Excel.Application();
			// 讓Excel文件可見
			excelApp.Visible = true;

			// 加入新的活頁簿
			excelApp.Workbooks.Add(Type.Missing);
			// 引用第一個活頁簿
			wBook = excelApp.Workbooks[1];
			// 設定活頁簿焦點
			wBook.Activate();

			// 引用第一個工作表
			wSheet = (Excel._Worksheet)wBook.Worksheets[1];
			// 命名工作表的名稱
			wSheet.Name = "工作表測試";
			// 設定工作表焦點
			wSheet.Activate();

			// 設定第1列資料
			excelApp.Cells[1, 1] = "IP Adress";
			excelApp.Cells[1, 2] = "Host";
			excelApp.Cells[1, 3] = "Status";
			excelApp.Cells[1, 4] = "OS version";
			excelApp.Cells[1, 7] = DateTime.Now.ToString("yyyy/MM/dd/HH:mm");

			// 設定第1列顏色
			wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 4]];
			wRange.Select();
			wRange.Font.Color = ColorTranslator.ToOle(Color.White);
			wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);





			try
			{
				//另存活頁簿
				wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
			}
			catch (Exception ex)
			{
				Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
			}




		}
		public void getExcelFile(object sender, EventArgs e, object filePath)
		{
			var Path = filePath.ToString();
			// 引用第一個工作表
			wSheet = (Excel._Worksheet)wBook.Worksheets[1];
			//Create COM Objects. Create a COM object for everything that is referenced
			Excel.Application xlApp = new Excel.Application();
			Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path);
			Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
			Excel.Range xlRange = xlWorksheet.UsedRange;

			int rowCount = xlRange.Rows.Count;
			int colCount = xlRange.Columns.Count;

			//iterate over the rows and columns and print to the console as it appears in the file
			//excel is not zero based!!
			for (int i = 1; i <= rowCount; i++)
			{
				for (int j = 1; j <= colCount; j++)
				{
					//new line
					if (j == 1)
					{
						
						Console.Write("\r\n");
					}
					//write the value to the console
					if (xlRange.Cells[i, j] != null)
					{
						//textBox2.Text += xlRange.Cells[i, j].Value + "\t";
						excelApp.Cells[i, j + 7] = xlRange.Cells[i, j].Value + "\t";
						//Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
					}
				}
			}

		}
	}
}
