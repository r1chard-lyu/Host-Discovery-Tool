using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace IP_SCAN
{

	class compare
	{

		Excel.Application excelApp;
		Excel._Workbook wBook;
		Excel._Worksheet wSheet;
		Excel.Range wRange;
		static string[,] a = new string[1000, 1000];
		static string[,] b = new string[1000, 1000];

		public compare()
		{
			open1();
			open2();
			con();

		}

		private void open1()
		{
			
			var fileContent = string.Empty;
			var filePath1 = string.Empty;
			var filePath2 = string.Empty;

			using (OpenFileDialog openFileDialog = new OpenFileDialog())
			{
				openFileDialog.InitialDirectory = "D:\\";
				openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
				openFileDialog.FilterIndex = 2;
				//openFileDialog.RestoreDirectory = true;
				openFileDialog.Multiselect = true;

				try
				{
					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						//Get the path of specified file
						filePath1 = openFileDialog.FileName;

						//Read the contents of the file into a stream

						foreach (String file in openFileDialog.FileNames)
						{
							var fileStream = openFileDialog.OpenFile();

							using (StreamReader reader = new StreamReader(fileStream))
							{
								fileContent = reader.ReadToEnd();
							}
						}

					}
					getExcelFile1(filePath1);
				}
				catch (Exception e)
				{
					MessageBox.Show(e.ToString() + "\n");
				}
			}

		}


		private void NewExcel()
		{
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
			//excelApp.Cells[1, 1] = "IP Adress";
			//excelApp.Cells[1, 2] = "Host";
			//excelApp.Cells[1, 3] = "Status";
			//excelApp.Cells[1, 4] = "OS version";
			//excelApp.Cells[1, 7] = DateTime.Now.ToString("yyyy/MM/dd/HH:mm");

			// 設定第1列顏色
			//wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 4]];
			//wRange.Select();
			//wRange.Font.Color = ColorTranslator.ToOle(Color.White);
			//Range.Interior.Color = ColorTranslator.ToOle(Color.DimGray);

		}

		private void getExcelFile1(object filePath)
		{
			NewExcel();
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

					//write the value to the console
					if (xlRange.Cells[i, j] != null)
					{
						//textBox2.Text += xlRange.Cells[i, j].Value + "\t";
						//excelApp.Cells[i, j] = xlRange.Cells[i, j].Value + "\t";
						a[i, j] = xlRange.Cells[i, j].Value + "\t";
						excelApp.Cells[i, j] = a[i, j];

						//Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

					}
					excelApp.Cells[1, 7] = "";
				}
				if (i != 1)
					excelApp.Cells[i, 14] = " ";
			}

		}

		/// <summary>
		/// ///////////////////////////////////////
		/// </summary>

		private void open2()
		{
			var fileContent = string.Empty;
			var filePath1 = string.Empty;
			var filePath2 = string.Empty;

			using (OpenFileDialog openFileDialog = new OpenFileDialog())
			{
				openFileDialog.InitialDirectory = "D:\\";
				openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
				openFileDialog.FilterIndex = 2;
				openFileDialog.RestoreDirectory = true;
				openFileDialog.Multiselect = true;

				try
				{
					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						//Get the path of specified file
						filePath1 = openFileDialog.FileName;

						//Read the contents of the file into a stream

						foreach (String file in openFileDialog.FileNames)
						{
							var fileStream = openFileDialog.OpenFile();

							using (StreamReader reader = new StreamReader(fileStream))
							{
								fileContent = reader.ReadToEnd();
							}
						}

					}
					getExcelFile2(filePath1);
				}
				catch (Exception e)
				{
					MessageBox.Show(e.ToString() + "\n");
				}
			}

		}
		private void getExcelFile2(object filePath)
		{
			NewExcel();
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

					//write the value to the console
					if (xlRange.Cells[i, j] != null)
					{
						//textBox2.Text += xlRange.Cells[i, j].Value + "\t";
						//excelApp.Cells[i, j] = xlRange.Cells[i, j].Value + "\t";
						b[i, j] = xlRange.Cells[i, j].Value + "\t";
						excelApp.Cells[i, j] = b[i, j];

						//Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

					}
					excelApp.Cells[1, 7] = "";                                                                                                                                                                                                       
				}
				if (i != 1)
					excelApp.Cells[i, 14] = " ";
			}

		}




		//////
		///compare
		private void con()
		{

			Excel.Application excelApp;
			Excel._Workbook wBook;
			Excel._Worksheet wSheet;
			Excel.Range wRange;

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


			for (int i = 1; i <= a.Length - 1; i++)
			{
				for (int j = 1; j <= b.Length - 1; j++)
				{
					if (a[i, 0] == b[j, 0])
					{
						if (a[i, 1] != b[j, 1])
						{
							excelApp.Cells[i, 0] = b[j, 0].ToString();
							excelApp.Cells[i, 1] = b[j, 1].ToString();

						}
					}
					else
					{


					}
					
				}
			}











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

	}
}


























