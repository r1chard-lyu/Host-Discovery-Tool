using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Threading;

namespace IP_SCAN
{
	class IpScan_Async
	{
		public static TextBox _textbox, _textbox2, _textbox3;

		private int timeout = 100;
		private int nFound = 0;

		static object x = new object();
		Stopwatch stopWatch = new Stopwatch();
		TimeSpan ts;


		List<IPdata> pdatas = new List<IPdata>();


		//IP data store struct
		struct IPdata
		{
			public int IPbin;
			public IPHostEntry host;
			public string status;
			
		}


		//constructor1 ()
		public IpScan_Async()
		{

			string StartIP = _textbox2.Text;
			string StopIP = _textbox3.Text;


			try
			{
				int IntStartIP = IpToInt(StartIP);
				int IntStopIP = IpToInt(StopIP);
				int n = IntStopIP - IntStartIP;
				if (n >= 0)
				{
					RunPingSweep_Async(StartIP, StopIP);


				}
				else
				{
					MessageBox.Show("IP終點位置不能小於IP起始位置");
				}

			}
			catch
			{

				MessageBox.Show("IP格式錯誤");
			}

		}


		//RunPingSweep_Async method1
		public async void RunPingSweep_Async(string StartIP, string StopIP)
		{
			nFound = 0;

			var tasks = new List<Task>();

			stopWatch.Start();
			int n = IpToInt(StopIP) - IpToInt(StartIP);


			//_textbox.AppendText("迴圈\n");
			for (int i = 0; i <= n; i++)
			{
				string ip = IntToIp(IpToInt(StartIP) + i);
				Ping p = new Ping();

				try
				{
					var task = PingAndUpdateAsync(p, ip, (i + 1));
					tasks.Add(task);
				}
				catch (Exception e)
				{
					_textbox.AppendText("\n exception:" + e + " , ");
				}
			}

			await Task.WhenAll(tasks).ContinueWith(t =>
			{
				stopWatch.Stop();
				//SORT pdatas compare to IPbin
				pdatas.Sort((s1, s2) => s1.IPbin.CompareTo(s2.IPbin));
				//_textbox.AppendText("排序中...\n");
				OpenExcel();
				

			});



			
			stopWatch.Stop();
			ts = stopWatch.Elapsed;



			//Print pdatas
			for (int j = 0; j <= pdatas.Count - 1; j++)
			{
				try
				{
					_textbox.AppendText(pdatas[j].host.HostName.ToString() + "   " + IntToIp(pdatas[j].IPbin) + " is " + pdatas[j].status + "\n");
				}
				catch
				{
					_textbox.AppendText(IntToIp(pdatas[j].IPbin) + "  ,   " + "   " + IntToIp(pdatas[j].IPbin) + " is " + pdatas[j].status + "\n");
				}
			}

			_textbox.AppendText("Finished\n");

			MessageBox.Show(nFound.ToString() + " devices found! Elapsed time: " + ts.ToString(), "Asynchronous");

		}

		private async Task PingAndUpdateAsync(Ping ping, string ip, int i)
		{

			var reply = await ping.SendPingAsync(ip, timeout);

			IPdata x = new IPdata();


			Thread thread1 = new Thread(()=> RunScanTcp(ip));
			thread1.Start();
			

			if (reply.Status == IPStatus.Success|| a )
			{
				a = false;
				x.IPbin = IpToInt(ip);
				x.status = "Up";
				
				try
				{
					x.host = await Dns.GetHostEntryAsync(ip);
				}
				catch
				{
					x.host = null;
				}
				finally
				{
					lock (IpScan_Async.x)
					{
						nFound++;
					}

					pdatas.Add(x);
				}
			}
			else
			{

				x.IPbin = IpToInt(ip);
				x.status = "Down";
				//_textbox.AppendText(IntToIp(x.IPbin).ToString() + " is " + x.status.ToString()+"\n");

			}


		}




		/// ////////////////////////////////////////////////////////
		/// Config txt data
		//constructor2 ()
		public IpScan_Async(int i)
		{
			run();

		}

		public async void run()
		{

			ConfigTxt con = new ConfigTxt();
			List<string> sline = con.ReadTxt();
			var tasks = new List<Task>();

			for (int j = 0; j <= sline.Count - 1; j++)
			{

				string StartIP = sline[j] + ".1";
				string StopIP = sline[j] + ".255";


				try
				{
					int IntStartIP = IpToInt(StartIP);
					int IntStopIP = IpToInt(StopIP);
					int n = IntStopIP - IntStartIP;

					if (n >= 0)
					{
						//_textbox.AppendText("J is " + j + "\n");
						var task = RunPingSweep_Async_a(StartIP, StopIP);
						tasks.Add(task);
					}

				}
				catch (Exception e)
				{

					MessageBox.Show(e + "IP格式錯誤\n");

				}


			}

			await Task.WhenAll(tasks).ContinueWith(t =>
			{
				//SORT pdatas compare to IPbin
				pdatas.Sort((s1, s2) => s1.IPbin.CompareTo(s2.IPbin));
				//_textbox.AppendText("排序中...\n");
				OpenExcel();

			});


		}

		//RunPingSweep_Async method2
		public async Task RunPingSweep_Async_a(string StartIP, string StopIP)
		{


			nFound = 0;

			var tasks = new List<Task>();

			stopWatch.Start();


			int n = IpToInt(StopIP) - IpToInt(StartIP);



			for (int i = 0; i <= n; i++)
			{


				string ip = IntToIp(IpToInt(StartIP) + i);
				Ping p = new Ping();



				try
				{

					var task = PingAndUpdateAsync(p, ip, (i + 1));
					tasks.Add(task);
				}
				catch (Exception e)
				{
					_textbox.AppendText("\n exception:" + e + " , ");
				}
			}

			await Task.WhenAll(tasks).ContinueWith(t =>
			{

				stopWatch.Stop();


			});



			//SORT pdatas compare to IPbin
			pdatas.Sort((s1, s2) => s1.IPbin.CompareTo(s2.IPbin));

			_textbox.Clear();
			_textbox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;

			for (int j = 0; j <= pdatas.Count - 1; j++)
			{
				try
				{
					_textbox.AppendText(pdatas[j].host.HostName.ToString() + "   " + IntToIp(pdatas[j].IPbin) + " is " + pdatas[j].status + "\n");
				}
				catch
				{

					_textbox.AppendText(IntToIp(pdatas[j].IPbin) + "  ,   " + "   " + IntToIp(pdatas[j].IPbin) + " is " + pdatas[j].status + "\n");
				}
			}

			ts = stopWatch.Elapsed;




			//MessageBox.Show(nFound.ToString() + " devices found! Elapsed time: " + ts.ToString(), "Asynchronous");

		}

		public void prin(int i)
		{


			//Print pdatas
			for (int j = 0; j <= pdatas.Count - 1; j++)
			{
				try
				{
					_textbox.AppendText(pdatas[j].host.HostName.ToString() + "   " + IntToIp(pdatas[j].IPbin) + " is " + pdatas[j].status + "\n");
				}
				catch
				{

					_textbox.AppendText(IntToIp(pdatas[j].IPbin) + "  ,   " + "   " + IntToIp(pdatas[j].IPbin) + " is " + pdatas[j].status + "\n");
				}
			}

			_textbox.AppendText("Finished\n");



		}




		/// <summary>
		/// ////////////////////////////////

		public void OpenExcel()
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
			
			excelApp.Cells[1, 5] = DateTime.Now.ToString("yyyy/MM/dd/HH:mm");

			// 設定第1列顏色
			wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 3]];
			wRange.Select();
			wRange.Font.Color = ColorTranslator.ToOle(Color.White);
			wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);



			for (int i = 0; i <= pdatas.Count - 1; i++)
			{
				try
				{
					excelApp.Cells[i + 2, 1] = IntToIp(pdatas[i].IPbin);
					excelApp.Cells[i + 2, 2] = pdatas[i].host.HostName.ToString();
					excelApp.Cells[i + 2, 3] = pdatas[i].status;
					
				}
				catch
				{
					try
					{
						excelApp.Cells[i + 2, 1] = IntToIp(pdatas[i].IPbin);
						excelApp.Cells[i + 2, 2] = IntToIp(pdatas[i].IPbin);
						excelApp.Cells[i + 2, 3] = pdatas[i].status;
						
					}
					catch { }

				}
				finally
				{
					//excelApp.Cells[i + 2, 7] = "= IF(AND(CODE(B" + (i + 2) + ") = CODE(I" + (i + 2) + "), B" + (i + 2) + " <> \"n/a\"), \"Yes\", IF(AND(CODE(B" + (i + 2) + ") = CODE(I" + (i + 2) + "), B" + (i + 2) + " = \"n/a\"), \"No\", \"Difference\"))";
				}
			}

			excelApp.Cells[1, 4] = " Up";
			excelApp.Cells[2, 4] = nFound;








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

		static bool a = false;
		static int b=0;
		//Tcp port connect
		public void RunScanTcp(string host)
		{

			TcpClient tcp = new TcpClient();


			try
			{
				tcp = new TcpClient(host, 22);
				a = true;
				b = 22;
			}
			catch
			{

			}
			finally
			{

				try
				{
					tcp.Close();


				}
				catch 
				{

				}
			}
			try
			{
				tcp = new TcpClient(host, 80);
				a = true;
				b = 80;
			}
			catch
			{

			}
			finally
			{

				try
				{
					tcp.Close();


				}
				catch 
				{

				}
			}
			try
			{
				tcp = new TcpClient(host, 135);
				a = true;
				b = 135;
			}
			catch
			{

			}
			finally
			{

				try
				{
					tcp.Close();


				}
				catch 
				{

				}
			}
			try
			{
				tcp = new TcpClient(host, 139);
				a = true;
				b = 139;
			}

			catch
			{

			}
			finally
			{

				try
				{
					tcp.Close();


				}
				catch 
				{

				}
			}
		}
	
		//Function transformation
		public int IpToInt(string ip)
		{
			char[] spertator = new char[] { '.' };
			string[] items = ip.Split(spertator);
			return int.Parse(items[0]) << 24 | int.Parse(items[1]) << 16 | int.Parse(items[2]) << 8 | int.Parse(items[3]);


		}
		public string IntToIp(int ipInt)
		{
			StringBuilder sb = new StringBuilder();
			sb.Append((ipInt >> 24) & 0xFF).Append(".");
			sb.Append((ipInt >> 16) & 0xFF).Append(".");
			sb.Append((ipInt >> 8) & 0xFF).Append(".");
			sb.Append(ipInt & 0xFF);
			return sb.ToString();
		}

	}
}



