using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IP_SCAN
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			
		}

		private void button1_Click(object sender, EventArgs e)
		{
			//textBox1.AppendText("button_click\n");
			//傳送Textbox的控制權
			IpScan_Async._textbox = textBox1;
			IpScan_Async._textbox2 = textBox2;
			IpScan_Async._textbox3 = textBox3;

			textBox1.Text = "";
			IpScan_Async sc1 = new IpScan_Async();





		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{

		}


		public void print(string str)
		{
			textBox1.AppendText(str);

		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}

		private void textBox4_TextChanged(object sender, EventArgs e)
		{

		}
		public void textBox5_TextChanged(object sender, EventArgs e)
		{

		}

		private void button2_Click(object sender, EventArgs e)
		{

			//傳送Textbox的控制權
			PortScan_Async._textbox4 = textBox4;
			PortScan_Async._textbox5 = textBox5;


			PortScan_Async p = new PortScan_Async();

		}

		private void button3_Click(object sender, EventArgs e)
		{
			compare cm = new compare();
			
				


			
		}

		private void button4_Click(object sender, EventArgs e)
		{
			IpScan_Async._textbox = textBox1;
			IpScan_Async._textbox2 = textBox2;
			IpScan_Async._textbox3 = textBox3;

			
			IpScan_Async sc2 = new IpScan_Async(1);

			
		}
	}
}
