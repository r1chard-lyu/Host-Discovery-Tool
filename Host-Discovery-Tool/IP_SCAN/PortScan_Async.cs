using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;

namespace IP_SCAN
{
    class PortScan_Async
    {
        public static TextBox _textbox4,_textbox5;
        string ip =_textbox4.Text;


        //Thread mythread = null;

        //Port scan constructor
        public PortScan_Async()
        {
            RunPortSweep_Async(ip);
            //mythread = new Thread(() => RunPortSweep_Async(ip));
            //mythread.Start();
        }
		



		//Port data
		struct PortStatus
        {
            public int portnumber;
            public bool tcp;
            //public bool udp;
  
        }
        List<PortStatus> ports = new List<PortStatus>();


        public async void RunPortSweep_Async(string ip )
        {
            var tasks = new List<Task>();

            
                _textbox5.AppendText("IP : "+ ip +"  \n");

            for (int i = 1; i <=1023; i++)
            {
				if(i==1)
				_textbox5.AppendText("1 \n");

                try
                {
                    var task = PortConnect(ip, i);
                    tasks.Add(task);
                }
                catch (Exception e)
                {
                    _textbox5.AppendText("\n exception:" + e + " , ");
                }

                

            }



            await Task.WhenAll(tasks).ContinueWith(t =>
            {
                _textbox5.AppendText(" Finished!! ");
            });
            _textbox5.AppendText(" Finished!! ");

        }

        private async Task<PortStatus> PortConnect(string ip,int port)
        {
            PortStatus x = new PortStatus();            
            TcpClient tcp = new TcpClient();
            //UdpClient udp = new UdpClient();

            x.portnumber = port;

            try
            {
                
                await tcp.ConnectAsync(ip, port);
                x.tcp = true;
                _textbox5.AppendText("TCP Port "+ port + "is open\n");
            }
            catch
            {
                //_textbox.AppendText("TCP Port " + port + "is close\n");

            }
            finally
            {
                try
                {
                    tcp.Close();


                }
                catch (Exception ex)
                {
                   _textbox5.AppendText(ex.Message);
                }
            }


            try
            {
                //udp = new UdpClient(ip, port);
                //await udp.ReceiveAsync();
                //x.udp = true;
                //_textbox5.AppendText("UDP Port " + port + "is open\n");
            }
            catch
            {
                //_textbox5.AppendText("UDP Port " + port + "is close\n");

            }
            finally
            {
                try
                {
                    
                    //udp.Close();


                }
                catch (Exception ex)
                {
                   _textbox5.AppendText(ex.Message);
                }
            }

            
            return x;


		} 

		
    }   
}
      