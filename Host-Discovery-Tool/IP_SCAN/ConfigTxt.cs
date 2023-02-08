using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IP_SCAN
{
	class ConfigTxt
	{
		

		public List<string> ReadTxt()
		{
			string path = "D:\\test.txt";
			StreamReader sr = new StreamReader(path, Encoding.Default);
			string line;
			//一行一IP
			List<string> sline = new List<string>();

			while ((line = sr.ReadLine()) != null)
			{
				sline.Add(line);
			}

			sline.Sort();

			return sline;
		}
		

	}
}
