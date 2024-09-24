using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintDialogX
{
	class Setting
	{
		// Folder location
		public static string g_BaseFolder = System.Windows.Forms.Application.StartupPath;
		public static string g_ConfigPath = g_BaseFolder + @"\Config.ini";
	}
}
