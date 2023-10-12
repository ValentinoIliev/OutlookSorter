using System;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace OutlookSorter.Workers;

public class Worker
{
	public Worker() {
		Outlook.Application outlookApp = new Outlook.Application();
		new Azure(outlookApp);
	}
}

