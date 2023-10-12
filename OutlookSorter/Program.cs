using System;
using System.Linq;
using OutlookSorter.Workers;
using Outlook = Microsoft.Office.Interop.Outlook;

class Program
{
	static void Main(string[] args) {
		new Worker();
		/*
		// Create an Outlook application object
		Outlook.Application outlookApp = new Outlook.Application();

		// Get the Outlook NameSpace
		Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

		// Get the root folder
		Outlook.MAPIFolder parentFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
		// Create a new folder
		Outlook.MAPIFolder child = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders[1];
		Outlook.MAPIFolder newFolder = child.Folders.Add("PR2", Outlook.OlDefaultFolders.olFolderInbox);

		// Release resources
		System.Runtime.InteropServices.Marshal.ReleaseComObject(newFolder);
		System.Runtime.InteropServices.Marshal.ReleaseComObject(parentFolder);
		System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNamespace);
		System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);


	*/
		//Console.WriteLine(rootFolder.Name);

		/*
		// Create a new email item
		Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

		// Set email properties
		mailItem.Subject = "Subject";
		mailItem.Body = "Email Body";
		mailItem.To = "valentino.iliev@elvexys.com";

		// Send the email
		mailItem.Send();

		// Release resources
		System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
		System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
		*/
	}
}