using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace GetOutlook
{
    public class GetFolder
    {
        public static MAPIFolder GetFolders(string folderPath, string account) 
        {
			if (string.IsNullOrWhiteSpace(folderPath))
			{
				return null;
			}
			Application application = null;
			try
			{
				application = InitOutlook();
				MAPIFolder mAPIFolder = null;
				if (string.IsNullOrWhiteSpace(account))
				{
					account = application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox).FolderPath;
					account = account.Substring(2);
					if (account.IndexOf('\\') < account.Length)
					{
						account = account.Substring(0, account.IndexOf('\\'));
					}
				}
				for (int i = 1; i <= application.Session.Folders.Count; i++)
				{
					if (application.Session.Folders[i].Name == account)
					{
						mAPIFolder = application.Session.Folders[i];
						break;
					}
				}
				if (mAPIFolder == null)
				{
					throw new ArgumentException("AccountNotMappedLocally");
				}
				return FindFolder(folderPath, mAPIFolder);
			}
			catch (System.Exception ex)
			{
				Trace.TraceWarning(ex.ToString());
			}
			finally
			{
				if (application != null)
				{
					Marshal.ReleaseComObject(application);
				}
			}
			return null;
		}

        public static Application InitOutlook()
        {
			Application application = null;
			try
			{
				application = (Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("0006F03A-0000-0000-C000-000000000046")));
				string version = application.Version;
				if (version != null && version.StartsWith("14.0"))
				{									
					//ss
  					 return application;											
				}
				return application;
			}
			catch (COMException ex)
			{
				if (ex.ErrorCode == -2147221164)
				{
					Trace.TraceError(ex.ToString());
					throw new SystemException("InitializeOutlookError" + Environment.NewLine + "IsOutlookInstalled", ex);
				}
				Trace.TraceError(ex.ToString());
				throw new SystemException(ex.Message, ex);
			}
			catch
			{
				if (application != null)
				{
					Marshal.ReleaseComObject(application);
				}
				throw;
			}
		}

        private static MAPIFolder FindFolder(string folderPath, MAPIFolder rootFolder)
        {
			Folders folders = null;
			try
			{
				folders = rootFolder.Folders;
				string[] array = folderPath.Split('\\');
				MAPIFolder mAPIFolder = null;
				string[] array2 = array;
				foreach (string stringToUnescape in array2)
				{
					string folderName = Uri.UnescapeDataString(stringToUnescape);
					mAPIFolder = folders.OfType<MAPIFolder>().FirstOrDefault((MAPIFolder f) => string.Equals(folderName, f.Name, StringComparison.InvariantCultureIgnoreCase));
					if (mAPIFolder == null)
					{
						throw new ArgumentException("MissingFolder");
					}
					folders = mAPIFolder.Folders;
				}
				return mAPIFolder;
			}
			finally
			{
				if (folders != null && Marshal.IsComObject(folders))
				{
					Marshal.ReleaseComObject(folders);
				}
			}
		}
    }
}
