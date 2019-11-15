using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WindowsInstaller;

namespace MsiFileReport
{
	internal static class MsiFileHelper
	{
		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Provides some information about an individual MSI file.
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		[Serializable]
		public class MsiFileInfo
		{
			public string FileName { get; set; }
			public string Manufacturer { get; set; }
			public string ProductName { get; set; }
			public string ProductVersion { get; set; }
			public string Modified { get; set; }
		}

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Gets the msi file information.
		/// </summary>
		/// <param name="folderPath">The folder path.</param>
		/// <param name="mfg">The manufacturer contains this value.</param>
		/// <returns></returns>
		//------------------------------------------------------------------------------------------------------------------------
		public static IEnumerable<MsiFileInfo> GetMsiFileInfo(string folderPath, string mfg = "")
		{
			FileInfo[] fileInfo = new DirectoryInfo(folderPath).GetFiles("*.msi", SearchOption.TopDirectoryOnly);
			Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
			var installer = Activator.CreateInstance(classType) as Installer;
			IOrderedEnumerable<MsiFileInfo> fi = fileInfo
                .Select
                (
                  msiFile => new MsiFileInfo
			        {
				        FileName = Path.GetFileName(msiFile.FullName),
				        Manufacturer = GetMsiProperty(installer, msiFile.FullName, "Manufacturer"),
				        ProductName = GetMsiProperty(installer, msiFile.FullName, "ProductName"),
				        ProductVersion = GetMsiProperty(installer, msiFile.FullName, "ProductVersion"),
				        Modified = new FileInfo(msiFile.FullName).LastWriteTime.ToString("MM-dd-yyyy hh:mm:ss tt")
			        }
                )
                .Where(a => a.Manufacturer.Contains(mfg))
                .OrderBy(a => a.Manufacturer).ThenBy(a => a.ProductName).ThenBy(a => a.ProductVersion);
			return fi;
		}

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Gets the msi property.
		/// </summary>
		/// <param name="installer"></param>
		/// <param name="msiFullFileName">The full name of the msi file.</param>
		/// <param name="property">The property.</param>
		/// <returns></returns>
		//------------------------------------------------------------------------------------------------------------------------
		private static string GetMsiProperty(Installer installer, string msiFullFileName, string property)
		{
			string retVal = string.Empty;
			// Open the msi file for reading. (0 - Read, 1 - Read/Write)
			if (installer == null) 
                return retVal;
			Database database = installer.OpenDatabase(msiFullFileName, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);
			View view = database.OpenView($"SELECT `Value` FROM `Property` WHERE Property='{ property }'");
			view.Execute();
			Record record = view.Fetch();
			if (record != null)
			{
				retVal = record.StringData[1];
				System.Runtime.InteropServices.Marshal.FinalReleaseComObject(record);
			}
			view.Close();
			System.Runtime.InteropServices.Marshal.FinalReleaseComObject(view);
			System.Runtime.InteropServices.Marshal.FinalReleaseComObject(database);
			return retVal;
		}
	}
}
