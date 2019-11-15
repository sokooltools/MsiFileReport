using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using MsiFileReport.Properties;

// Make sure to add a reference to Interop.WindowsInstaller.dll 

namespace MsiFileReport
{
    internal static class Program
    {
        private const string MSI_FLDR = @"C:\Windows\Installer";
        private const string RPT_FILE = "MsiFileReport.htm";

        //----------------------------------------------------------------------------------------------------
        /// <summary>
        /// The HTML file
        /// </summary>
        //----------------------------------------------------------------------------------------------------
        private static readonly string HtmlFile = Resources.HtmlFile;

        //----------------------------------------------------------------------------------------------------
        /// <summary>
        /// The main entry point of this application.
        /// </summary>
        /// <param name="args">
        /// The optional arguments. The first argument indicating that the html output should be constrained 
        /// to only those 'Manufacturers' whose name contains the specified value (e.g. "Parker").
        /// </param>
        //----------------------------------------------------------------------------------------------------
        public static void Main(string[] args)
        {
            ConsoleHelper.Setup();

            Console.WriteLine(@"Collecting data... Please wait.");
            IEnumerable<MsiFileHelper.MsiFileInfo> msiFiles = MsiFileHelper.GetMsiFileInfo(MSI_FLDR, args?.Length > 0 ? args[0] : string.Empty);

            Console.WriteLine(@"Creating HTML...");
            string table = HtmlHelper.GetTableFromList(msiFiles,
                m => m.Manufacturer,
                m => m.ProductName,
                m => m.ProductVersion,
                m => m.FileName,
                m => m.Modified);

            Console.WriteLine(@"Inserting Table into HTML...");
            string html = HtmlFile.Replace("<table></table>", table);
            string outputFile = Path.Combine(Path.GetTempPath(), RPT_FILE);
            File.WriteAllText(outputFile, html);

            Console.WriteLine(@"Launching HTML in Browser...");
            var psi = new ProcessStartInfo(outputFile) { UseShellExecute = true };
            Process.Start(psi);
        }
    }
}
