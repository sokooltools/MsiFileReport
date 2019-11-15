using System;
using System.Runtime.InteropServices;

namespace MsiFileReport
{
	//----------------------------------------------------------------------------------------------------------------------------
	/// <summary>
	/// Provides helper methods for the <see cref="System.Console"/>
	/// </summary>
	//----------------------------------------------------------------------------------------------------------------------------
	internal static class ConsoleHelper
	{
		private const int SW_HIDE = 0;
		private const int SW_SHOW = 5;

		[DllImport("kernel32.dll")]
		private static extern IntPtr GetConsoleWindow();

		[DllImport("user32.dll")]
		private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Setups the console with some initial properties such as its title, window size, buffer size, etc.
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		public static void Setup()
		{
			IntPtr handle = GetConsoleWindow();
			ShowWindow(handle, SW_HIDE);
			Console.Title = @"MSI File Report";
			Console.WindowHeight = 14;
			Console.WindowWidth = 50;
			Console.BufferHeight = 14;
			Console.BufferWidth = 50;
			ShowWindow(handle, SW_SHOW);
		}
	}
}
