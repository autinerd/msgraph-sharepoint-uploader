using System.Globalization;
using System.Text;

namespace SharePointUploader
{
	internal static class Logger
	{
		internal static void Log(string line)
		{
			using var sw = new StreamWriter("client.log", true, Encoding.UTF8) {AutoFlush = true};
			sw.WriteLine($"[{DateTime.Now.ToString("u", DateTimeFormatInfo.InvariantInfo)}, {Environment.ProcessId}] {line}");
		}
	}
}
