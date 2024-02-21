using System.Globalization;
using System.Text;

namespace SharePointInterface
{
	internal static class Logger
	{
		internal static void Log(string line)
		{
			using StreamWriter sw = new("client.log", true, Encoding.UTF8) { AutoFlush = true };
			sw.WriteLine($"[{DateTime.Now.ToString("u", DateTimeFormatInfo.InvariantInfo)}, PID: {Environment.ProcessId}] {line}");
		}
	}
}
