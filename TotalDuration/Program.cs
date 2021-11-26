using System;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace TotalDuration
{
    class Program
    {
        static void Main(string[] args)
        {
            var dir = Environment.CurrentDirectory;

            if (args.Length > 0)
                if (IsValidDirectory(args[0]))
                    dir = args[0];

            var totalDuration = Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories)
                .Select(f => (long)(GetFileDuration(f)/10000))
                .Sum();

            var span = TimeSpan.FromMilliseconds(totalDuration);

            Console.WriteLine($"Total Duration is: {FormatSpan(span)}");
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static string FormatSpan(TimeSpan span)
        {
            var b = new StringBuilder();

            if (span.Days > 0)
                b.Append($"{span.Days} days ");

            if (span.Hours > 0)
                b.Append($"{span.Hours} hours ");

            if (span.Minutes > 0)
                b.Append($"{span.Minutes} minutes ");

            if (span.Seconds > 0)
                b.Append($"{span.Seconds} seconds ");

            if (span.Milliseconds > 0)
                b.Append($"{span.Milliseconds} milliseconds ");

            var formatted = b.ToString();

            if (string.IsNullOrWhiteSpace(formatted))
                return "None";

            return formatted.Trim();
        }

        private static bool IsValidDirectory(string dir)
        {
            try
            {
                var result = Directory.GetFiles(dir);
                return result != null;
            }
            catch
            {
                return false;
            }
        }

        private static ulong GetFileDuration(string filePath) //in 100ns
        {
            try
            {
                using (ShellObject shell = ShellObject.FromParsingName(filePath))
                {
                    IShellProperty prop = shell.Properties.System.Media.Duration;
                    return (ulong)prop.ValueAsObject;
                }
            }
            catch
            {
                return 0;
            }
        }
    }
}
