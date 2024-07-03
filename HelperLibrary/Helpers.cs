using HelperLibrary.POCOs;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace HelperLibrary
{
    public static class Helpers
    {
        #region Log Functions
        public static string GetLogPath(string appName, string mode, string extension)
        {
            var powerShell = PowerShell.Create();
            var currentDateTime = DateTime.Now.ToString("MM-dd-yyyy_HH-mm-ss");
            var version = powerShell.GetPSVersion();
            var username = Environment.UserName;
            var hostname = Environment.MachineName;
            var currentDir = Environment.CurrentDirectory;
            // create a logs directory within the current directory if it doesn't exist
            if (!Directory.Exists(Path.Combine(currentDir, "Logs")))
            {
                Directory.CreateDirectory(Path.Combine(currentDir, "Logs"));
            }
            string logName = $"Log-{appName}-{mode}-{username}-{hostname}-{version}-{currentDateTime}.{extension}";
            powerShell.Dispose();
            return Path.Combine($"{currentDir}\\Logs", logName);
        }

        public static void CreateLogFile(string logPath)
        {
            if (!File.Exists(logPath))
            {
                File.Create(logPath).Dispose();
            }
        }

        public static void AddLogEntry(string logPath, string message)
        {
            using (var sw = File.AppendText(logPath))
            {
                sw.WriteLine($"{DateTime.Now}: {message}");
            }
        }

        #endregion

        public static T GetValidUserInput<T>(string prompt)
        {
            while (true)
            {
                Console.Write(prompt);
                string input = Console.ReadLine().Trim();
                try
                {
                    return (T)Convert.ChangeType(input, typeof(T));
                }
                catch
                {
                    Console.Write("Invalid input. Press ENTER to try again...");
                    Console.ReadLine();
                }
            }
        }

        public static int GetValidListSelection<T>(List<T> values, string prompt, string heading)
        {
            int result = 0;
            while (true)
            {
                Console.WriteLine(heading);
                for (int i = 0; i < values.Count; i++)
                    Console.WriteLine($"{i + 1}. {values[i]}");
                result = GetValidUserInput<int>(prompt);
                if (result < 1 || result > values.Count)
                {
                    Console.Write("Invalid selection. Press ENTER to try again...");
                    Console.ReadLine();
                }
                else
                    break;
            }
            return result;
        }

        internal static List<QUserSession> ConvertProcessResultToSessions(ProcessResult processResult)
        {
            var sessions = new List<QUserSession>();
            string pattern = @"\s{2,}";
            string replacement = ",";
            string[] separators = { "\r\n" };
            string[] lines = processResult.StandardOutput.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 1; i < lines.Length; i++)
            {
                lines[i] = lines[i].Replace(">", "");
                lines[i] = lines[i].Trim();
                lines[i] = Regex.Replace(lines[i], pattern, replacement);
                string[] parts = lines[i].Split(',');
                sessions.Add(new QUserSession
                {
                    ComputerName = processResult.PCAddress,
                    UserName = parts.Length >= 1 ? parts[0] : string.Empty,
                    SessionName = parts.Length >= 2 ? parts[1] : string.Empty,
                    Id = parts.Length >= 3 ? parts[2] : string.Empty,
                    State = parts.Length >= 4 ? parts[3] : string.Empty,
                    IdleTime = parts.Length >= 5 ? parts[4] : string.Empty,
                    LogonTime = parts.Length >= 6 ? parts[5] : string.Empty
                });
            }
            return sessions;
        }

        public static void PrintLoggedInComputers(string usersso, List<QUserSession> sessions)
        {
            var userSessions = sessions.FindAll(s => s.UserName == usersso);
            if (userSessions.Count == 0)
            {
                Console.WriteLine("No computers are currently logged in with the specified username.");
                return;
            }
            Console.WriteLine($"Computers currently logged in with the username '{usersso}':");
            ConsoleTable.From(userSessions).Write();
        }

        private static string FormatTimeTaken(TimeSpan timeTaken) =>
            timeTaken.TotalSeconds < 1 ? "Less than a second" : timeTaken.ToString(@"hh\:mm\:ss");

        private static string EscapeCsvField(string field)
        {
            if (field == null || field == string.Empty) return string.Empty;
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                field = field.Replace("\"", "\"\"");
                field = field.Replace("\r\n", string.Empty);
                field = $"\"{field}\"";
            }
            return field;
        }

        public static void ExportProcessResultsToCsv(List<ProcessResult> processResults, string logPath)
        {
            var csv = new StringBuilder();
            csv.AppendLine("PC Address,Exit Code,Standard Output,Standard Error");
            foreach (var processResult in processResults)
            {
                csv.AppendLine($"{EscapeCsvField(processResult.PCAddress)},{processResult.ExitCode},{EscapeCsvField(processResult.StandardOutput)},{EscapeCsvField(processResult.StandardError)}");
            }
            File.WriteAllText(logPath, csv.ToString());
        }

        public static void AddTimeTakenToCsv(string logPath, TimeSpan timeTaken)
        {
            var csv = new StringBuilder();
            csv.AppendLine($"Time Taken: {FormatTimeTaken(timeTaken)}");
            File.AppendAllText(logPath, csv.ToString());
        }

        public static bool IsValidComputer(string pc, PowerShell powerShell)
        {
            return pc.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                   (powerShell.IsValidADComputer(pc) && powerShell.IsPCOnline(pc));
        }

        public static void PrinStatus(string message)
        {
            Console.WriteLine(message);
        }
    }
}