using HelperLibrary.POCOs;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace HelperLibrary
{
    public static class ProcessExtensions
    {
        private static ProcessResult RunProcessInternal(this Process process, string fileName, string arguments, string pcAddress = "localhost", bool isRemote = false)
        {
            process.StartInfo.FileName = fileName;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.CreateNoWindow = true;

            try
            {
                process.Start();
            }
            catch (Exception ex)
            {
                return new ProcessResult
                {
                    PCAddress = pcAddress,
                    StandardError = ex.Message,
                };
            }

            process.WaitForExit();

            return new ProcessResult
            {
                PCAddress = pcAddress,
                ExitCode = process.ExitCode,
                StandardOutput = process.StandardOutput.ReadToEnd(),
                StandardError = process.StandardError.ReadToEnd()
            };
        }

        public static ProcessResult RunProcess(this Process process, string fileName, string arguments)
        {
            return process.RunProcessInternal(fileName, arguments);
        }

        public static ProcessResult RunPSCommand(this Process process, string command)
        {
            return process.RunProcessInternal("powershell.exe", $"-noninteractive -executionpolicy bypass -WindowStyle Hidden -command \"{command}\"");
        }

        public static ProcessResult RunPSFile(this Process process, string scriptPath)
        {
            return process.RunProcessInternal("powershell.exe", $"-noninteractive -executionpolicy bypass -WindowStyle Hidden -file \"{scriptPath}\"");
        }

        public static ProcessResult RunBatchFile(this Process process, string scriptPath)
        {
            return process.RunProcessInternal("cmd.exe", $"/c \"{scriptPath}\"");
        }

        public static ProcessResult InstallMSI(this Process process, string msiPath, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/i \"{msiPath}\" {arguments}");
        }

        public static ProcessResult RepairMSI(this Process process, string productCode, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/fa {productCode} {arguments}");
        }

        public static ProcessResult UninstallMSI(this Process process, string msiPath, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/x \"{msiPath}\" {arguments}");
        }

        public static ProcessResult UninstallMSIByProductCode(this Process process, string productCode, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/x {productCode} {arguments}");
        }

        public static ProcessResult InstallMSP(this Process process, string mspPath, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/p \"{mspPath}\" {arguments}");
        }

        public static ProcessResult RepairMSP(this Process process, string mspPath, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/fa \"{mspPath}\" {arguments}");
        }

        public static ProcessResult UninstallMSP(this Process process, string mspPath, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/p \"{mspPath}\" {arguments}");
        }

        public static ProcessResult UninstallMSPByProductCode(this Process process, string productCode, string arguments = "/qn /norestart")
        {
            return process.RunProcessInternal("msiexec.exe", $"/p {productCode} {arguments}");
        }

        public static ProcessResult InstallEXE(this Process process, string exePath, string arguments = "/s /v\"/qn /norestart\"")
        {
            return process.RunProcessInternal(exePath, arguments);
        }

        public static ProcessResult RepairEXE(this Process process, string exePath, string arguments = "/s /v\"/qn /norestart\"")
        {
            return process.RunProcessInternal(exePath, arguments);
        }

        public static ProcessResult UninstallEXE(this Process process, string exePath, string arguments = "/s /v\"/qn /norestart\"")
        {
            return process.RunProcessInternal(exePath, arguments);
        }

        public static ProcessResult InstallRegFile(this Process process, string regPath)
        {
            return process.RunProcessInternal("regedit.exe", $"/s \"{regPath}\"");
        }

        public static ProcessResult UninstallRegFile(this Process process, string regPath)
        {
            return process.RunProcessInternal("regedit.exe", $"/s \"{regPath}\"");
        }

        public static ProcessResult RunRemoteProcess(this Process process, string pcAddress, string arguments)
        {
            return process.RunProcessInternal("PsExec.exe", $"\\\\{pcAddress} -accepteula -nobanner cmd /c \"{arguments}\"", pcAddress, true);
        }

        public static ProcessResult RunRemotePSCommand(this Process process, string pcAddress, string command)
        {
            return process.RunRemoteProcess(pcAddress, $"powershell -noninteractive -executionpolicy bypass -WindowStyle Hidden -command \"{command}\"");
        }

        public static ProcessResult RunRemotePSFile(this Process process, string pcAddress, string scriptPath)
        {
            return process.RunRemoteProcess(pcAddress, $"powershell -noninteractive -executionpolicy bypass -WindowStyle Hidden -file \"{scriptPath}\"");
        }

        public static ProcessResult RunRemoteBatchFile(this Process process, string pcAddress, string scriptPath)
        {
            return process.RunRemoteProcess(pcAddress, $"cmd /c \"{scriptPath}\"");
        }

        public static ProcessResult InstallRemoteMSI(this Process process, string pcAddress, string msiPath, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /i \"{msiPath}\" {arguments}");
        }

        public static ProcessResult RepairRemoteMSI(this Process process, string pcAddress, string productCode, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /fa {productCode} {arguments}");
        }

        public static ProcessResult UninstallRemoteMSI(this Process process, string pcAddress, string msiPath, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /x \"{msiPath}\" {arguments}");
        }

        public static ProcessResult UninstallRemoteMSIByProductCode(this Process process, string pcAddress, string productCode, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /x {productCode} {arguments}");
        }

        public static ProcessResult InstallRemoteMSP(this Process process, string pcAddress, string mspPath, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /p \"{mspPath}\" {arguments}");
        }

        public static ProcessResult RepairRemoteMSP(this Process process, string pcAddress, string mspPath, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /fa \"{mspPath}\" {arguments}");
        }

        public static ProcessResult UninstallRemoteMSP(this Process process, string pcAddress, string mspPath, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /p \"{mspPath}\" {arguments}");
        }

        public static ProcessResult UninstallRemoteMSPByProductCode(this Process process, string pcAddress, string productCode, string arguments = "/qn /norestart")
        {
            return process.RunRemoteProcess(pcAddress, $"msiexec /p {productCode} {arguments}");
        }

        public static ProcessResult InstallRemoteEXE(this Process process, string pcAddress, string exePath, string arguments = "/s /v\"/qn /norestart\"")
        {
            return process.RunRemoteProcess(pcAddress, $"{exePath} {arguments}");
        }

        public static ProcessResult RepairRemoteEXE(this Process process, string pcAddress, string exePath, string arguments = "/s /v\"/qn /norestart\"")
        {
            return process.RunRemoteProcess(pcAddress, $"{exePath} {arguments}");
        }

        public static ProcessResult UninstallRemoteEXE(this Process process, string pcAddress, string exePath, string arguments = "/s /v\"/qn /norestart\"")
        {
            return process.RunRemoteProcess(pcAddress, $"{exePath} {arguments}");
        }

        public static ProcessResult InstallRemoteRegFile(this Process process, string pcAddress, string regPath)
        {
            return process.RunRemoteProcess(pcAddress, $"regedit /s \"{regPath}\"");
        }

        public static ProcessResult UninstallRemoteRegFile(this Process process, string pcAddress, string regPath)
        {
            return process.RunRemoteProcess(pcAddress, $"regedit /s \"{regPath}\"");
        }

        public static List<QUserSession> RunQUser(this Process process)
        {
            var result = process.RunProcessInternal("quser.exe", null);
            return Helpers.ConvertProcessResultToSessions(result);
        }

        public static List<QUserSession> RunQuser(this Process process, string username, string pcAddress)
        {
            var result = process.RunRemoteProcess(pcAddress, $"{username} /server:{pcAddress}");
            if (result.StandardOutput.Contains("No User exists for"))
            {
                return null;
            }
            return Helpers.ConvertProcessResultToSessions(result);
        }
    }
}