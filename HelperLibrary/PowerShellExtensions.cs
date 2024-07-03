using HelperLibrary.POCOs;
using Microsoft.ActiveDirectory.Management;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;

namespace HelperLibrary
{
    public static class PowerShellExtensions
    {
        #region Invoke and Clear

        private static Collection<T> InvokeAndClearInternal<T>(this PowerShell powershell)
        {
            try
            {
                return powershell.Invoke<T>();
            }
            finally
            {
                powershell.Commands.Clear();
            }
        }

        private static Collection<PSObject> InvokeAndClearInternal(this PowerShell powershell)
        {
            try
            {
                return powershell.Invoke();
            }
            finally
            {
                powershell.Commands.Clear();
            }
        }

        private static Collection<T> InvokeAndClear<T>(this PowerShell powershell) => powershell.InvokeAndClearInternal<T>();

        private static Collection<PSObject> InvokeAndClear(this PowerShell powershell) => powershell.InvokeAndClearInternal();

        #endregion

        #region Utilities

        public static string GetPSVersion(this PowerShell powerShell) =>
            powerShell.AddScript("$PSVersionTable.PSVersion.ToString()").InvokeAndClear<string>().FirstOrDefault();

        public static string GetExecutionPolicy(this PowerShell powerShell) =>
            powerShell.AddScript("(Get-ExecutionPolicy).ToString()").InvokeAndClear<string>().FirstOrDefault();

        public static bool IsPCOnline(this PowerShell powerShell, string pcAddress) =>
            powerShell.AddScript($"Test-Connection -ComputerName {pcAddress} -Count 1 -Quiet").InvokeAndClear<bool>().FirstOrDefault();

        public static bool IsPSModuleAvailable(this PowerShell powerShell, string moduleName) =>
            powerShell.AddScript($"Get-Module -ListAvailable -Name {moduleName}").InvokeAndClear().Any();

        public static bool IsADModuleAvailable(this PowerShell powerShell) =>
            powerShell.IsPSModuleAvailable("ActiveDirectory");

        public static bool IsValidADComputer(this PowerShell powerShell, string computerName) =>
            powerShell.GetADComputer(computerName) != null;

        public static void RestartPC(this PowerShell powerShell, string pcAddress)
        {
            if (!powerShell.IsPCOnline(pcAddress))
                throw new Exception("PC is offline");

            powerShell.AddScript($"Restart-Computer -ComputerName {pcAddress} -Force").InvokeAndClear();
        }

        public static void RestartPCs(this PowerShell powerShell, List<string> pcAddresses)
        {
            foreach (var pcAddress in pcAddresses)
            {
                try
                {
                    powerShell.RestartPC(pcAddress);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to restart {pcAddress}: {ex.Message}");
                }
            }
        }

        public static void ShutdownPC(this PowerShell powerShell, string pcAddress)
        {
            if (!powerShell.IsPCOnline(pcAddress))
                throw new Exception("PC is offline");

            powerShell.AddScript($"Stop-Computer -ComputerName {pcAddress} -Force").InvokeAndClear();
        }

        public static void ShutdownPCs(this PowerShell powerShell, List<string> pcAddresses)
        {
            foreach (var pcAddress in pcAddresses)
            {
                try
                {
                    powerShell.ShutdownPC(pcAddress);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to shutdown {pcAddress}: {ex.Message}");
                }
            }
        }

        public static bool IsPCBackOnline(this PowerShell powerShell, string pcAddress, int? timeout = null)
        {
            if (timeout == null)
                timeout = 60;

            DateTime startTime = DateTime.Now;
            while (!powerShell.IsPCOnline(pcAddress))
            {
                if ((DateTime.Now - startTime).TotalSeconds > timeout)
                    return false;
            }
            return true;
        }

        #endregion

        #region IO Helpers

        private static bool TestPath(this PowerShell powerShell, string path) =>
            powerShell.AddScript($"Test-Path -Path '{path}'").InvokeAndClear<bool>().FirstOrDefault();

        private static IOResult ExecuteIOScript(this PowerShell powerShell, string script, string sourcePath, string destinationPath = null)
        {
            try
            {
                powerShell.AddScript(script).InvokeAndClear();
                if (powerShell.Streams.Error.Count > 0)
                {
                    return new IOResult
                    {
                        SourcePath = sourcePath,
                        DestinationPath = destinationPath,
                        OperationSuccessful = false,
                        ErrorMessage = powerShell.Streams.Error[0].Exception.Message
                    };
                }
                return new IOResult
                {
                    SourcePath = sourcePath,
                    DestinationPath = destinationPath,
                    OperationSuccessful = true
                };
            }
            catch (Exception ex)
            {
                return new IOResult
                {
                    SourcePath = sourcePath,
                    DestinationPath = destinationPath,
                    OperationSuccessful = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private static IOResult CreateDirectory(this PowerShell powerShell, string path) =>
            powerShell.ExecuteIOScript($"New-Item -Path '{path}' -ItemType Directory -Force -ErrorAction Stop", path);

        public static IOResult ExpandArchive(this PowerShell powerShell, string sourcePath, string destinationPath)
        {
            if (!powerShell.TestPath(sourcePath))
                return new IOResult
                {
                    SourcePath = sourcePath,
                    DestinationPath = destinationPath,
                    OperationSuccessful = false,
                    ErrorMessage = "Source path does not exist"
                };

            return powerShell.ExecuteIOScript($"Expand-Archive -Path '{sourcePath}' -DestinationPath '{destinationPath}' -Force -ErrorAction Stop", sourcePath, destinationPath);
        }

        public static IOResult ExpandArchiveOnPC(this PowerShell powerShell, string pcAddress, string sourcePath, string destinationPath)
        {
            string newSourcePath = $@"\\{pcAddress}\{sourcePath[0]}${sourcePath.Substring(2)}";
            string newDestinationPath = $@"\\{pcAddress}\{destinationPath[0]}${destinationPath.Substring(2)}";
            return powerShell.ExpandArchive(newSourcePath, newDestinationPath);
            
        }

        public static IOResult CompressArchive(this PowerShell powerShell, string sourcePath, string destinationPath)
        {
            if (!powerShell.TestPath(sourcePath))
                return new IOResult
                {
                    SourcePath = sourcePath,
                    DestinationPath = destinationPath,
                    OperationSuccessful = false,
                    ErrorMessage = "Source path does not exist"
                };

            return powerShell.ExecuteIOScript($"Compress-Archive -Path '{sourcePath}' -DestinationPath '{destinationPath}' -Force -ErrorAction Stop", sourcePath, destinationPath);
        }

        public static IOResult CompressArchiveOnPC(this PowerShell powerShell, string pcAddress, string sourcePath, string destinationPath)
        {
            string newSourcePath = $@"\\{pcAddress}\{sourcePath[0]}${sourcePath.Substring(2)}";
            string newDestinationPath = $@"\\{pcAddress}\{destinationPath[0]}${destinationPath.Substring(2)}";
            return powerShell.CompressArchive(newSourcePath, newDestinationPath);
        }

        private static IOResult CopyItemInternal(this PowerShell powerShell, string source, string destination, bool recurse, string pcAddress = null)
        {
            if (!powerShell.TestPath(source))
                return new IOResult
                {
                    SourcePath = source,
                    DestinationPath = destination,
                    OperationSuccessful = false,
                    ErrorMessage = "Source path does not exist"
                };

            if (!powerShell.TestPath(destination))
            {
                var result = powerShell.CreateDirectory(destination);
                if (!result.OperationSuccessful)
                    return result;
            }

            string recurseFlag = recurse ? "-Recurse" : "";
            string script = $"Copy-Item -Path '{source}' -Destination '{destination}' {recurseFlag} -Force -ErrorAction Stop";
            return powerShell.ExecuteIOScript(script, source, destination);
        }

        private static IOResult CopyItem(this PowerShell powerShell, string source, string destination, bool recurse) =>
            powerShell.CopyItemInternal(source, destination, recurse);

        private static IOResult CopyItemToPC(this PowerShell powerShell, string pcAddress, string source, string destination, bool recurse)
        {
            string newDestinationPath = $@"\\{pcAddress}\{destination[0]}${destination.Substring(2)}";
            return powerShell.CopyItemInternal(source, newDestinationPath, recurse, pcAddress);
        }

        private static IOResult CopyItemFromPC(this PowerShell powerShell, string pcAddress, string source, string destination, bool recurse)
        {
            string newSourcePath = $@"\\{pcAddress}\{source[0]}${source.Substring(2)}";
            return powerShell.CopyItemInternal(newSourcePath, destination, recurse, pcAddress);
        }

        public static IOResult CopyFile(this PowerShell powerShell, string sourcePath, string destinationPath) =>
            powerShell.CopyItem(sourcePath, destinationPath, false);

        public static IOResult CopyFileToPC(this PowerShell powerShell, string pcAddress, string sourcePath, string destinationPath) =>
            powerShell.CopyItemToPC(pcAddress, sourcePath, destinationPath, false);

        public static IOResult CopyFileFromPC(this PowerShell powerShell, string pcAddress, string sourcePath, string destinationPath) =>
            powerShell.CopyItemFromPC(pcAddress, sourcePath, destinationPath, false);

        public static IOResult CopyDirectory(this PowerShell powerShell, string sourcePath, string destinationPath) =>
            powerShell.CopyItem($@"{sourcePath}\*", destinationPath, true);

        public static IOResult CopyDirectoryToPC(this PowerShell powerShell, string pcAddress, string sourcePath, string destinationPath) =>
            powerShell.CopyItemToPC(pcAddress, $@"{sourcePath}\*", destinationPath, true);

        public static IOResult CopyDirectoryFromPC(this PowerShell powerShell, string pcAddress, string sourcePath, string destinationPath) =>
            powerShell.CopyItemFromPC(pcAddress, $@"{sourcePath}\*", destinationPath, true);

        private static IOResult RemoveItemInternal(this PowerShell powerShell, string path, bool recurse, string pcAddress = null)
        {
            if (!powerShell.TestPath(path))
                return new IOResult
                {
                    SourcePath = path,
                    OperationSuccessful = false,
                    ErrorMessage = "Path does not exist"
                };

            string recurseFlag = recurse ? "-Recurse" : "";
            string script = $"Remove-Item -Path '{path}' {recurseFlag} -Force -ErrorAction Stop";
            return powerShell.ExecuteIOScript(script, path);
        }

        private static IOResult RemoveItem(this PowerShell powerShell, string path, bool recurse) =>
            powerShell.RemoveItemInternal(path, recurse);

        private static IOResult RemoveItemOnPc(this PowerShell powerShell, string pcAddress, string path, bool recurse)
        {
            string newPath = $@"\\{pcAddress}\{path[0]}${path.Substring(2)}";
            return powerShell.RemoveItemInternal(newPath, recurse, pcAddress);
        }

        public static IOResult RemoveFile(this PowerShell powerShell, string path) =>
            powerShell.RemoveItem(path, false);

        public static IOResult RemoveFileOnPc(this PowerShell powerShell, string pcAddress, string path) =>
            powerShell.RemoveItemOnPc(pcAddress, path, false);

        public static IOResult RemoveDirectory(this PowerShell powerShell, string path) =>
            powerShell.RemoveItem(path, true);

        public static IOResult RemoveDirectoryOnPc(this PowerShell powerShell, string pcAddress, string path) =>
            powerShell.RemoveItemOnPc(pcAddress, path, true);

        public static Collection<PSDriveInfo> GetPSDrives(this PowerShell powerShell) =>
            powerShell.AddScript("Get-PSDrive").InvokeAndClear<PSDriveInfo>();

        public static PSDriveInfo GetPSDrive(this PowerShell powerShell, string name)
        {
            try
            {
                return powerShell.AddScript($"Get-PSDrive -Name {name}").InvokeAndClear<PSDriveInfo>().FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region AD Helpers

        private static ADComputer GetADComputerInternal(this PowerShell powerShell, string script)
        {
            if (!powerShell.IsADModuleAvailable())
                throw new Exception("Active Directory module is not available");

            try
            {
                return powerShell.AddScript(script).InvokeAndClear<ADComputer>().FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        public static ADComputer GetADComputer(this PowerShell powerShell, string computerName) =>
            powerShell.GetADComputerInternal($"Get-ADComputer -Identity {computerName} -Properties *");

        public static Collection<ADComputer> GetADComputers(this PowerShell powerShell, string filter)
        {
            if (!powerShell.IsADModuleAvailable())
                throw new Exception("Active Directory module is not available");

            try
            {
                return powerShell.AddScript($"Get-ADComputer -Filter '{filter}' -Properties *").InvokeAndClear<ADComputer>();
            }
            catch
            {
                return null;
            }
        }

        public static Collection<ADComputer> GetADUserComputers(this PowerShell powerShell) =>
            powerShell.GetADComputers("Name -like 'WKS*' -or Name -like 'VDI*'");

        public static Collection<ADComputer> GetADServers(this PowerShell powerShell) =>
            powerShell.GetADComputers("OperatingSystem -like '*Server*'");

        private static ADUser GetADUserInternal(this PowerShell powerShell, string script)
        {
            if (!powerShell.IsADModuleAvailable())
                throw new Exception("Active Directory module is not available");

            try
            {
                return powerShell.AddScript(script).InvokeAndClear<ADUser>().FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        public static ADUser GetADUser(this PowerShell powerShell, string username) =>
            powerShell.GetADUserInternal($"Get-ADUser -Identity {username} -Properties *");

        public static ADUser GetADUserBySSO(this PowerShell powerShell)
        {
            ADUser user = null;
            while (user == null)
            {
                string sso = Helpers.GetValidUserInput<string>("Enter the user's SSO: ");
                user = powerShell.GetADUser(sso);
                if (user == null)
                {
                    Console.Write("User not found. Press ENTER to try again.");
                    Console.ReadLine();
                }
            }
            return user;
        }

        public static Collection<ADUser> GetADUsers(this PowerShell powerShell, string filter)
        {
            if (!powerShell.IsADModuleAvailable())
                throw new Exception("Active Directory module is not available");

            try
            {
                return powerShell.AddScript($"Get-ADUser -Filter '{filter}' -Properties *").InvokeAndClear<ADUser>();
            }
            catch
            {
                return null;
            }
        }

        private static ADGroup GetADGroupInternal(this PowerShell powerShell, string script)
        {
            if (!powerShell.IsADModuleAvailable())
                throw new Exception("Active Directory module is not available");

            try
            {
                return powerShell.AddScript(script).InvokeAndClear<ADGroup>().FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        public static ADGroup GetADGroup(this PowerShell powerShell, string groupName) =>
            powerShell.GetADGroupInternal($"Get-ADGroup -Identity {groupName} -Properties *");

        public static Collection<ADGroup> GetADGroups(this PowerShell powerShell, string filter)
        {
            if (!powerShell.IsADModuleAvailable())
                throw new Exception("Active Directory module is not available");

            try
            {
                return powerShell.AddScript($"Get-ADGroup -Filter '{filter}' -Properties *").InvokeAndClear<ADGroup>();
            }
            catch
            {
                return null;
            }
        }

        private static List<DHCPScope> GetActiveDHCPScopes(this PowerShell powerShell, string DHCPServer)
        {
            var scopes = new List<DHCPScope>();
            try
            {
                var results = powerShell.AddScript($"Get-DhcpServerv4Scope -ComputerName {DHCPServer}").InvokeAndClear();
                foreach (var result in results)
                {
                    if (result.Properties["State"].Value.ToString() == "Active")
                    {
                        scopes.Add(new DHCPScope
                        {
                            ScopeId = result.Properties["ScopeId"].Value.ToString(),
                            SubnetMask = result.Properties["SubnetMask"].Value.ToString(),
                            Name = result.Properties["Name"].Value.ToString(),
                            State = result.Properties["State"].Value.ToString(),
                            StartRange = result.Properties["StartRange"].Value.ToString(),
                            EndRange = result.Properties["EndRange"].Value.ToString(),
                            LeaseDuration = result.Properties["LeaseDuration"].Value.ToString()
                        });
                    }
                }
                return scopes;
            }
            catch
            {
                return null;
            }
        }

        private static List<string> GetUserComputersInDHCPScope(this PowerShell powerShell, string DHCPServer, string scopeId)
        {
            var leases = new List<string>();
            try
            {
                var results = powerShell.AddScript($"Get-DhcpServerv4Lease -ComputerName {DHCPServer} -ScopeId {scopeId}").InvokeAndClear();
                foreach (var result in results)
                {
                    var hostname = result.Properties["HostName"].Value.ToString();
                    hostname = hostname.Substring(0, hostname.IndexOf('.'));
                    if (hostname.ToLower().Substring(0, 3) == "wks" || hostname.ToLower().Substring(0, 3) == "vdi")
                    {
                        leases.Add(hostname);
                    }
                }
                return leases;
            }
            catch
            {
                return null;
            }
        }

        private static void DisplayDHCPScopes(List<DHCPScope> dHCPScopes)
        {
            for (int i = 0; i < dHCPScopes.Count; i++)
            {
                Console.WriteLine($"{i + 1}: {dHCPScopes[i].Name}");
            }
        }

        private static int GetDHCPScopesSelection(List<DHCPScope> dHCPScopes)
        {
            int selection = 0;
            while (true)
            {
                DisplayDHCPScopes(dHCPScopes);
                selection = Helpers.GetValidUserInput<int>("Enter the number of the DHCP scope: ");
                if (selection < 1 || selection > dHCPScopes.Count)
                {
                    Console.Write("Invalid selection. Press ENTER to try again...");
                    Console.ReadLine();
                }
                else
                    break;
            }
            return selection;
        }

        public static List<string> GetComputerNamesFromDHCPScope(this PowerShell powerShell, string DHCPServer)
        {
            var computerNames = new List<string>();
            var scopes = powerShell.GetActiveDHCPScopes(DHCPServer);
            if (scopes == null)
                return null;

            scopes.Add(new DHCPScope { ScopeId = "All", Name = "All Computers" });

            do
            {
                int selection = GetDHCPScopesSelection(scopes);
                if (selection == scopes.Count)
                {
                    var allComputers = powerShell.GetADComputers("*");
                    foreach (var computer in allComputers)
                    {
                        computerNames.Add(computer.Name);
                    }
                }
                else
                {
                    computerNames = powerShell.GetUserComputersInDHCPScope(DHCPServer, scopes[selection - 1].ScopeId);
                }

                if (computerNames.Count < 1)
                {
                    Console.Write("No computers found in the selected DHCP scope. Press ENTER to try again.");
                    Console.ReadLine();
                }

            } while (computerNames.Count < 1);

            return computerNames;
        }

        #endregion
    }
}