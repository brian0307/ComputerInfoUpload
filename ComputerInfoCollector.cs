using ComputerInfoUpload;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Text;

namespace ComputerInfoUpload
{
    public class ComputerInfoCollector
    {
        public Dictionary<string, object> CollectComputerInfo()
        {
            var info = new Dictionary<string, object>();
            Console.OutputEncoding = Encoding.UTF8;

            // 電腦與系統資訊
            info["CpuId"] = GetCpuId();
            info["ComputerName"] = Environment.MachineName;
            info["OSVersion"] = Environment.OSVersion.VersionString;
            info["BIOSSerialNumber"] = GetWMIValue("Win32_BIOS", "SerialNumber");
            info["Cpu"] = GetWMIValue("Win32_Processor", "Name");
            info["CpuCoreNumbers"] = Environment.ProcessorCount;
            info["Motherboard"] = new
            {
                Manufacturer = GetWMIValue("Win32_BaseBoard", "Manufacturer"),
                Product = GetWMIValue("Win32_BaseBoard", "Product"),
                SerialNumber = GetWMIValue("Win32_BaseBoard", "SerialNumber")
            };
            info["Gpu"] = GetWMIValue("Win32_VideoController", "Name");
            info["SystemArchitecture"] = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit";
            var rawBootTime = GetWMIValue("Win32_OperatingSystem", "LastBootUpTime");
            info["LastBootUpTime"] = FormatWMIDate(rawBootTime);

            // 網路資訊
            var (ip, mac) = GetNetworkInfo();
            info["IpAddress"] = ip;
            info["MacAddress"] = mac;

            // 使用者資訊
            info["UserName"] = Environment.UserName;

            // 已安裝軟體
            info["InstalledSoftware"] = GetInstalledSoftware();

            // 硬碟
            GetDrivesInfo(info);

            // 記憶體
            info["TotalRAM"] = GetTotalRAM();

            var now = DateTime.Now;
            info["UploadTime"] = now.ToString("yyyy/MM/dd HH:mm:ss");
            return info;
        }

        private static string GetWMIValue(string className, string propertyName, bool convertToGB = false)
        {
            try
            {
                using var searcher = new ManagementObjectSearcher($"SELECT {propertyName} FROM {className}");
                foreach (ManagementObject obj in searcher.Get())
                {
                    var value = obj[propertyName];
                    if (value != null)
                    {
                        if (convertToGB && ulong.TryParse(value.ToString(), out ulong bytes))
                            return (bytes / 1024 / 1024 / 1024).ToString();
                        return value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return "Null";
        }

        private static (List<string>, List<string>) GetNetworkInfo()
        {
            var ips = new List<string>();
            var macs = new List<string>();

            foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (ni.OperationalStatus == OperationalStatus.Up &&
                    ni.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                {
                    var ip = ni.GetIPProperties().UnicastAddresses
                        .FirstOrDefault(a => a.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);

                    if (ip != null)
                    {
                        ips.Add(ip.Address.ToString());
                        macs.Add(BitConverter.ToString(ni.GetPhysicalAddress().GetAddressBytes()));
                    }
                }
            }
            return (ips, macs);
        }

        private static List<string> GetInstalledSoftware()
        {
            var software = new List<string>();
            string[] uninstallKeys = {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (string key in uninstallKeys)
            {
                using RegistryKey rk = Registry.LocalMachine.OpenSubKey(key);
                if (rk == null) continue;

                foreach (string subKeyName in rk.GetSubKeyNames())
                {
                    using RegistryKey subKey = rk.OpenSubKey(subKeyName);
                    object displayName = subKey?.GetValue("DisplayName");
                    if (displayName != null)
                        software.Add(displayName.ToString());
                }
            }
            return software;
        }

        private static void GetDrivesInfo(Dictionary<string, object> info)
        {
            var driveNames = new List<string>();
            var driveMemories = new List<double>();
            var driveAvailableMemories = new List<double>();

            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady)
                {
                    driveNames.Add(drive.Name);
                    driveMemories.Add(drive.TotalSize / (1024.0 * 1024 * 1024));
                    driveAvailableMemories.Add(drive.AvailableFreeSpace / (1024.0 * 1024 * 1024));
                }
            }

            info["DriveName"] = driveNames;
            info["DriveMemory"] = driveMemories;
            info["DriveAvailableMemory"] = driveAvailableMemories;
        }

        private static string GetTotalRAM()
        {
            try
            {
                var psi = new ProcessStartInfo("wmic", "computersystem get TotalPhysicalMemory")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var process = Process.Start(psi);
                string output = process.StandardOutput.ReadToEnd();
                var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length >= 2 && ulong.TryParse(lines[1].Trim(), out ulong bytes))
                {
                    return (bytes / 1073741824.0).ToString("F2") + "GB";
                }
            }
            catch { }
            return "-1";
        }
        private static string FormatWMIDate(string wmiDate)
        {
            if (string.IsNullOrEmpty(wmiDate)) return "Unknown";

            try
            {
                // 截掉微秒與時區，只取前 14 碼
                string baseTime = wmiDate.Substring(0, 14);
                if (DateTime.TryParseExact(baseTime, "yyyyMMddHHmmss",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out DateTime parsed))
                {
                    return parsed.ToString("yyyy/MM/dd HH:mm:ss");
                }
            }
            catch { }
            return "Invalid";
        }

        private static string GetCpuId()
        {
            string cpuId = string.Empty;

            // Create a ManagementObjectSearcher to query WMI
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("select ProcessorId from Win32_Processor"))
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    cpuId = obj["ProcessorId"]?.ToString() ?? "Unknown";
                    break; // typically only one processor entry
                }
            }
            return cpuId;
        }
    }
}
