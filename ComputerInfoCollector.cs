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
using System.Windows.Forms;

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

            // 顯示器/螢幕解析度
            info["Displays"] = GetDisplayInfos();

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

            // 記憶體（插槽、DDR4/DDR5種類、模組清單）
            info["Memory"] = GetMemorySummary();

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
                ClientLog.Warn("GetWMIValue failed" + ex.Message);
            }
            return "Null";
        }

        private static (List<string>, List<string>) GetNetworkInfo()
        {
            var ips = new List<string>();
            var macs = new List<string>();

            try
            {
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
            }
            catch (Exception ex)
            {
                ClientLog.Warn("GetNetworkInfo failed: " + ex.Message);
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

            try
            {
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
            }
            catch (Exception ex)
            {
                ClientLog.Warn("GetInstalledSoftware failed: " + ex.Message);
            }

            return software;
        }

        private static void GetDrivesInfo(Dictionary<string, object> info)
        {
            var driveNames = new List<string>();
            var driveMemories = new List<double>();
            var driveAvailableMemories = new List<double>();

            try
            {
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
            catch (Exception ex)
            {
                ClientLog.Warn("GetDrivesInfo failed: " + ex.Message);
            }
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
            catch (Exception ex)
            {
                ClientLog.Warn("GetTotalRAM failed: " + ex.Message);
            }
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
            catch (Exception ex) 
            {
                ClientLog.Warn("FormatWMIDate failed: " + ex.Message);
            }
            return "Invalid";
        }

        private static string GetCpuId()
        {
            string cpuId = string.Empty;

            try
            {
                // Create a ManagementObjectSearcher to query WMI
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("select ProcessorId from Win32_Processor"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        cpuId = obj["ProcessorId"]?.ToString() ?? "Unknown";
                        break; // typically only one processor entry
                    }
                }
            }
            catch (Exception ex)
            {
                ClientLog.Warn("GetCpuId failed: " + ex.Message);
            }

            return cpuId;
        }

        private static List<object> GetDisplayInfos()
        {
            var list = new List<object>();

            try
            {
                // 先抓到每個螢幕的基本資訊
                foreach (var screen in Screen.AllScreens)
                {
                    var bounds = screen.Bounds;
                    list.Add(new
                    {
                        DeviceName = screen.DeviceName,
                        Width = bounds.Width,
                        Height = bounds.Height,
                        Primary = screen.Primary,
                        Manufacturer = GetMonitorManufacturer(screen.DeviceName),
                        MonitorModel = GetMonitorModel(screen.DeviceName)
                    });
                }

                // 若抓不到（理論上不會發生）
                if (list.Count == 0)
                {
                    using var searcher = new ManagementObjectSearcher(
                        "SELECT Name, CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController");
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        list.Add(new
                        {
                            DeviceName = obj["Name"]?.ToString() ?? "Display",
                            Width = Convert.ToInt32(obj["CurrentHorizontalResolution"] ?? 0),
                            Height = Convert.ToInt32(obj["CurrentVerticalResolution"] ?? 0),
                            Primary = (bool?)null,
                            Manufacturer = string.Empty,
                            MonitorModel = string.Empty
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                ClientLog.Warn("GetDisplayInfos failed: " + ex.Message);
            }

            return list;
        }

        private static string GetMonitorManufacturer(string deviceName)
        {
            try
            {
                using var searcher = new ManagementObjectSearcher("root\\wmi", "SELECT ManufacturerName FROM WmiMonitorID");
                foreach (ManagementObject mo in searcher.Get())
                {
                    var arr = mo["ManufacturerName"] as ushort[];
                    if (arr != null)
                    {
                        string manuf = new string(arr.Select(c => (char)c).ToArray()).TrimEnd('\0');
                        if (!string.IsNullOrWhiteSpace(manuf))
                            return manuf;
                    }
                }
            }
            catch (Exception ex)
            {
                ClientLog.Warn("GetMonitorManufacturer failed: " + ex.Message);
            }
            return string.Empty;
        }

        private static string GetMonitorModel(string deviceName)
        {
            try
            {
                using var searcher = new ManagementObjectSearcher("root\\wmi", "SELECT UserFriendlyName FROM WmiMonitorID");
                foreach (ManagementObject mo in searcher.Get())
                {
                    var arr = mo["UserFriendlyName"] as ushort[];
                    if (arr != null)
                    {
                        string model = new string(arr.Select(c => (char)c).ToArray()).TrimEnd('\0');
                        if (!string.IsNullOrWhiteSpace(model))
                            return model;
                    }
                }
            }
            catch (Exception ex)
            {
                ClientLog.Warn("GetMonitorModel failed: " + ex.Message);
            }
            return string.Empty;
        }

        private static object GetMemorySummary()
        {
            var modules = new List<object>();
            int installedCount = 0;
            int maxSlots = 0;
            ulong totalBytes = 0;
            var types = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Max slots from PhysicalMemoryArray.MemoryDevices
                using (var arrSearcher = new ManagementObjectSearcher("SELECT MemoryDevices FROM Win32_PhysicalMemoryArray"))
                {
                    foreach (ManagementObject obj in arrSearcher.Get())
                    {
                        if (obj["MemoryDevices"] != null && int.TryParse(obj["MemoryDevices"].ToString(), out var md))
                        {
                            if (md > maxSlots) maxSlots = md;
                        }
                    }
                }

                using (var memSearcher = new ManagementObjectSearcher("SELECT Capacity,Speed,Manufacturer,PartNumber,BankLabel,DeviceLocator,SMBIOSMemoryType FROM Win32_PhysicalMemory"))
                {
                    foreach (ManagementObject mo in memSearcher.Get())
                    {
                        installedCount++;
                        ulong cap = 0;
                        if (mo["Capacity"] != null) ulong.TryParse(mo["Capacity"].ToString(), out cap);
                        totalBytes += cap;

                        int speed = mo["Speed"] != null ? Convert.ToInt32(mo["Speed"]) : 0;
                        string manufacturer = mo["Manufacturer"]?.ToString() ?? string.Empty;
                        string pn = mo["PartNumber"]?.ToString() ?? string.Empty;
                        string bank = mo["BankLabel"]?.ToString() ?? string.Empty;
                        string locator = mo["DeviceLocator"]?.ToString() ?? string.Empty;

                        string ddrType = MapSmbiosMemoryType(mo["SMBIOSMemoryType"] as IConvertible);
                        if (!string.IsNullOrEmpty(ddrType)) types.Add(ddrType);

                        modules.Add(new
                        {
                            SizeGB = Math.Round(cap / 1073741824.0, 2),
                            SpeedMHz = speed,
                            Type = ddrType,
                            Manufacturer = manufacturer,
                            PartNumber = pn,
                            BankLabel = bank,
                            DeviceLocator = locator
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                ClientLog.Warn("GetMemorySummary failed: " + ex.Message);
            }

            return new
            {
                InstalledCount = installedCount,
                MaxSlots = maxSlots, // 0 表示無法偵測
                TotalGB = Math.Round(totalBytes / 1073741824.0, 2),
                Types = types.ToArray(), // e.g. ["DDR4"] or ["DDR4","DDR5"]
                Modules = modules
            };
        }

        private static string MapSmbiosMemoryType(IConvertible? raw)
        {
            try
            {
                if (raw == null) return string.Empty;
                int code = raw.ToInt32(System.Globalization.CultureInfo.InvariantCulture);
                // Common SMBIOSMemoryType codes:
                // 24=DDR3, 26=DDR4, 34=DDR5
                return code switch
                {
                    24 => "DDR3",
                    26 => "DDR4",
                    34 => "DDR5",
                    _ => string.Empty
                };
            }
            catch { return string.Empty; }
        }
    }
}
