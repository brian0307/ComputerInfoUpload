using System;
using System.IO;
using System.ServiceProcess;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ComputerInfoUpload
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            // --- Parse flags ---
            // Debug 日誌:
            //   computerInfoUpload.exe -D C:\Logs
            //   computerInfoUpload.exe --debug C:\Logs
            //
            // JSON 輸出:
            //   computerInfoUpload.exe -J C:\Desktop
            //   computerInfoUpload.exe --json C:\Desktop
            ConfigureFromArgs(args);

            try
            {
                ClientLog.Info("Program start.");
                try
                {
                    // 檢查 WMI
                    var service = new ServiceController("winmgmt");
                    if (service.Status != ServiceControllerStatus.Running)
                    {
                        Console.WriteLine("WMI 未啟動，嘗試啟動中…");
                        ClientLog.Warn("WMI not running. Attempting to start...");

                        // 若服務是停止狀態才啟動
                        if (service.Status == ServiceControllerStatus.Stopped ||
                            service.Status == ServiceControllerStatus.StopPending)
                        {
                            service.Start();
                            service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                        }

                        if (service.Status == ServiceControllerStatus.Running)
                        {
                            Console.WriteLine("WMI 已成功啟動\n");
                            ClientLog.Info("WMI started successfully.");
                        }
                        else
                        {
                            Console.WriteLine("WMI 啟動失敗或超時\n");
                            ClientLog.Warn("WMI start failed or timed out.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("WMI 已在執行中\n");
                        ClientLog.Info("WMI already running.");
                    }
                }
                catch (Exception ex)
                {
                    ClientLog.Error("WMI check/start failed.", ex);
                }

                // 收集電腦資訊
                var collector = new ComputerInfoCollector();
                var info = collector.CollectComputerInfo();

                // 序列化成 JSON
                string json = JsonSerializer.Serialize(info, new JsonSerializerOptions
                {
                    WriteIndented = true, // 美化資料 : 換行與縮排等等
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping // 支援中文
                });

                Console.WriteLine(" 電腦資訊 JSON:\n");
                Console.WriteLine(json);

                // 交給 JsonOutput 控管是否輸出檔案與路徑
                string? outputPath = await JsonOutput.WriteAsync(json);
                if (!string.IsNullOrEmpty(outputPath))
                {
                    Console.WriteLine($"\n 已輸出 JSON 檔案至：{outputPath}\n");
                }

                /* （可選）上傳到 API
                try
                {
                    await ApiUploader.UploadToApi("https://sample.com.tw/api/?pn=ItDeviceReceive", info);
                }
                catch (Exception ex)
                {
                    ClientLog.Error("API 上傳失敗。", ex);
                }
                */

                ClientLog.Info("Program finish.");
            }
            catch (Exception ex)
            {
                ClientLog.Error("Program fatal error.", ex);
                Console.WriteLine("錯誤：" + ex.Message);
            }
        }

        private static void ConfigureFromArgs(string[] args)
        {
            try
            {
                // -------- Debug 日誌設定 (-D / --debug) --------
                int dIdx = Array.FindIndex(args, a =>
                    string.Equals(a, "-D", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(a, "--debug", StringComparison.OrdinalIgnoreCase));

                if (dIdx >= 0)
                {
                    string? logPath = null;
                    if (dIdx + 1 < args.Length && !args[dIdx + 1].StartsWith("-"))
                    {
                        logPath = args[dIdx + 1];
                    }
                    if (string.IsNullOrWhiteSpace(logPath))
                    {
                        // default to .\Logs next to exe
                        logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                    }
                    ClientLog.Configure(true, logPath);
                    Console.WriteLine($"[Debug] Logging enabled to: {logPath}");
                }
                else
                {
                    // Default: logging disabled
                    ClientLog.Configure(false);
                }

                // -------- JSON 輸出設定 (-J / --json) --------
                int jIdx = Array.FindIndex(args, a =>
                    string.Equals(a, "-J", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(a, "--json", StringComparison.OrdinalIgnoreCase));

                if (jIdx >= 0)
                {
                    string? jsonDir = null;
                    if (jIdx + 1 < args.Length && !args[jIdx + 1].StartsWith("-"))
                    {
                        jsonDir = args[jIdx + 1];
                    }

                    if (!string.IsNullOrWhiteSpace(jsonDir))
                    {
                        JsonOutput.Configure(true, jsonDir);
                    }
                    else
                    {
                        // 若未指定路徑，預設用執行檔目錄
                        JsonOutput.Configure(true, AppDomain.CurrentDomain.BaseDirectory);
                    }

                    Console.WriteLine($"[JSON] Output enabled. Directory: {jsonDir ?? AppDomain.CurrentDomain.BaseDirectory}");
                    ClientLog.Info($"JSON output enabled. Directory: {jsonDir ?? AppDomain.CurrentDomain.BaseDirectory}");
                }
                else
                {
                    JsonOutput.Configure(false);
                }
            }
            catch (Exception ex)
            {
                ClientLog.Error("ConfigureFromArgs failed.", ex);
                ClientLog.Configure(false);
                JsonOutput.Configure(false);
            }
        }
    }
}
