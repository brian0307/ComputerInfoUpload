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
            // Usage: computerInfoUpload.exe -D C:\Desktop  (enable debug logging to that folder)
            //        computerInfoUpload.exe --debug C:\Logs
            // If -D/--debug present but path missing, will use ".\Logs" next to exe.
            ConfigureLoggingFromArgs(args);

            try
            {
                ClientLog.Info("Program start.");
                try
                {
                    // 檢查 WMI
                    var service = new ServiceController("winmgmt");
                    if (service.Status != ServiceControllerStatus.Running)
                    {
                        //Console.WriteLine("WMI 未啟動，嘗試啟動中…");
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
                            //Console.WriteLine("WMI 已成功啟動\n");
                            ClientLog.Info("WMI started successfully.");
                        }
                        else
                        {
                            //Console.WriteLine("WMI 啟動失敗或超時\n");
                            ClientLog.Warn("WMI start failed or timed out.");
                        }
                    }
                    else
                    {
                        //Console.WriteLine("WMI 已在執行中\n");
                        ClientLog.Info("WMI already running.");
                    }
                }
                catch (Exception ex)
                {
                    ClientLog.Error("WMI check/start failed.", ex);
                }

                /* 建立自動啟動 .bat
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                string batPath = Path.Combine(startupPath, "MyAutoRun.bat");

                if (!File.Exists(batPath))
                {
                    File.WriteAllText(batPath, $"start \"\" \"{exePath}\"");
                    Console.WriteLine("已建立自動啟動 .bat 檔案於 Startup 資料夾。\n");
                }*/

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

                /* 將 JSON 寫入檔案
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "computerInfo.json");
                await File.WriteAllTextAsync(outputPath, json, Encoding.UTF8);
                Console.WriteLine($"\n 已輸出 JSON 檔案至：{outputPath}\n");
                */

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
                //Console.WriteLine("錯誤：" + ex.Message);
            }
        }

        private static void ConfigureLoggingFromArgs(string[] args)
        {
            try
            {
                // Find -D or --debug
                int idx = Array.FindIndex(args, a => string.Equals(a, "-D", StringComparison.OrdinalIgnoreCase)
                                                  || string.Equals(a, "--debug", StringComparison.OrdinalIgnoreCase));
                if (idx >= 0)
                {
                    string? path = null;
                    if (idx + 1 < args.Length && !args[idx + 1].StartsWith("-"))
                    {
                        path = args[idx + 1];
                    }
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        // default to .\Logs next to exe
                        path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                    }
                    ClientLog.Configure(true, path);
                    //Console.WriteLine($"[Debug] Logging enabled to: {path}");
                }
                else
                {
                    // Default: logging disabled
                    ClientLog.Configure(false);
                }
            }
            catch
            {
                ClientLog.Configure(false);
            }
        }
    }
}
