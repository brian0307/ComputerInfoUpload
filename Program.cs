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

            try
            {
                // 檢查 WMI
                var service = new ServiceController("winmgmt");
                if (service.Status != ServiceControllerStatus.Running)
                {
                    Console.WriteLine("WMI 未啟動，嘗試啟動中…");

                    // 若服務是停止狀態才啟動
                    if (service.Status == ServiceControllerStatus.Stopped ||
                        service.Status == ServiceControllerStatus.StopPending)
                    {
                        service.Start();
                        service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                    }

                    if (service.Status == ServiceControllerStatus.Running)
                        Console.WriteLine("WMI 已成功啟動\n");
                    else
                        Console.WriteLine("WMI 啟動失敗或超時\n");
                }
                else
                {
                    Console.WriteLine("WMI 已在執行中\n");
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

                // （可選）上傳到 API
                //await ApiUploader.UploadToApi("https://tap.turvo.com.tw/api/?pn=ItDeviceReceive", info);
            }
            catch (Exception ex)
            {
                Console.WriteLine("錯誤：" + ex.Message);
            }
        }
    }
}
