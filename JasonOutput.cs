using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ComputerInfoUpload
{
    /// <summary>
    /// 控管 JSON 輸出行為：
    /// - 預設關閉，不會輸出任何檔案。
    /// - 透過 Configure(enable, dir) 啟用與設定輸出資料夾。
    /// - WriteAsync(json) 會在啟用時將 JSON 寫入檔案。
    /// </summary>
    public static class JsonOutput
    {
        private static bool _enabled = false;
        private static string _directory = AppDomain.CurrentDomain.BaseDirectory;
        private const string DefaultFileName = "computerInfo.json";

        /// <summary>
        /// 設定是否啟用 JSON 輸出與輸出路徑。
        /// enable = true 時若有指定 directory 會更新目錄。
        /// </summary>
        public static void Configure(bool enable, string? directory = null)
        {
            try
            {
                _enabled = enable;

                if (!string.IsNullOrWhiteSpace(directory))
                {
                    _directory = directory!;
                }
                else
                {
                    _directory = AppDomain.CurrentDomain.BaseDirectory;
                }
            }
            catch (Exception ex)
            {
                _enabled = false;
                ClientLog.Error("JsonOutput.Configure failed.", ex);
            }
        }

        /// <summary>
        /// 將 JSON 寫入檔案（當 _enabled = true 時）。
        /// 傳回實際寫入的檔案完整路徑（失敗或未啟用時傳回 null）。
        /// </summary>
        public static async Task<string?> WriteAsync(string json, string? fileName = null)
        {
            if (!_enabled)
            {
                return null;
            }

            try
            {
                Directory.CreateDirectory(_directory);

                string name = string.IsNullOrWhiteSpace(fileName) ? DefaultFileName : fileName!;
                string outputPath = Path.Combine(_directory, name);

                await File.WriteAllTextAsync(outputPath, json, Encoding.UTF8);
                ClientLog.Info($"JSON written to {outputPath}");

                return outputPath;
            }
            catch (Exception ex)
            {
                ClientLog.Error("JsonOutput.WriteAsync failed.", ex);
                return null;
            }
        }
    }
}