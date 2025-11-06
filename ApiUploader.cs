using ComputerInfoUpload;
using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ComputerInfoUpload
{
    public class ApiUploader
    {
        public static async Task UploadToApi(string url, object data)
        {
            try
            {
                using var httpClient = new HttpClient();
                string json = JsonSerializer.Serialize(data, new JsonSerializerOptions
                {
                    WriteIndented = false,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping // 支援中文
                });
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync(url, content);
                Console.WriteLine($"API Response: {response.StatusCode}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("API 上傳失敗：" + ex.Message);
            }
        }
    }
}