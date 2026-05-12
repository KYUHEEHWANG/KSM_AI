using System;
using System.IO;
using System.Web.Script.Serialization;

namespace KSM_SolidEdge
{
    public class AppConfig
    {
        // --- 설정 항목들 ---
        public int ReleaseMemoryInterval { get; set; } = 50;
        public int RestartSolidEdgeInterval { get; set; } = 100;
        public int GridClearInterval { get; set; } = 100;
        public int ImageSizeWidth { get; set; } = 1024;
        public int ImageSizeHeight { get; set; } = 1024;

        // --- 관리용 로직 ---
        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

        /// <summary>
        /// JSON 파일로부터 설정을 로드합니다. 파일이 없으면 기본값으로 생성합니다.
        /// </summary>
        public static AppConfig Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    var serializer = new JavaScriptSerializer();
                    return serializer.Deserialize<AppConfig>(json);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Config Load Error: {ex.Message}");
            }

            // 파일이 없거나 오류 발생 시 기본값 생성 및 저장
            AppConfig defaultConfig = new AppConfig();
            defaultConfig.Save();
            return defaultConfig;
        }

        /// <summary>
        /// 현재 설정을 JSON 파일로 저장합니다.
        /// </summary>
        public void Save()
        {
            try
            {
                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(this);
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Config Save Error: {ex.Message}");
            }
        }
    }
}