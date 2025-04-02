using System.Data.SqlClient;
using System.Drawing;
using System.Text.Json;
using Tesseract;
using Microsoft.Data.SqlClient;
using SqlConnection = Microsoft.Data.SqlClient.SqlConnection;
using SqlCommand = Microsoft.Data.SqlClient.SqlCommand;
using SqlDataReader = Microsoft.Data.SqlClient.SqlDataReader;


public class AppSettings
{
    public string ScreenshotFolder { get; set; } = string.Empty;
    public string CsvFolder { get; set; } = string.Empty;
    public string DbUser { get; set; } = string.Empty;
    public string DbPassword { get; set; } = string.Empty;

    public AppSettings()
    {
    }

    public static AppSettings LoadFromFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("設定ファイルが見つかりません。", filePath);
        }

        var json = File.ReadAllText(filePath);
        var settings = JsonSerializer.Deserialize<AppSettings>(json);
        if (settings == null)
        {
            throw new InvalidOperationException("設定ファイルの読み込みに失敗しました。");
        }
        return settings;
    }

    public void SaveToFile(string filePath)
    {
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(filePath, json);
    }
}

public class CompleatedItem
{
    public string ItemName { get; set; } = string.Empty;
    public DateTime TimeStamp { get; set; } = DateTime.Now;
}

class Program
{
    static AppSettings? settings = null;
    static List<CompleatedItem>? compleatedFolders = null;
    static string completedFoldersList = string.Empty;
    static string TessDataFolder = string.Empty;

    static void Main(string[] args)
    {
        string ProgramDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), AppDomain.CurrentDomain.FriendlyName);

        // 設定ファイルの読み込み
        string jsonSettingsFilePath = Path.Combine(ProgramDataFolder, "settings.json");
        if (File.Exists(jsonSettingsFilePath))
        {
            settings = AppSettings.LoadFromFile(jsonSettingsFilePath);
            Console.WriteLine("JSON設定ファイルから読み込みました。");
        }
        else
        {
            if (!Directory.Exists(ProgramDataFolder))
            {
                Directory.CreateDirectory(ProgramDataFolder);
            }

            settings = new AppSettings();
            settings.SaveToFile(jsonSettingsFilePath);
            Console.WriteLine("デフォルト設定を下記ファイルに保存しました。内容を確認して再実行！");
            Console.WriteLine(jsonSettingsFilePath);

            // アプリケーションを終了
            Environment.Exit(0);
        }

        // 処理済みフォルダリストの読み込み
        completedFoldersList = Path.Combine(ProgramDataFolder, "completedFoldersList.json");
        if (File.Exists(completedFoldersList))
        {
            compleatedFolders = JsonSerializer.Deserialize<List<CompleatedItem>>(File.ReadAllText(completedFoldersList));
        }
        else
        {
            compleatedFolders = new List<CompleatedItem>();
        }

        // Tesseractのデータフォルダを確認
        TessDataFolder = Path.Combine(ProgramDataFolder, "tessdata");
        if (!Directory.Exists(TessDataFolder))
        {
            Console.WriteLine("OCRに必要な下記フォルダが見つかりません。");
            Console.WriteLine(TessDataFolder);
            Environment.Exit(1);
        }

        // ロットフォルダのリストを取得
        List<string> LotFolders = new List<string>();
        if (Directory.Exists(settings.ScreenshotFolder))
        {
            var directories = Directory.GetDirectories(settings.ScreenshotFolder);
            foreach (var dir in directories)
            {
                if (GetDeepestDirectoryName(dir).Length == "910963".Length)
                    LotFolders.Add(dir);
            }
        }
        else
        {
            Console.WriteLine("TargetFolderが存在しません。");
        }

        // コマンドライン引数にファイル名が与えられた場合、標準出力をそのファイルにリダイレクト
        if (args.Length > 0)
        {
            string logFilePath = args[0];
            using (StreamWriter writer = new StreamWriter(logFilePath))
            {
                Console.SetOut(writer);
                Run(LotFolders);
            }
        }
        else
        {
            Run(LotFolders);
        }
    }

    static void Run(List<string> LotFolders)
    {
        // 開始時刻を表示
        Console.WriteLine("開始：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

        FindScreenshot(LotFolders);

        // 終了時刻を表示
        Console.WriteLine("終了：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

    }

    static string GetDeepestDirectoryName(string path)
    {
        return Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
    }

    static void FindScreenshot(List<string> Folders)
    {
        if (settings == null)
        {
            return;
        }

        string connectionString = $"Server=192.168.11.15;Database=AOI;User Id={settings.DbUser};Password={settings.DbPassword};TrustServerCertificate=True;";
        string PassOrFail = string.Empty;
        string ParentFolder = string.Empty;
        string ParentFolderTemp = string.Empty;
        string AoiJudgement = string.Empty;
        string OcrJudgement = string.Empty;
        (int w, int h) ImageDimensions = (0, 0);
        List<string>? fullPathList = null;
        DateTime folderTimeStamp = DateTime.MinValue;

        foreach (var folder in Folders)
        {
            if (ShouldExitLoop())
            {
                return;
            }

            // 処理対象フォルダを表示
            Console.Write(folder);

            if (compleatedFolders != null)
            {
                folderTimeStamp = Directory.GetLastWriteTime(folder);
                if (compleatedFolders.Exists(x => x.ItemName == folder && x.TimeStamp == folderTimeStamp))
                {
                    Console.WriteLine(" -> フォルダは処理済みです");
                    continue;
                }
            }

            // 改行出力
            Console.WriteLine();

            var directories = Directory.GetFiles(folder, "*.jpeg", SearchOption.AllDirectories);
            foreach (var file in directories)
            {
                if (ShouldExitLoop())
                {
                    return;
                }

                // AOIが含まれていないファイルはスキップ
                if (file.Contains("AOI") != true)
                {
                    continue;
                }

                PassOrFail = CheckPassOrFail(file);

                // 合格、不合格の文字列が含まれていないファイルはスキップ
                if (PassOrFail == string.Empty)
                {
                    continue;
                }

                fullPathList = GetLotInDatabase(connectionString, GetDeepestDirectoryName(folder));

                // ParentFolderが変わったときの処理
                ParentFolderTemp = Path.GetDirectoryName(file) ?? string.Empty;
                if (ParentFolder != ParentFolderTemp)
                {
                    ParentFolder = ParentFolderTemp;

                    // ParentFolderの中に拡張子がcsvのファイルがあるか確認
                    var csvFiles = Directory.GetFiles(ParentFolder, "*.csv", SearchOption.TopDirectoryOnly);
                    AoiJudgement = csvFiles.Length != 0 ? "OK" : "NG";
                }

                Console.Write(file.Replace(settings.ScreenshotFolder, "") + " ");

                bool exists = fullPathList.Contains(file);
                if (exists)
                {
                    Console.WriteLine("ファイル名はデータベースに存在します");
                }
                else
                {
                    Console.WriteLine("ファイル名をデータベースに登録します");

                    OcrJudgement = CountWordInRectangle("NG", file, new Rectangle(1130, 170, 50, 160)) == 0 ? "OK" : "NG";
                    ImageDimensions = GetImageDimensions(file);

                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        string query = "INSERT INTO ScreenshotList (タイトル, TimeStamp, ロット, 合否, AOI判定, 画像横サイズ, 画像縦サイズ, フルパス) VALUES (@Title, @TimeStamp, @Lot, @PassOrFail, @AoiJudgement, @Width, @Height, @filePath)";
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@Title", Path.GetFileName(folder));
                            command.Parameters.AddWithValue("@TimeStamp", File.GetCreationTime(file));
                            command.Parameters.AddWithValue("@Lot", GetDeepestDirectoryName(folder));
                            command.Parameters.AddWithValue("@PassOrFail", PassOrFail);
                            command.Parameters.AddWithValue("@AoiJudgement", $"{AoiJudgement}|{OcrJudgement}");
                            command.Parameters.AddWithValue("@Width", ImageDimensions.w);
                            command.Parameters.AddWithValue("@Height", ImageDimensions.h);
                            command.Parameters.AddWithValue("@filePath", file);
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }

            // folderのタイムスタンプを取得する
            folderTimeStamp = Directory.GetLastWriteTime(folder);

            if (compleatedFolders != null)
            {
                compleatedFolders.Add(new CompleatedItem { ItemName = folder, TimeStamp = folderTimeStamp });

                var json = JsonSerializer.Serialize(compleatedFolders, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(completedFoldersList, json);
            }
        }
    }

    static bool ShouldExitLoop()
    {
        if (Console.KeyAvailable)
        {
            Console.ReadKey(true); // キー入力を処理
            return true;
        }
        return false;
    }

    public static List<string> GetLotInDatabase(string connectionString, string lotName)
    {
        using (SqlConnection connection = new(connectionString))
        {
            connection.Open();
            string query = "SELECT フルパス FROM ScreenshotList WHERE ロット = @lotName";
            using (SqlCommand command = new(query, connection))
            {
                command.Parameters.AddWithValue("@lotName", lotName);
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    List<string> list = new List<string>();
                    while (reader.Read())
                    {
                        list.Add(reader.GetString(0));
                    }
                    return list;
                }
            }
        }
    }

    public static (int width, int height) GetImageDimensions(string filePath)
    {
        if (OperatingSystem.IsWindows())
        {
            using (var image = System.Drawing.Image.FromFile(filePath)) // Use System.Drawing.Image
            {
                return (image.Width, image.Height);
            }
        }
        else
        {
            throw new PlatformNotSupportedException("この機能はWindowsでのみサポートされています。");
        }
    }

    public static string CheckPassOrFail(string path)
    {
        if (path.Contains("合格"))
        {
            return "合格";
        }
        else if (path.Contains("不合格"))
        {
            return "不合格";
        }
        else
        {
            return string.Empty;
        }
    }
    public static int CountWordInRectangle(string word, string imagePath, Rectangle rectangle)
    {
        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException("画像ファイルが見つかりません。", imagePath);
        }

        using (var image = System.Drawing.Image.FromFile(imagePath))
        using (var bitmap = new Bitmap(image))
        {
            // 矩形領域を切り取る
            var croppedBitmap = bitmap.Clone(rectangle, bitmap.PixelFormat);

            // BitmapをPixに変換
            using (var pix = ConvertBitmapToPix(croppedBitmap))
            {
                // OCRエンジンを使用して文字列を認識する
                using (var ocrEngine = new TesseractEngine(TessDataFolder, "jpn", EngineMode.Default))
                {
                    using (var page = ocrEngine.Process(pix))
                    {
                        var text = page.GetText();

                        // wordの出現回数を数える
                        int count = text.Split(new string[] { word }, StringSplitOptions.None).Length - 1;
                        return count;
                    }
                }
            }
        }
    }
    private static Pix ConvertBitmapToPix(Bitmap bitmap)
    {
        if (OperatingSystem.IsWindows())
        {
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);
                stream.Position = 0;
                return Pix.LoadFromMemory(stream.ToArray());
            }
        }
        else
        {
            throw new PlatformNotSupportedException("この機能はWindowsでのみサポートされています。");
        }
    }
}
