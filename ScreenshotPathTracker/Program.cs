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

public class CompletedItem
{
    public string ItemName { get; set; } = string.Empty;
    public DateTime TimeStamp { get; set; } = DateTime.Now;
}

class Program
{
    static AppSettings? settings = null;
    static List<CompletedItem>? completedFolders = null;
    static string completedFoldersList = string.Empty;
    static List<CompletedItem>? completedFiles = null;
    static string completedFilesList = string.Empty;
    static string TessDataFolder = string.Empty;
    static string connectionString = string.Empty;

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

        // 接続文字列の作成
        connectionString = $"Server=192.168.11.15;Database=AOI;User Id={settings.DbUser};Password={settings.DbPassword};TrustServerCertificate=True;";

        // 処理済みフォルダリストの読み込み
        completedFoldersList = Path.Combine(ProgramDataFolder, "completedFoldersList.json");
        if (File.Exists(completedFoldersList))
        {
            completedFolders = JsonSerializer.Deserialize<List<CompletedItem>>(File.ReadAllText(completedFoldersList));
        }
        else
        {
            completedFolders = new List<CompletedItem>();
        }

        // 処理済みファイルリストの読み込み
        completedFilesList = Path.Combine(ProgramDataFolder, "completedFilesList.json");
        if (File.Exists(completedFilesList))
        {
            completedFiles = JsonSerializer.Deserialize<List<CompletedItem>>(File.ReadAllText(completedFilesList));
        }
        else
        {
            completedFiles = new List<CompletedItem>();
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
                {
                    LotFolders.Add(dir);
                }
            }
        }
        else
        {
            Console.WriteLine("TargetFolderが存在しません。");
        }

        // コマンドライン引数にファイル名が与えられた場合、標準出力とファイルの両方に出力
        if (args.Length > 0)
        {
            string logFilePath = args[0];
            using (StreamWriter writer = new StreamWriter(logFilePath))
            using (var multiWriter = new MultiTextWriter(Console.Out, writer))
            {
                Console.SetOut(multiWriter); // 標準出力をマルチライターに設定
                Run(LotFolders);
            }
        }
        else
        {
            Run(LotFolders);
        }

        // コマンドライン引数にファイル名が与えられた場合、標準出力をそのファイルにリダイレクト
        //if (args.Length > 0)
        //{
        //    string logFilePath = args[0];
        //    using (StreamWriter writer = new StreamWriter(logFilePath))
        //    {
        //        Console.SetOut(writer);
        //        Run(LotFolders);
        //    }
        //}
        //else
        //{
        //    Run(LotFolders);
        //}
    }

    static void Run(List<string> LotFolders)
    {
        // 開始時刻を表示
        Console.WriteLine("開始：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

        FindScreenshot(LotFolders);
        CsvToDatabase();

        // 終了時刻を表示
        Console.WriteLine("終了：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

    }

    static void CsvToDatabase()
    {
        if (settings == null)
        {
            return;
        }

        DateTime folderTimeStamp;

        // csvFolder内の全てのcsvファイルをリストアップ
        var csvFiles = Directory.GetFiles(settings.CsvFolder, "*.csv", SearchOption.AllDirectories);

        foreach (var csvFile in csvFiles)
        {
            if (ShouldExitLoop())
            {
                return;
            }

            // 処理対象フォルダを表示
            Console.Write(csvFile);

            if (completedFiles != null)
            {
                folderTimeStamp = File.GetLastWriteTime(csvFile);
                if (completedFiles.Exists(x => x.ItemName == csvFile && x.TimeStamp == folderTimeStamp))
                {
                    Console.WriteLine(" -> ファイルは処理済みです");
                    continue;
                }
            }

            // 改行出力
            Console.WriteLine();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // csvFileを1行ずつ読込んで処理
                using (var reader = new StreamReader(csvFile))
                {
                    // 1行目はヘッダー行なのでスキップする
                    _ = reader.ReadLine();

                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (line == null)
                        {
                            continue;
                        }

                        var values = line.Split(',');
                        Console.Write($"{values[4]}-{values[5]} -> ");

                        // レコードが存在するか確認
                        string selectQuery = "SELECT COUNT(*) FROM ResultAOI WHERE LOT = @lot AND SLOT = @slot";
                        using (SqlCommand selectCommand = new SqlCommand(selectQuery, connection))
                        {
                            selectCommand.Parameters.AddWithValue("@lot", values[4]);
                            selectCommand.Parameters.AddWithValue("@slot", values[5]);
                            int count = (int)selectCommand.ExecuteScalar();

                            // レコードが存在する場合、更新
                            string query = string.Empty;
                            if (count > 0)
                            {
                                query = "UPDATE ResultAOI SET Class020 = @class020, Class025 = @class025, Class029 = @class029, Class050 = @class050, Class104 = @class104, Class300 = @class300, Class500 = @class500, Total = @total, Judge = @judge, Comment = @comment WHERE LOT = @lot AND SLOT = @slot";
                                Console.WriteLine("更新");
                            }
                            else
                            {
                                query = "INSERT INTO ResultAOI (LOT, SLOT, Class020, Class025, Class029, Class050, Class104, Class300, Class500, Total, Judge, Comment) VALUES (@lot, @slot, @class020, @class025, @class029, @class050, @class104, @class300, @class500, @total, @judge, @comment)";
                                Console.WriteLine("新規登録");
                            }

                            using (SqlCommand sqlCommand = new SqlCommand(query, connection))
                            {
                                ExecuteSQL(sqlCommand, line);
                            }
                        }
                    }
                }
            }

            // folderのタイムスタンプを取得する
            folderTimeStamp = Directory.GetLastWriteTime(csvFile);

            if (completedFiles != null)
            {
                var completedItem = completedFiles.Find(x => x.ItemName == csvFile);
                if (completedItem != null)
                {
                    completedItem.TimeStamp = folderTimeStamp;
                }
                else
                {
                    completedFiles.Add(new CompletedItem { ItemName = csvFile, TimeStamp = folderTimeStamp });
                }

                var json = JsonSerializer.Serialize(completedFiles, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(completedFilesList, json);
            }
        }
    }

    static void ExecuteSQL(SqlCommand sqlCommand, string csvLine)
    {
        var values = csvLine.Split(',');
        sqlCommand.Parameters.AddWithValue("@class020", int.Parse(values[6]));
        sqlCommand.Parameters.AddWithValue("@class025", int.Parse(values[7]));
        sqlCommand.Parameters.AddWithValue("@class029", int.Parse(values[8]));
        sqlCommand.Parameters.AddWithValue("@class050", int.Parse(values[9]));
        sqlCommand.Parameters.AddWithValue("@class104", int.Parse(values[10]));
        sqlCommand.Parameters.AddWithValue("@class300", int.Parse(values[11]));
        sqlCommand.Parameters.AddWithValue("@class500", int.Parse(values[12]));
        sqlCommand.Parameters.AddWithValue("@total", int.Parse(values[13]));
        sqlCommand.Parameters.AddWithValue("@judge", values[14]);
        sqlCommand.Parameters.AddWithValue("@comment", values[15]);
        sqlCommand.Parameters.AddWithValue("@lot", values[4]);
        sqlCommand.Parameters.AddWithValue("@slot", values[5]);
        sqlCommand.ExecuteNonQuery();
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

            if (completedFolders != null)
            {
                folderTimeStamp = Directory.GetLastWriteTime(folder);
                if (completedFolders.Exists(x => x.ItemName == folder && x.TimeStamp == folderTimeStamp))
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

            if (completedFolders != null)
            {
                var completedItem = completedFolders.Find(x => x.ItemName == folder);
                if (completedItem != null)
                {
                    completedItem.TimeStamp = folderTimeStamp;
                }
                else
                {
                    completedFolders.Add(new CompletedItem { ItemName = folder, TimeStamp = folderTimeStamp });
                }

                var json = JsonSerializer.Serialize(completedFolders, new JsonSerializerOptions { WriteIndented = true });
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
