using Aspose.Cells;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Newtonsoft.Json;
using NLog;

namespace wp_boot
{
  
    class Program
    {
        static void Main(string[] args)
        {
            var app = new FileProcessorApp();
            app.Run();
        }
    }

    public class FileProcessorApp
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Config _config = new Config();
        private readonly IInventoryProcessor _inventoryProcessor;
        private readonly ITextSummaryProcessor _textSummaryProcessor;
        private readonly IFileConverter _fileConverter;
        private readonly IBatchProcessor _batchProcessor;
        private readonly ISearchProcessor _searchProcessor;
        private readonly IReportGenerator _reportGenerator;
        private readonly IBackupService _backupService;

        public FileProcessorApp()
        {
            _inventoryProcessor = new InventoryProcessor(_config);
            _textSummaryProcessor = new TextSummaryProcessor();
            _fileConverter = new FileConverter();
            _batchProcessor = new BatchProcessor(_inventoryProcessor, _textSummaryProcessor, _fileConverter);
            _searchProcessor = new SearchProcessor();
            _reportGenerator = new ReportGenerator();
            _backupService = new BackupService();

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
        }

        public void Run()
        {
            while (true)
            {
                ShowMainMenu();
                string choice = Console.ReadLine();

                try
                {
                    ProcessChoice(choice);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Hata oluştu");
                    ShowError($"Hata: {ex.Message}");
                }

                Console.WriteLine("\nDevam etmek için bir tuşa basın...");
                Console.ReadKey();
            }
        }

        private void ShowMainMenu()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("=== DOSYA Düzenleyici ===");
            Console.ResetColor();
            Console.WriteLine("1. Envanter Kontrolü");
            Console.WriteLine("2. Metin Özeti");
            Console.WriteLine("3. Dosya Dönüştürme");
            Console.WriteLine("4. Toplu İşlem");
            Console.WriteLine("5. Arama ve Filtreleme");
            Console.WriteLine("6. Rapor Oluştur");
            Console.WriteLine("7. Yedekle");
            Console.WriteLine("0. Çıkış");
            Console.Write("Seçiminiz: ");
        }

        private void ProcessChoice(string choice)
        {
            switch (choice)
            {
                case "1":
                    ProcessInventory();
                    break;
                case "2":
                    ProcessTextSummary();
                    break;
                case "3":
                    ProcessConversion();
                    break;
                case "4":
                    ProcessBatch();
                    break;
                case "5":
                    ProcessSearch();
                    break;
                case "6":
                    ProcessReport();
                    break;
                case "7":
                    ProcessBackup();
                    break;
                case "0":
                    Environment.Exit(0);
                    break;
                default:
                    ShowError("Geçersiz seçim!");
                    break;
            }
        }

        private void ProcessInventory()
        {
            Console.Write("Dosya yolunu girin: ");
            string filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                ShowError("Dosya bulunamadı!");
                return;
            }

            Console.WriteLine("\n1. Mükerrer kontrol");
            Console.WriteLine("2. Stok analizi");
            Console.WriteLine("3. Eksik veri kontrolü");
            Console.Write("Seçiminiz: ");
            string subChoice = Console.ReadLine();

            switch (subChoice)
            {
                case "1":
                    _inventoryProcessor.CheckDuplicates(filePath);
                    break;
                case "2":
                    _inventoryProcessor.AnalyzeInventory(filePath);
                    break;
                case "3":
                    _inventoryProcessor.CheckMissingData(filePath);
                    break;
                default:
                    ShowError("Geçersiz seçim!");
                    break;
            }
        }

        private void ProcessTextSummary()
        {
            Console.Write("Dosya yolunu girin: ");
            string filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                ShowError("Dosya bulunamadı!");
                return;
            }

            string summary = _textSummaryProcessor.GenerateSummary(filePath);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nMetin Özeti:");
            Console.ResetColor();
            Console.WriteLine(summary);

            string summaryPath = Path.ChangeExtension(filePath, "_summary.txt");
            File.WriteAllText(summaryPath, summary);
            Console.WriteLine($"\nÖzet kaydedildi: {summaryPath}");
        }

        private void ProcessConversion()
        {
            Console.Write("Kaynak dosya yolunu girin: ");
            string sourcePath = Console.ReadLine();

            if (!File.Exists(sourcePath))
            {
                ShowError("Dosya bulunamadı!");
                return;
            }

            Console.WriteLine("\n1. Excel -> PDF");
            Console.WriteLine("2. Excel -> DOCX");
            Console.WriteLine("3. Excel -> PPTX");
            Console.Write("Seçiminiz: ");
            string choice = Console.ReadLine();

            Console.Write("Sayfa numarası (tüm sayfalar için boş bırakın): ");
            string sheetInput = Console.ReadLine();
            int sheetNumber = string.IsNullOrWhiteSpace(sheetInput) ? -1 : int.Parse(sheetInput);

            string outputPath = Path.ChangeExtension(sourcePath, choice == "1" ? ".pdf" : choice == "2" ? ".docx" : ".pptx");

            _fileConverter.ConvertFile(sourcePath, outputPath, choice, sheetNumber);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\nDönüştürme tamamlandı: {outputPath}");
            Console.ResetColor();
        }

        private void ProcessBatch()
        {
            Console.Write("Klasör yolunu girin: ");
            string folderPath = Console.ReadLine();

            if (!Directory.Exists(folderPath))
            {
                ShowError("Klasör bulunamadı!");
                return;
            }

            Console.WriteLine("\n1. Tüm Excel dosyalarını PDF'ye dönüştür");
            Console.WriteLine("2. Tüm metin dosyalarının özetini oluştur");
            Console.Write("Seçiminiz: ");
            string choice = Console.ReadLine();

            _batchProcessor.ProcessBatch(folderPath, choice);
        }

        private void ProcessSearch()
        {
            Console.Write("Dosya yolunu girin: ");
            string filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                ShowError("Dosya bulunamadı!");
                return;
            }

            Console.Write("Aranacak metin: ");
            string searchTerm = Console.ReadLine();

            var results = _searchProcessor.SearchInFile(filePath, searchTerm);

            if (results.Any())
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\n{results.Count} sonuç bulundu:");
                Console.ResetColor();

                foreach (var result in results)
                {
                    Console.WriteLine($"- Satır {result.Line}: {result.Text}");
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\nSonuç bulunamadı");
                Console.ResetColor();
            }
        }

        private void ProcessReport()
        {
            Console.Write("Dosya yolunu girin: ");
            string filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                ShowError("Dosya bulunamadı!");
                return;
            }

            Console.WriteLine("\n1. Envanter raporu");
            Console.WriteLine("2. Özet raporu");
            Console.Write("Seçiminiz: ");
            string choice = Console.ReadLine();

            string reportPath = Path.ChangeExtension(filePath, "_rapor.html");

            if (choice == "1")
            {
                var inventory = _inventoryProcessor.LoadInventory(filePath);
                _reportGenerator.GenerateReport(inventory, "Envanter Raporu", reportPath);
            }
            else if (choice == "2")
            {
                string summary = _textSummaryProcessor.GenerateSummary(filePath);
                _reportGenerator.GenerateReport(summary, "Metin Özeti Raporu", reportPath);
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\nRapor oluşturuldu: {reportPath}");
            Console.ResetColor();
        }

        private void ProcessBackup()
        {
            Console.Write("Yedeklenecek dosya yolunu girin: ");
            string filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                ShowError("Dosya bulunamadı!");
                return;
            }

            string backupPath = _backupService.CreateBackup(filePath);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\nYedek oluşturuldu: {backupPath}");
            Console.ResetColor();
        }

        private void ShowError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n{message}");
            Console.ResetColor();
        }
    }

    // Yapılandırma 
    public class Config
    {
        public int LowStockThreshold { get; set; } = 10;
        public int CriticalStockThreshold { get; set; } = 5;
        public string DefaultOutputPath { get; set; } = "Cikti";
        public bool EnableLogging { get; set; } = true;
    }

    // Envanter 
    public interface IInventoryProcessor
    {
        void CheckDuplicates(string filePath);
        void AnalyzeInventory(string filePath);
        void CheckMissingData(string filePath);
        List<InventoryItem> LoadInventory(string filePath);
    }

    // Envanter 
    public class InventoryProcessor : IInventoryProcessor
    {
        private readonly Config _config;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public InventoryProcessor(Config config)
        {
            _config = config;
        }

        public void CheckDuplicates(string filePath)
        {
            Console.WriteLine("Mükerrer kontrol yapılıyor...");
            var duplicates = FindDuplicates(filePath);

            if (duplicates.Any())
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n{duplicates.Count} adet mükerrer kayıt bulundu:");
                Console.ResetColor();

                foreach (var item in duplicates)
                {
                    Console.WriteLine($"- {item}");
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\nMükerrer kayıt bulunamadı!");
                Console.ResetColor();
            }
        }

        public void AnalyzeInventory(string filePath)
        {
            Console.WriteLine("Stok analizi yapılıyor...");
            var inventory = LoadInventory(filePath);

            if (!inventory.Any())
            {
                Console.WriteLine("Envanter verisi bulunamadı!");
                return;
            }

            var lowStock = inventory.Where(i => i.Quantity < _config.LowStockThreshold).ToList();
            var criticalStock = inventory.Where(i => i.Quantity < _config.CriticalStockThreshold).ToList();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"\nDüşük stok ({_config.LowStockThreshold} adet altında): {lowStock.Count} ürün");
            Console.ResetColor();

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Kritik stok ({_config.CriticalStockThreshold} adet altında): {criticalStock.Count} ürün");
            Console.ResetColor();

            // Basit metin tabanlı grafik
            GenerateTextStockChart(inventory);
        }

        public void CheckMissingData(string filePath)
        {
            Console.WriteLine("Eksik veri kontrolü yapılıyor...");
            var inventory = LoadInventory(filePath);

            var missingSerial = inventory.Where(i => string.IsNullOrWhiteSpace(i.SerialNumber)).ToList();
            var missingName = inventory.Where(i => string.IsNullOrWhiteSpace(i.ProductName)).ToList();
            var invalidQuantity = inventory.Where(i => i.Quantity < 0).ToList();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"\nEksik seri no: {missingSerial.Count} kayıt");
            Console.WriteLine($"Eksik ürün adı: {missingName.Count} kayıt");
            Console.WriteLine($"Geçersiz miktar: {invalidQuantity.Count} kayıt");
            Console.ResetColor();
        }

        public List<InventoryItem> LoadInventory(string filePath)
        {
            var inventory = new List<InventoryItem>();

            if (filePath.EndsWith(".xlsx"))
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var item = new InventoryItem
                        {
                            SerialNumber = worksheet.Cells[row, 1].Text?.Trim(),
                            ProductName = worksheet.Cells[row, 2].Text?.Trim(),
                            Quantity = int.TryParse(worksheet.Cells[row, 3].Text, out int qty) ? qty : 0
                        };

                        inventory.Add(item);
                    }
                }
            }
            else if (filePath.EndsWith(".csv"))
            {
                var lines = File.ReadAllLines(filePath);
                foreach (var line in lines.Skip(1))
                {
                    var values = line.Split(',');
                    if (values.Length >= 3)
                    {
                        var item = new InventoryItem
                        {
                            SerialNumber = values[0].Trim(),
                            ProductName = values[1].Trim(),
                            Quantity = int.TryParse(values[2].Trim(), out int qty) ? qty : 0
                        };

                        inventory.Add(item);
                    }
                }
            }
            else if (filePath.EndsWith(".json"))
            {
                string json = File.ReadAllText(filePath);
                inventory = JsonConvert.DeserializeObject<List<InventoryItem>>(json);
            }

            return inventory;
        }

        private List<string> FindDuplicates(string filePath)
        {
            var duplicates = new List<string>();
            var seen = new HashSet<string>();

            if (filePath.EndsWith(".xlsx"))
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var value = worksheet.Cells[row, col].Text?.Trim();
                            if (!string.IsNullOrEmpty(value))
                            {
                                if (seen.Contains(value))
                                {
                                    duplicates.Add(value);
                                }
                                else
                                {
                                    seen.Add(value);
                                }
                            }
                        }
                    }
                }
            }
            else if (filePath.EndsWith(".csv"))
            {
                var lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    var values = line.Split(',');
                    foreach (var value in values)
                    {
                        var trimmed = value.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                        {
                            if (seen.Contains(trimmed))
                            {
                                duplicates.Add(trimmed);
                            }
                            else
                            {
                                seen.Add(trimmed);
                            }
                        }
                    }
                }
            }
            else if (filePath.EndsWith(".json"))
            {
                string json = File.ReadAllText(filePath);
                var data = JsonConvert.DeserializeObject<List<InventoryItem>>(json);

                foreach (var item in data)
                {
                    if (seen.Contains(item.SerialNumber))
                    {
                        duplicates.Add(item.SerialNumber);
                    }
                    else
                    {
                        seen.Add(item.SerialNumber);
                    }
                }
            }

            return duplicates;
        }

        private void GenerateTextStockChart(List<InventoryItem> inventory)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nStok Grafiği (Metin Tabanlı):");
            Console.ResetColor();

            // En fazla 10 ürünü göster
            var topItems = inventory.OrderByDescending(i => i.Quantity).Take(10).ToList();
            int maxQuantity = topItems.Max(i => i.Quantity);

            foreach (var item in topItems)
            {
                int barLength = (int)((double)item.Quantity / maxQuantity * 50);
                string bar = new string('█', barLength);

                Console.ForegroundColor = item.Quantity < _config.CriticalStockThreshold ? ConsoleColor.Red :
                                  item.Quantity < _config.LowStockThreshold ? ConsoleColor.Yellow : ConsoleColor.Green;

                Console.WriteLine($"{item.ProductName.PadRight(20)} | {bar} {item.Quantity}");
                Console.ResetColor();
            }

            // CSV olarak kaydet
            string csvPath = "stok_grafik.csv";
            using (var writer = new StreamWriter(csvPath))
            {
                writer.WriteLine("Ürün Adı,Miktar");
                foreach (var item in inventory)
                {
                    writer.WriteLine($"{item.ProductName},{item.Quantity}");
                }
            }

            Console.WriteLine($"\nStok verileri CSV olarak kaydedildi: {csvPath}");
        }
    }

    // Metin özeti 
    public interface ITextSummaryProcessor
    {
        string GenerateSummary(string filePath);
    }

    // Metin özeti  
    public class TextSummaryProcessor : ITextSummaryProcessor
    {
        public string GenerateSummary(string filePath)
        {
            if (filePath.EndsWith(".txt"))
            {
                return GenerateTextSummary(filePath);
            }
            else if (filePath.EndsWith(".docx"))
            {
                return GenerateDocxSummary(filePath);
            }
            else
            {
                throw new NotSupportedException("Desteklenmeyen dosya formatı");
            }
        }

        private string GenerateTextSummary(string filePath)
        {
            string text = File.ReadAllText(filePath);
            string[] sentences = Regex.Split(text, @"(?<=[.!?])\s+");

            var meaningfulSentences = sentences
                .Where(s => s.Split(' ').Count(w => !IsStopWord(w)) > 3)
                .Take(3)
                .ToList();

            return string.Join(". ", meaningfulSentences) + ".";
        }

        private string GenerateDocxSummary(string filePath)
        {
            var summary = new StringBuilder();

            using (var doc = WordprocessingDocument.Open(filePath, false))
            {
                var body = doc.MainDocumentPart.Document.Body;

                foreach (var para in body.Elements<Paragraph>())
                {
                    string text = para.InnerText;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        summary.AppendLine(text);
                        if (summary.ToString().Split(' ').Length > 100)
                            break;
                    }
                }
            }

            return summary.ToString();
        }

        private bool IsStopWord(string word)
        {
            string[] stopWords = { "ve", "veya", "ancak", "sonuç", "olarak", "için", "bu", "şu", "o", "bir", "ile" };
            return stopWords.Contains(word.ToLower());
        }
    }

    // Dosya dönüştürücü
    public interface IFileConverter
    {
        void ConvertFile(string sourcePath, string outputPath, string conversionType, int sheetNumber);
    }

    // Dosya dönüştürücü 
    public class FileConverter : IFileConverter
    {
        public void ConvertFile(string sourcePath, string outputPath, string conversionType, int sheetNumber)
        {
            Console.WriteLine("Dönüştürme başlatılıyor...");
            ShowProgress();

            var workbook = new Workbook(sourcePath);
            var saveFormat = GetSaveFormat(conversionType);

            if (sheetNumber > 0 && sheetNumber <= workbook.Worksheets.Count)
            {
                // Belirli bir sayfayı dönüştürmek için geçici çalışma kitabı oluşturmak için
                var tempWorkbook = new Workbook();
                var tempWorksheet = tempWorkbook.Worksheets[0];

                // Kaynak çalışma sayfasından verileri kopyala
                var sourceWorksheet = workbook.Worksheets[sheetNumber - 1];
                CopyWorksheet(sourceWorksheet, tempWorksheet);

                tempWorkbook.Save(outputPath, saveFormat);
            }
            else
            {
                workbook.Save(outputPath, saveFormat);
            }
        }

        private SaveFormat GetSaveFormat(string conversionType)
        {
            return conversionType switch
            {
                "1" => SaveFormat.Pdf,
                "2" => SaveFormat.Docx,
                "3" => SaveFormat.Pptx,
                _ => SaveFormat.Pdf
            };
        }

        private void CopyWorksheet(Worksheet source, Worksheet destination)
        {
            // Hücre değerlerini kopyala
            for (int row = 1; row <= source.Cells.MaxDataRow; row++)
            {
                for (int col = 1; col <= source.Cells.MaxDataColumn; col++)
                {
                    var sourceCell = source.Cells[row, col];
                    var destCell = destination.Cells[row, col];

                    destCell.Value = sourceCell.Value;
                    destCell.Value = sourceCell.Value;
                }
            }
        }

        private void ShowProgress()
        {
            for (int i = 0; i <= 100; i += 10)
            {
                Console.Write($"\rİşleniyor: {i}%");
                Thread.Sleep(100);
            }
            Console.WriteLine();
        }
    }

    // Toplu işlemci 
    public interface IBatchProcessor
    {
        void ProcessBatch(string folderPath, string operationType);
    }

    // Toplu işlemci 
    public class BatchProcessor : IBatchProcessor
    {
        private readonly IInventoryProcessor _inventoryProcessor;
        private readonly ITextSummaryProcessor _textSummaryProcessor;
        private readonly IFileConverter _fileConverter;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public BatchProcessor(
            IInventoryProcessor inventoryProcessor,
            ITextSummaryProcessor textSummaryProcessor,
            IFileConverter fileConverter)
        {
            _inventoryProcessor = inventoryProcessor;
            _textSummaryProcessor = textSummaryProcessor;
            _fileConverter = fileConverter;
        }

        public void ProcessBatch(string folderPath, string operationType)
        {
            var files = Directory.GetFiles(folderPath);
            int total = files.Length;
            int processed = 0;

            foreach (var file in files)
            {
                try
                {
                    if (operationType == "1" && file.EndsWith(".xlsx"))
                    {
                        string outputPath = Path.ChangeExtension(file, ".pdf");
                        _fileConverter.ConvertFile(file, outputPath, "1", -1);
                        processed++;
                    }
                    else if (operationType == "2" && (file.EndsWith(".txt") || file.EndsWith(".docx")))
                    {
                        string summary = _textSummaryProcessor.GenerateSummary(file);
                        string summaryPath = Path.ChangeExtension(file, "_summary.txt");
                        File.WriteAllText(summaryPath, summary);
                        processed++;
                    }

                    Console.Write($"\rİşleniyor: {processed}/{total}");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, $"Batch processing failed for {file}");
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\n{processed}/{total} dosya başarıyla işlendi");
            Console.ResetColor();
        }
    }

    // Arama işlemci arayüzü
    public interface ISearchProcessor
    {
        List<SearchResult> SearchInFile(string filePath, string searchTerm);
    }

    // Arama işlemci 
    public class SearchProcessor : ISearchProcessor
    {
        public List<SearchResult> SearchInFile(string filePath, string searchTerm)
        {
            var results = new List<SearchResult>();

            if (filePath.EndsWith(".xlsx"))
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var text = worksheet.Cells[row, col].Text;
                            if (text.Contains(searchTerm, StringComparison.OrdinalIgnoreCase))
                            {
                                results.Add(new SearchResult
                                {
                                    Line = row,
                                    Text = text
                                });
                            }
                        }
                    }
                }
            }
            else if (filePath.EndsWith(".txt") || filePath.EndsWith(".csv"))
            {
                var lines = File.ReadAllLines(filePath);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].Contains(searchTerm, StringComparison.OrdinalIgnoreCase))
                    {
                        results.Add(new SearchResult
                        {
                            Line = i + 1,
                            Text = lines[i]
                        });
                    }
                }
            }
            else if (filePath.EndsWith(".docx"))
            {
                using (var doc = WordprocessingDocument.Open(filePath, false))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    int line = 1;

                    foreach (var para in body.Elements<Paragraph>())
                    {
                        string text = para.InnerText;
                        if (text.Contains(searchTerm, StringComparison.OrdinalIgnoreCase))
                        {
                            results.Add(new SearchResult
                            {
                                Line = line,
                                Text = text
                            });
                        }
                        line++;
                    }
                }
            }

            return results;
        }
    }

    // Rapor oluşturucu 
    public interface IReportGenerator
    {
        void GenerateReport(object data, string title, string filePath = null);
    }

    // Rapor oluşturucu 
    public class ReportGenerator : IReportGenerator
    {
        public void GenerateReport(object data, string title, string filePath = null)
        {
            if (filePath == null)
            {
                filePath = $"{title.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd}.html";
            }

            string html = $@"
<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1 {{ color: #2c3e50; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        .warning {{ color: #f39c12; }}
        .error {{ color: #e74c3c; }}
        .success {{ color: #27ae60; }}
    </style>
</head>
<body>
    <h1>{title}</h1>
    {GenerateReportContent(data)}
    <p><small>Rapor oluşturulma tarihi: {DateTime.Now}</small></p>
</body>
</html>";

            File.WriteAllText(filePath, html);
        }

        private string GenerateReportContent(object data)
        {
            if (data is List<InventoryItem> inventory)
            {
                var sb = new StringBuilder();
                sb.Append("<table>");
                sb.Append("<tr><th>Seri No</th><th>Ürün Adı</th><th>Miktar</th></tr>");

                foreach (var item in inventory)
                {
                    string rowClass = item.Quantity < 5 ? "error" :
                                     item.Quantity < 10 ? "warning" : "";

                    sb.Append($"<tr class='{rowClass}'>");
                    sb.Append($"<td>{item.SerialNumber}</td>");
                    sb.Append($"<td>{item.ProductName}</td>");
                    sb.Append($"<td>{item.Quantity}</td>");
                    sb.Append("</tr>");
                }

                sb.Append("</table>");
                return sb.ToString();
            }
            else if (data is string summary)
            {
                return $"<p>{summary.Replace("\n", "<br>")}</p>";
            }
            else if (data is List<string> duplicates)
            {
                var sb = new StringBuilder();
                sb.Append("<ul>");
                foreach (var item in duplicates)
                {
                    sb.Append($"<li class='error'>{item}</li>");
                }
                sb.Append("</ul>");
                return sb.ToString();
            }
            else
            {
                return $"<pre>{JsonConvert.SerializeObject(data, Formatting.Indented)}</pre>";
            }
        }
    }

    // Yedekleme servisi 
    public interface IBackupService
    {
        string CreateBackup(string filePath);
    }

    // Yedekleme servisi 
    public class BackupService : IBackupService
    {
        public string CreateBackup(string filePath)
        {
            string backupDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Yedekler");
            Directory.CreateDirectory(backupDir);

            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupPath = Path.Combine(backupDir, $"{fileName}_{timestamp}{extension}");

            File.Copy(filePath, backupPath);
            return backupPath;
        }
    }

    // Veri modelleri
    public class InventoryItem
    {
        public string SerialNumber { get; set; }
        public string ProductName { get; set; }
        public int Quantity { get; set; }
    }

    public class SearchResult
    {
        public int Line { get; set; }
        public string Text { get; set; }
    }
}