using System;
using System.IO;
using System.Runtime.InteropServices;
using ClosedXML.Excel;

namespace CsvToExcelConverter
{
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetConsoleOutputCP(uint wCodePageID);
        static void Main(string[] args)
        {

            SetConsoleOutputCP(65001);
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            // Verifica os parâmetros: 
            // Uso: programa.exe <input.csv> <output.xlsx> <max_linhas> <modo(sheet/excel)>
            if (args.Length < 4)
            {
                Console.WriteLine("Uso: programa.exe <input.csv> <output.xlsx> <max_linhas> <modo(sheet/excel)>");
                return;
            }

            string inputCsv = args[0];
            string outputExcel = args[1];
            if (!int.TryParse(args[2], out int maxRows) || maxRows <= 0)
            {
                Console.WriteLine("O número máximo de linhas deve ser um inteiro positivo.");
                return;
            }
            string mode = args[3].ToLower();
            if (mode != "sheet" && mode != "excel")
            {
                Console.WriteLine("O modo deve ser 'sheet' ou 'excel'.");
                return;
            }

            try
            {
                ProcessCsvToExcel(inputCsv, outputExcel, maxRows, mode);
                Console.WriteLine("\n✅ Processamento concluído com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ Erro: {ex.Message}");
            }
        }

        static void ProcessCsvToExcel(string inputCsv, string outputExcel, int maxRows, string mode)
        {
            // Obtém o tamanho total do ficheiro para cálculo da percentagem
            FileInfo fileInfo = new FileInfo(inputCsv);
            long totalSize = fileInfo.Length;

            // Abre o ficheiro CSV para leitura (streaming)
            using (StreamReader sr = new StreamReader(inputCsv))
            {
                if (mode == "sheet")
                {
                    // Modo: várias sheets num único ficheiro Excel
                    using (var workbook = new XLWorkbook())
                    {
                        int sheetIndex = 1, rowIndex = 1;
                        var worksheet = workbook.AddWorksheet("Sheet" + sheetIndex);
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            // Se exceder o número máximo de linhas, cria uma nova sheet
                            if (rowIndex > maxRows)
                            {
                                sheetIndex++;
                                rowIndex = 1;
                                worksheet = workbook.AddWorksheet("Sheet" + sheetIndex);
                            }

                            // Divide a linha com base no delimitador (alterar se necessário)
                            string[] columns = line.Split(';');
                            for (int col = 0; col < columns.Length; col++)
                                worksheet.Cell(rowIndex, col + 1).Value = columns[col];

                            rowIndex++;

                            // Atualiza a percentagem com base na posição lida do ficheiro
                            double progress = sr.BaseStream.Position / (double)totalSize * 100;
                            Console.Write($"\r🔄 Processando: {progress:F2}% concluído...");
                        }
                        workbook.SaveAs(outputExcel);
                    }
                }
                else if (mode == "excel")
                {
                    // Modo: vários ficheiros Excel
                    int fileIndex = 1, rowIndex = 1;
                    var workbook = new XLWorkbook();
                    var worksheet = workbook.AddWorksheet("Sheet1");
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        // Se exceder o número máximo de linhas, salva o ficheiro atual e inicia um novo
                        if (rowIndex > maxRows)
                        {
                            string fileName = outputExcel.Replace(".xlsx", $"_{fileIndex}.xlsx");
                            workbook.SaveAs(fileName);
                            workbook.Dispose();

                            fileIndex++;
                            rowIndex = 1;
                            workbook = new XLWorkbook();
                            worksheet = workbook.AddWorksheet("Sheet1");
                        }

                        string[] columns = line.Split(';');
                        for (int col = 0; col < columns.Length; col++)
                            worksheet.Cell(rowIndex, col + 1).Value = columns[col];

                        rowIndex++;

                        double progress = sr.BaseStream.Position / (double)totalSize * 100;
                        Console.Write($"\r🔄 Processando: {progress:F2}% concluído...");
                    }

                    // Salva o último ficheiro se houver linhas não salvas
                    string finalFileName = outputExcel.Replace(".xlsx", $"_{fileIndex}.xlsx");
                    workbook.SaveAs(finalFileName);
                    workbook.Dispose();
                }
            }
        }
    }
}
