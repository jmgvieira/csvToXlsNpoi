using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace MyApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Uso correto: programa.exe <caminho_csv> <caminho_excel> <max_linhas> <modo>");
                Console.WriteLine("Modo: 'sheet' para dividir em planilhas, 'file' para criar novos arquivos");
                return;
            }

            string csvPath = args[0];
            string excelBasePath = args[1];
            int maxLinhas = int.Parse(args[2]);
            string modo = args[3].ToLower();

            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Erro: O arquivo CSV '{csvPath}' não foi encontrado.");
                return;
            }

            if (modo != "sheet" && modo != "file")
            {
                Console.WriteLine("Erro: O modo deve ser 'sheet' ou 'file'.");
                return;
            }

            try
            {
                using var reader = new StreamReader(csvPath);
                int arquivoIndex = 1;
                int linhaIndex = 0;
                int totalLinhas = CountLines(csvPath);
                int progressoAtual = 0;

                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Dados");

                while (!reader.EndOfStream)
                {
                    string linha = reader.ReadLine();
                    string[] colunas = linha.Split(';'); // Altere para ',' se necessário
                    IRow row = sheet.CreateRow(linhaIndex);

                    for (int j = 0; j < colunas.Length; j++)
                    {
                        row.CreateCell(j).SetCellValue(colunas[j]);
                    }

                    linhaIndex++;

                    // Atualizar progresso a cada 5%
                    int progresso = (linhaIndex * 100) / totalLinhas;
                    if (progresso >= progressoAtual + 5)
                    {
                        Console.WriteLine($"Progresso: {progresso}% ({linhaIndex}/{totalLinhas} linhas processadas)");
                        progressoAtual = progresso;
                    }

                    if (linhaIndex >= maxLinhas)
                    {
                        string filePath = GetNewFilePath(excelBasePath, arquivoIndex);
                        using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(fs);
                        }

                        Console.WriteLine($"Arquivo Excel salvo: {filePath}");
                        workbook = new XSSFWorkbook();
                        sheet = workbook.CreateSheet("Dados");
                        linhaIndex = 0;
                        arquivoIndex++;
                    }
                }

                // Salvar o último arquivo
                string finalPath = GetNewFilePath(excelBasePath, arquivoIndex);
                using (var fs = new FileStream(finalPath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                Console.WriteLine($"Arquivo Excel final salvo: {finalPath}");
                Console.WriteLine("Conversão concluída com sucesso! ✅");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao converter CSV para Excel: {ex.Message}");
            }
        }

        static int CountLines(string filePath)
        {
            int count = 0;
            using (var reader = new StreamReader(filePath))
            {
                while (reader.ReadLine() != null) count++;
            }
            return count;
        }

        static string GetNewFilePath(string basePath, int index)
        {
            if (index == 1) return basePath;
            string dir = Path.GetDirectoryName(basePath);
            string nomeArquivo = Path.GetFileNameWithoutExtension(basePath);
            string extensao = Path.GetExtension(basePath);
            return Path.Combine(dir, $"{nomeArquivo}_{index}{extensao}");
        }
    }
}