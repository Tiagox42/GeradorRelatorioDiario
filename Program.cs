using System;
using System.Globalization;
using System.IO;
using System.Text;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        // Caminho do arquivo Excel
        string excelFilePath = "Base.xlsx";

        // Abrir o arquivo Excel
        using (var workbook = new XLWorkbook(excelFilePath))
        {
            var worksheet = workbook.Worksheet(1); // Considerando que a planilha de interesse é a primeira

            StringBuilder markdownContent = new StringBuilder();

            // Cabeçalho
            string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
            int itemCount = 0;
            TimeSpan totalTimeSpent = TimeSpan.Zero;

            for (int i = 3; i <= 30; i++)
            {
                var itemName = worksheet.Cell($"A{i}").GetString();
                var timeSpent = worksheet.Cell($"B{i}").GetString();
                var itemDescription = worksheet.Cell($"C{i}").GetString();

                if (!string.IsNullOrWhiteSpace(itemName))
                {
                    itemCount++;

                    if (TimeSpan.TryParseExact(timeSpent, @"hh\:mm", CultureInfo.InvariantCulture, out TimeSpan time))
                    {
                        totalTimeSpent += time;
                    }

                    markdownContent.AppendLine($"## {itemName}");
                    markdownContent.AppendLine($"**Tempo Gasto:** {timeSpent}");
                    markdownContent.AppendLine();
                    markdownContent.AppendLine(itemDescription);
                    markdownContent.AppendLine();
                    markdownContent.AppendLine("---");
                    markdownContent.AppendLine();
                }
            }

            // Adicionando cabeçalho no início do arquivo markdown
            StringBuilder headerContent = new StringBuilder();
            headerContent.AppendLine($"# Relatório de Atividades - {currentDate}");
            headerContent.AppendLine($"**Quantidade de Itens Trabalhados:** {itemCount}");
            headerContent.AppendLine($"**Tempo Total Gasto:** {totalTimeSpent:hh\\:mm}");
            headerContent.AppendLine();
            headerContent.AppendLine("---");
            headerContent.AppendLine();

            headerContent.Append(markdownContent.ToString());

            // Caminho do arquivo Markdown a ser gerado
            string markdownFilePath = "Relatorio.md";
            File.WriteAllText(markdownFilePath, headerContent.ToString());
        }
    }
}
