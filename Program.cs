using System;
using System.Globalization;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using LibGit2Sharp;

class Program
{
    static void Main()
    {
        try
        {
            // Caminho do diretório base
            string baseDirectory = @"C:\Users\Dantas\Documents\GitHub\GeradorRelatorioDiario";

            // Caminho do arquivo Excel
            string excelFilePath = Path.Combine(baseDirectory, "Base.xlsx");

            // Caminho do diretório de relatórios
            string reportsDirectory = Path.Combine(baseDirectory, "Relatorios");

            // Caminho do diretório de logs
            string logsDirectory = Path.Combine(baseDirectory, "Logs");

            // Verificar se o diretório "Relatorios" existe, se não, criar
            if (!Directory.Exists(reportsDirectory))
            {
                Directory.CreateDirectory(reportsDirectory);
            }

            // Verificar se o diretório "Logs" existe, se não, criar
            if (!Directory.Exists(logsDirectory))
            {
                Directory.CreateDirectory(logsDirectory);
            }

            var caminhoMd = "";

            // Abrir o arquivo Excel
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1); // Considerando que a planilha de interesse é a primeira

                StringBuilder markdownContent = new StringBuilder();

                // Cabeçalho
                string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                string reportDate = DateTime.Now.ToString("yyyyMMdd_HHmmss"); // Data formatada para o nome do arquivo
                int itemCount = 0;
                TimeSpan totalTimeSpent = TimeSpan.Zero;

                for (int i = 3; i <= 30; i++)
                {
                    var itemName = worksheet.Cell($"A{i}").GetString();
                    var timeSpent = worksheet.Cell($"B{i}").GetString();
                    var itemDescription = worksheet.Cell($"C{i}").GetString();

                    if (!string.IsNullOrWhiteSpace(itemName))
                    {
                        Console.WriteLine($"Item: {itemName}, Tempo: {timeSpent}"); // Log para verificar os valores lidos

                        itemCount++;

                        // Converter a string de tempo para o formato "HH:mm:ss"
                        if (DateTime.TryParseExact(timeSpent, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime timeAsDateTime))
                        {
                            TimeSpan time = timeAsDateTime.TimeOfDay;
                            Console.WriteLine($"Tempo convertido: {time}"); // Log para verificar o tempo convertido
                            totalTimeSpent += time;
                        }
                        else
                        {
                            Console.WriteLine($"Falha ao converter tempo: {timeSpent}"); // Log para indicar falha na conversão do tempo
                        }

                        markdownContent.AppendLine($"## {itemName}");
                        markdownContent.AppendLine($"**Data e tempo gasto:** {timeSpent}");
                        markdownContent.AppendLine();
                        markdownContent.AppendLine(itemDescription);
                        markdownContent.AppendLine();
                        markdownContent.AppendLine("---");
                        markdownContent.AppendLine();
                    }
                }

                // Formatar o tempo total gasto
                string totalTimeFormatted = $"{(int)totalTimeSpent.TotalHours:D2}:{totalTimeSpent.Minutes:D2}";

                // Adicionando cabeçalho no início do arquivo markdown
                StringBuilder headerContent = new StringBuilder();
                headerContent.AppendLine($"# Relatório de Atividades - {currentDate}");
                headerContent.AppendLine($"**Quantidade de Itens Trabalhados:** {itemCount}");
                headerContent.AppendLine();
                headerContent.AppendLine($"**Tempo Total Gasto:** {totalTimeFormatted}");
                headerContent.AppendLine();
                headerContent.AppendLine("---");
                headerContent.AppendLine();


                headerContent.Append(markdownContent.ToString());

                // Caminho do arquivo Markdown a ser gerado com a data no nome
                string markdownFilePath = Path.Combine(reportsDirectory, $"Relatorio_{reportDate}.md");

                // Escrever o conteúdo no arquivo Markdown
                File.WriteAllText(markdownFilePath, headerContent.ToString());

                caminhoMd = markdownFilePath;
            }

            // Perguntar se o usuário quer fazer commit e push para o GitHub
            Console.Write("Deseja fazer commit e enviar para o GitHub? (s/n): ");
            string response = Console.ReadLine().ToLower();

            if (response == "s")
            {
                CommitAndPushToGitHub(baseDirectory);
            }

            Console.WriteLine("Abrindo o arquivo Markdown no GitHub...");
            OpenMarkdownFileOnGitHub(caminhoMd);
            Console.WriteLine("Arquivo Markdown aberto no GitHub.");
        }
        catch (Exception ex)
        {
            // Caminho do diretório de logs
            string baseDirectory = @"C:\Users\Dantas\Documents\GitHub\GeradorRelatorioDiario";
            string logsDirectory = Path.Combine(baseDirectory, "Logs");

            // Nome do arquivo de log
            string logFileName = $"Log_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.txt";
            string logFilePath = Path.Combine(logsDirectory, logFileName);

            // Escrever o erro no arquivo de log
            File.WriteAllText(logFilePath, ex.ToString());

            Console.WriteLine("Ocorreu um erro. Verifique o arquivo de log para mais detalhes.");
        }
    }

    static void CommitAndPushToGitHub(string repoPath)
    {
        try
        {
            // Caminho do arquivo que contém o token
            string tokenFilePath = Path.Combine(repoPath, "token.txt");

            // Ler o token do arquivo
            string token = File.ReadAllText(tokenFilePath).Trim();

            // Mensagem de commit
            string commitMessage = "Adicionando novos relatórios.";

            // Configurações do autor do commit
            var signature = new Signature("Tiago Dantas", "tiagodantas42@gmail.com", DateTimeOffset.Now);

            using (var repo = new Repository(repoPath))
            {
                // Adicionar os arquivos ao índice
                Commands.Stage(repo, "*");

                // Realizar o commit
                repo.Commit(commitMessage, signature, signature);

                // Push para o repositório remoto
                var remote = repo.Network.Remotes["origin"];
                var options = new PushOptions
                {
                    CredentialsProvider = (_, __, ___) =>
                        new UsernamePasswordCredentials
                        {
                            Username = "Tiagox42", // Seu nome de usuário no GitHub
                            Password = token  // Usando o token lido do arquivo
                        }
                };

                repo.Network.Push(remote, @"refs/heads/main", options);
            }
        }
        catch (Exception ex)
        {
            // Tratamento de erro
            Console.WriteLine($"Erro ao fazer commit e push para o GitHub: {ex.Message}");
        }
    }

    static void OpenMarkdownFileOnGitHub(string filePath)
    {
        try
        {
            // URL base do repositório no GitHub
            string baseGitHubUrl = "https://github.com/Tiagox42/GeradorRelatorioDiario/blob/main/";

            // Constrói a URL completa para o arquivo Markdown
            string gitHubUrl = baseGitHubUrl + "Relatorios/" + Path.GetFileName(filePath);

            // Abre a URL no navegador padrão
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = gitHubUrl,
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao abrir o arquivo no GitHub: {ex.Message}");
        }
    }
}
