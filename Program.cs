using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OfficeOpenXml;
using System.Diagnostics;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

class Program
{
static void Main(string[] args){

        DateTime ini = DateTime.Now;
        Console.WriteLine($"Processo iniciado em: {ini:HH:mm:ss}");
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Definir contexto de licença

        var edgeOptions = new EdgeOptions();

        edgeOptions.AddArguments("headless");

        IWebDriver driver = new EdgeDriver(edgeOptions);
    
        inicio(driver);
 
        driver.Quit();

        DateTime fim = DateTime.Now;
        Console.WriteLine($"Processo Finalizado em: {fim:HH:mm:ss}");
    }
static void inicio(IWebDriver driver){
        //modificar o caminho 
        string caminhoArquivo = @"C:\AUTOMACAO\JURIDICO\LEVA JANEIRO 25\base_juridico.xlsm";
    
        using (var package = new ExcelPackage(new FileInfo(caminhoArquivo)))
        {
            var sheet = package.Workbook.Worksheets["DADOS"];
            string contrato = "";

            //verificar a ultima linha digitada e comecar o loop da linha 6
            //string ultimaLinha = sheet.Dimension.End.Row;
            int ultimaLinha = ObterUltimaLinhaComDados(sheet);

            Console.WriteLine($"LInhas na plan com dados {ultimaLinha},Vamos iniciar os relatorios e ao todo teremos {ultimaLinha - 6} linhas para processar");

            for (int i = 2; i < ultimaLinha; i++ ){
                contrato = sheet.Cells[i,1].Text;
                string status = sheet.Cells[i,5].Text;

                if (contrato != "" && status != "Finalizado" ){
                    DateTime ini = DateTime.Now;
                    Console.WriteLine($"---CONTRATO - {contrato} - iniciado em: {ini:HH:mm:ss}---");
                    
                    nome = sheet.Cells[i,4].Text;
                    cpf = sheet.Cells[i,3].Text;

                    Automacao(contrato,driver);
                    
                    sheet.Cells[i,6].Value = caminhoArquivo_;
                    sheet.Cells[i,5].Value = "Finalizado";
                    sheet.Cells[i,8].Value = bruto;
                    sheet.Cells[i,9].Value = devedor;


                    //define horario de fim para calcular o tempo e salvar na planilha para conferir
                    DateTime fim = DateTime.Now;
                    Console.WriteLine($"-LINHA-{i}-CONTRATO-{contrato}-FINALIZADO-{fim:HH:mm:ss}-");
                    TimeSpan diferenca = fim - ini;
                    string formatoMinSeg = $"{(int)diferenca.TotalMinutes:00}:{diferenca.Seconds:00}";
                    sheet.Cells[i,7].Value = formatoMinSeg;
                    package.Save();
                }
            }
            
        }
    }
static int ObterUltimaLinhaComDados(ExcelWorksheet? sheet)
{
    // Se a planilha ou a dimensão da planilha for nula, retorna 0
    if (sheet?.Dimension == null)
        return 0;

    // Retorna a última linha com dados
    return sheet.Dimension.End.Row;
}
static string nome = "";
static string cpf = "";
static string bruto = "";
static string devedor = "";
static string caminhoArquivo_ = "";
static void Automacao(string contrato,IWebDriver driver)
    {
        // Caminho para o ChromeDriver (ajuste para o seu ambiente)
        //IWebDriver driver = new EdgeDriver(@"C:\chrome-win64\chrome.exe");
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

        //Console.WriteLine($"Contrato - {contrato}");
        //string contrato = Console.ReadLine();
        string urlContrato = $"http://18.217.139.90/WebAppBackOffice/Pages/Contrato/ICContratoDetalhes?nrOper={contrato}";
        //string urlBackoffice = "http://18.217.139.90/WebAppBackOffice/Pages/Contrato";

        try
        {
            // 1. Acessar o site
            driver.Navigate().GoToUrl(urlContrato);
            //Console.WriteLine("Site acessado com sucesso!");

            // 2. Verificar se o campo de login existe
            if (ElementoExiste(driver, By.Id("txtUsuario_CAMPO"))) // Verifica se o campo de login existe
            {
                //Console.WriteLine("Campos de login encontrados! Realizando login...");

                // 3. Preencher os campos de login
                IWebElement campoUsuario = driver.FindElement(By.Id("txtUsuario_CAMPO"));
                IWebElement campoSenha = driver.FindElement(By.Id("txtSenha_CAMPO"));
                IWebElement botaoLogin = driver.FindElement(By.Id("bbConfirmar"));

                campoUsuario.SendKeys("OTCWDJQA"); // Insira seu usuário
                campoSenha.SendKeys("YIXS@KfJ");  // Insira sua senha
                botaoLogin.Click(); // Clica no botão de login

               //Console.WriteLine("Login realizado com sucesso!");

                // 4. Esperar carregamento pós-login
                Thread.Sleep(1000); // Aguarde 3 segundos
                driver.Navigate().GoToUrl(urlContrato);
            }
            else
            {
                //Console.WriteLine("Campos de login não encontrados. Continuando com o código...");

                // Coloque aqui as ações alternativas
            }

                    // Navega até o contrato e aguarda o elemento da página carregar
            wait.Until(ExpectedConditions.ElementExists(By.Id("ctl00_Cph_lblNumeroContrato")));

            // Valida se o elemento foi encontrado na página
            if (ElementoExiste(driver, By.Id("ctl00_Cph_lblNumeroContrato")))
            {
                // Captura o convênio e a taxa para uso posterior
                IWebElement taxaWeb = driver.FindElement(By.Id("ctl00_Cph_TabContainer1_panelContrato_lblTaxaCLmes"));
                string taxa = taxaWeb.Text;

                // Troca para a aba de parcelas
                IWebElement abaParcelas = driver.FindElement(By.Id("__tab_ctl00_Cph_TabContainer1_panelParcelas"));
                abaParcelas.Click();

                // Aguarda um breve tempo para garantir o carregamento da aba (substitua Thread.Sleep por uma espera explícita)
                Thread.Sleep(500);

                // Instancia a tabela para pegar os dados das parcelas
                IWebElement tabelaParcelas = driver.FindElement(By.Id("ctl00_Cph_TabContainer1_panelParcelas_gridParcelas"));
                IList<IWebElement> linhas = tabelaParcelas.FindElements(By.XPath(".//tbody/tr"));

                List<List<string>> dadosTabela = new List<List<string>>();

                // Itera pelas linhas da tabela e extrai os dados
                foreach (var linha in linhas)
                {
                    IList<IWebElement> celulas = linha.FindElements(By.TagName("td"));
                    List<string> linhaDados = new List<string>();

                    // Adiciona o conteúdo de cada célula à linha de dados
                    foreach (var celula in celulas)
                    {
                        linhaDados.Add(celula.Text.Trim()); // Remove espaços desnecessários
                    }

                    // Adiciona apenas linhas que não estejam vazias
                    if (linhaDados.Count > 0)
                    {
                        dadosTabela.Add(linhaDados);
                    }
                }

                // Exporta os dados capturados para o Excel
                string caminhoTemplate = @"C:\AUTOMACAO\JURIDICO\LEVA JANEIRO 25\TEMPLATE\template_ccb.xlsm";
                ExportarParaExcel(dadosTabela, caminhoTemplate, contrato, taxa);

                //Console.WriteLine("Dados exportados com sucesso!");
            }
            else
            {
                //Console.WriteLine("Elemento do contrato não foi encontrado. Verifique a página.");
            }


            //FIM DO CODIGO
            //Console.WriteLine("Execução finalizada.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Erro: " + ex.Message);
        }
        finally
        {
            Console.WriteLine("");
        }
                  
    }

    /// <summary>
    /// Função que verifica se um elemento existe.
    /// </summary>
    /// <param name="driver"></param>
    /// <param name="by"></param>
    /// <returns></returns>

static bool ElementoExiste(IWebDriver driver, By by)
    {
        try
        {
            return driver.FindElement(by) != null;
        }
        catch
        {
            return false;
        }
    }
static void ExportarParaExcel(List<List<string>> dados, string caminhoArquivo, string contrato, string taxa)
{
    try
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Definir contexto de licença

        // Abrir o arquivo de template existente
        using (var package = new ExcelPackage(new FileInfo(caminhoArquivo)))
        {
            var sheet = package.Workbook.Worksheets[0]; // Seleciona a primeira planilha do arquivo de template
            string aux = "";
            double valor = 0;
            int auxi = 0;
            int lin1 = 5; // Linha inicial para os dados
            int lin2 = 2; // Coluna inicial para os dados

            // Escrever os dados na planilha
            for (int i = 0; i < dados.Count; i++)
            {
                for (int j = 0; j < dados[i].Count; j++)
                {
                    if (j < 6)
                    {
                        switch (j)
                        {
                            case 0: // Conversão para inteiro
                                auxi = Convert.ToInt32(dados[i][j]);
                                sheet.Cells[i + lin1, j + lin2].Value = auxi;
                                break;

                            case 2: // Valor numérico (monetário)
                            case 5:
                                aux = dados[i][j].Replace("R$ ", "");
                            
                                valor = Convert.ToDouble(aux);
                                sheet.Cells[i + lin1, j + lin2].Value = valor;
                                sheet.Cells[i + lin1, j + lin2].Style.Numberformat.Format = "#,##0.00";
                                break;
                                
                            default: // Texto genérico
                                sheet.Cells[i + lin1, j + lin2].Value = dados[i][j];
                                break;
                        }
                    }
                }
            }

            // Determinar a última linha preenchida na planilha
            int lastRow = sheet.Dimension.End.Row;

            // Inserir o valor da taxa em K2
            double taxaaux = Convert.ToDouble(taxa.Replace("%", ""));
            sheet.Cells[2, 14].Value = taxaaux; // Atualizar célula K2
            sheet.Cells[3,9].Value = $"Contrato: {contrato}";
            sheet.Cells[3,3].Value = $"Nome: {nome}";
            sheet.Cells[3,6].Value = $"CPF: {cpf}";
            

            // Garantir que o diretório de destino exista
            string nome_ = nome;

            string diretorio = Path.Combine(@$"C:\AUTOMACAO\JURIDICO\LEVA JANEIRO 25\RELATORIOS\{nome_}\");
            if (!Directory.Exists(diretorio))
            {
                Directory.CreateDirectory(diretorio);
            }

            // Salvar o arquivo Excel modificado
            string novoCaminho = Path.Combine(diretorio, $"{contrato}.xlsm");
            package.SaveAs(new FileInfo(novoCaminho));

            caminhoArquivo_ = novoCaminho;

            string caminhoVBS = @"C:\AUTOMACAO\JURIDICO\LEVA JANEIRO 25\TEMPLATE\teste.vbs";
            string nomeMacro = "teste";
            //caminhoArquivo = @$"C:\AUTOMACAO\RELATORIOS PARCELAS\RELATORIOS\{contrato}.xlsm";

            montarVBS(novoCaminho,nomeMacro,caminhoVBS);

            devedor = sheet.Cells[2,16].Text;
            bruto = sheet.Cells[2,15].Text;

            }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erro ao exportar para Excel e PDF: {ex.Message}");
    }
}
static void montarVBS(string caminhoArquivo, string nomeMacro,string caminhoVBS)
    {

        //string caminhoVBS = @"C:\AUTOMACAO\RELATORIOS PARCELAS\TEMPLATE\teste.vbs";
        string conteudoVBS = $@"
Dim xlApp 
Dim xlWorkbook
Dim caminhoArquivo

' Caminho para o arquivo Excel
caminhoArquivo = ""{caminhoArquivo}""

On Error Resume Next ' Inicia o tratamento de erros

' Cria uma instância do Excel
Set xlApp = CreateObject(""Excel.Application"")
xlApp.Visible = False  ' O Excel não será visível

' Abre o arquivo Excel
Set xlWorkbook = xlApp.Workbooks.Open(caminhoArquivo)

' Verifica se o arquivo foi aberto com sucesso
If Err.Number <> 0 Then
    WScript.Echo ""Erro ao abrir o arquivo: "" & Err.Description
    WScript.Quit
End If

' Executa a macro chamada ""teste""
xlApp.Run ""{nomeMacro}""

' Salva e fecha o arquivo
xlWorkbook.Save
xlWorkbook.Close

' Libera os objetos
Set xlWorkbook = Nothing
xlApp.Quit
Set xlApp = Nothing
";

        // Cria o arquivo VBS
        File.WriteAllText(caminhoVBS, conteudoVBS);
        
        RunVbsScriptViaCmd(caminhoVBS);
    }
static void RunVbsScriptViaCmd(string caminhoVBS)
    {
        try
        {
            // Verifica se o arquivo VBS existe
            if (!System.IO.File.Exists(caminhoVBS))
            {
               // Console.WriteLine($"O arquivo VBS não existe: {caminhoVBS}");
                return;
            }

            // Configura o processo para executar o script VBS via CMD
            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                FileName = "cmd.exe",  // Executa o CMD
                Arguments = $"/c wscript \"{caminhoVBS}\"", // Comando para executar o script VBS
                UseShellExecute = false,   // Deve ser false para usar variáveis de ambiente
                CreateNoWindow = true      // Se você não quiser que uma janela apareça
            };

            // Inicia o processo
            using (Process process = Process.Start(startInfo))
            {
                process.WaitForExit();  // Aguarda o script VBS terminar de rodar
            }

           //Console.WriteLine("Script VBS executado com sucesso.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao executar o script VBS: {ex.Message}");
        }
    }
}