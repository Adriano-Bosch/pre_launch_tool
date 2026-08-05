using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Data.OleDb;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Windows.Threading;
using System.Windows.Media;
using System.Diagnostics;
using System.ComponentModel;
using System.Globalization;
using Xceed.Words.NET;
using Xceed.Document.NET;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using Microsoft.Win32;
using System.DirectoryServices.AccountManagement;
using CheckBox = System.Windows.Controls.CheckBox;
using ComboBox = System.Windows.Controls.ComboBox;
using Label = System.Windows.Controls.Label;

namespace Pre_Launch_Tool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        // defined strings for paths and etc
        
        // Loaded AI prompt template
        private string? viabilityPrompt;

        // Prefer UNC path because mapped drive letters may not exist for all users.
        private static readonly string[] DbPathCandidates =
        {
            @"\\bosch.com\dfsrb\DfsBR\loc\Ca1\AA\tr_rbr\Inter_Setor\Master_Data\09. Pre-Launch Tool\[não alterar] banco de dados\PE_DB_v1.accdb",
            @"S:\AA\tr_rbr\Inter_Setor\Master_Data\09. Pre-Launch Tool\[não alterar] banco de dados\PE_DB_v1.accdb"
        };
        private static readonly string[] DbSearchRoots =
        {
            @"\\bosch.com\dfsrb\DfsBR\loc\Ca1\AA\tr_rbr\Inter_Setor\Master_Data\09. Pre-Launch Tool",
            @"S:\AA\tr_rbr\Inter_Setor\Master_Data\09. Pre-Launch Tool"
        };
        private static readonly string[] OleDbProviderCandidates =
        {
            "Microsoft.ACE.OLEDB.16.0",
            "Microsoft.ACE.OLEDB.12.0"
        };
        private const string AppLogFolderName = "Pre_Launch_Tool\\Logs";

        readonly string sourceDbPath;
        readonly string dbPath;
        readonly string activeOleDbProvider;
        readonly string startupDiagnosticLogPath;
        private string templatePath;
        private string emailAddress;
        readonly string connectionString;

        // COM Office objects removed — using EPPlus/Xceed instead

        // Cache para resultados de queries
        private Dictionary<string, object> queryCache = new Dictionary<string, object>();

        // Controle de operações em andamento
        private CancellationTokenSource currentFilterOperation;

        // Dados estáticos para uso offline (reduz acesso ao banco)
        private List<FleetDataRow> baseDataTable;
        // Último resultado filtrado usado para geração de relatórios (inclui motorização)
        private List<FleetDataRow> lastFilteredFleetList = new List<FleetDataRow>();

        // Items source for datagrid
        ObservableCollection<PreviewItem> previewItems;

        // competitor (MARKENBEZ -> TW -> TWNR_VERD) loading
        private sealed record CompetitorRow(string Markenbez, string Tw, string? TwnrVerd);
        private List<CompetitorRow> competitorTable = new();
        private string? selectedCompetitorVerdCode;

        private const string AiBasesPath = @"\\bosch.com\dfsrb\DfsBR\loc\Ca1\AA\tr_rbr\Inter_Setor\MEETINGS_AND_OPLS\2) Meetings and WS\Data Driven\4. Foco de trabalho 1 - Pré-Launch\1. Arquivos\Bases IA - PDF";
        private const string M365AgentUrl = "https://m365.cloud.microsoft/chat/?titleId=T_bc53de94-da30-5705-59d8-b21b92a6773f&source=embedded-builder";
        private const string AllowedRegion = "Brazil";
        private const string SelectAllOption = "Select All";
        private bool isBulkChecking;
                private const string EmbeddedViabilityPrompt = """
# CONTEXTO E OBJETIVO

Você é um analista de produtos especialista no mercado de reposição automotiva.
Faça uma análise de viabilidade de lançamento usando exclusivamente os dados fornecidos no TSV e as bases internas.

Restrição crítica:
- Não usar internet.
- Não usar conhecimento externo.
- Não inferir dados não suportados pelas bases.

# DADOS A CONSULTAR NAS BASES

Com base nos dados de entrada, monte tabelas curtas (uma linha por registro, uma coluna por atributo, sem parágrafos longos em célula):

1) 01.BASE_FROTA
- 1.1.Volume_Total_da_Frota_Circulante_(FAS_POPULATION)
- 1.2.Código_do_veículo_(SHORT)

2) 02.BASE_FRAGA
- 2.1.Competidor
- 2.2.Código_do_Competidor
- 2.3.Preço_do_Competidor

3) 03.BASE_EPOS
- 3.3.Market_Share_Atual (MS)
- 3.4.Taxa_de_Troca_do_Produto (%)
- 3.5.Quantidade_Trocada_por_Veículo (%)

4) 04.BASE_PM_PREMISSES
- 4.1.Business_Unit
- 4.2.Nome_do_produto
- 4.3.Premissa_de_volume
- 4.4.Premissa_de_Faturamento

5) 05.BASE_OE
- 5.1.SHORT (Bosch Key)
- 5.2.OE_number

# ESTRUTURA OBRIGATÓRIA DA ANÁLISE

Passo 1.0 - Tabelas de entrada
- Tabela A (2 colunas):
    - 1.2.Business_Unit | valor informado
    - 1.3.Nome_do_produto | valor informado
    - 1.1.País_de_circulação | valor informado
    - 1.4.Market_Share_estipulado_pelo_usuário | valor informado
- Tabela B (9 colunas): Bosch Key | Código OE | Marca | Veículo | Ano de Fabricação | Motorização | Carroceria | Tipo de Combustível | Códigos OE.
- Um Bosch Key por linha.
- Se faltar dado de entrada, buscar nas bases.
- Se houver múltiplos valores, usar linhas adicionais (não concatenar texto longo).

Passo 1.1 - Potencial de Mercado Volume
- Calcular PMV por veículo e total:
    PM = 1.5.Volume_Total_da_Frota_Circulante * 3.3.Market_Share_Atual * 3.4.Taxa_de_Troca_do_Produto * 3.5.Quantidade_Trocada_por_Veículo
- Calcular PMU por veículo e total:
    PMU = 1.5.Volume_Total_da_Frota_Circulante * 1.1.Market_Share_estipulado_pelo_usuário * 3.4.Taxa_de_Troca_do_Produto * 3.5.Quantidade_Trocada_por_Veículo
- Expor resultados na coluna VOLUME_PAÍS.

Passo 2.1 - Preços de mercado
- Considerar 3.1.Business_Unit, 3.2.Nome_do_produto, 1.11.Nome_Competidor e 1.12.Código_Competidor.
- Gerar tabela: Código do Concorrente | Preço do Concorrente.

Passo 2.2 - Potencial de Mercado Faturamento (PMF)
- Cálculo 1: média de preço concorrente * PMV
- Cálculo 2: média de preço concorrente * PMU
- Cálculo 3: menor preço concorrente * PMV
- Cálculo 4: menor preço concorrente * PMU
- Cálculo 5: maior preço concorrente * PMV
- Cálculo 6: maior preço concorrente * PMU

Passo 3 - Análise mercadológica
- Estimar volume anual capturável com PM e PMU.
- Considerar Market Share atual e Market Share estipulado pelo usuário.

Passo 4 - Comparativo com premissas de Portfolio Management
- Comparar:
    - PMV vs 4.3.Premissa_de_volume
    - PMF vs 4.4.Premissa_de_Faturamento
- Tabela com 7 colunas:
    4.1.Business_Unit | 4.2.Nome_do_produto | 4.3.Premissa_de_volume | Cálculo de Potencial de Mercado Volume | 4.4.Premissa_de_faturamento | Cálculo de Potencial de Mercado Faturamento | Ok/NOk
- Regra de status:
    - Ok: PMV > premissa de volume e/ou PMF > premissa de faturamento
    - NOk: PMV < premissa de volume e/ou PMF < premissa de faturamento

Passo 5 - Síntese e viabilidade
- Consolidar cálculos.
- Explicar pontos fortes e riscos.

Passo 6 - Recomendação final
- Escolher exatamente um status:
    - [RECOMENDADO]
    - [RECOMENDADO COM RESSALVAS]
    - [NÃO RECOMENDADO]
- Justificar em 2-3 frases conectando dados e conclusão.

# ESTRUTURA DO RELATÓRIO PDF FINAL

Formato e visual:
- Fonte principal: Arial.
- Cabeçalho: logo Bosch (esquerda) e título Análise de Viabilidade de Produto (direita).
- Rodapé: Página X de Y + CONFIDENCIAL | BOSCH Mobility Aftermarket.
- Cores:
    - H1: #005691
    - H2/H3 e texto: #333333
    - Cabeçalho de tabela: #F0F0F0
    - Status Ok: #28a745
    - Status NOk: #dc3545

Estrutura:
- Página 1: capa (produto, aplicação, país, data).
- Página 2: glossário das siglas e abreviações usadas.
- Página 3+:
    1. Sumário executivo (status + justificativa)
    2. Cenário e dados de entrada (Passo 1.0)
    3. Potencial de mercado volume (Passo 1.1)
    4. Comparativo com premissas (Passo 4)
    5. Análise mercadológica e concorrência (Passos 2 e 3)
    6. Síntese, riscos e recomendação final (Passos 5 e 6)
- Se não houver dados de preço de concorrente, incluir nota explícita.

### REGRAS OBRIGATÓRIAS DE CONTROLE DE TABELAS (NÃO NEGOCIÁVEL)

Antes de finalizar o PDF execute obrigatoriamente esta validação:

ETAPA 1 - Cálculo de largura
- Somar a largura de todas as colunas.
- Garantir que a largura total seja <= largura útil da página.
- Nunca usar larguras automáticas.

ETAPA 2 - Controle de quebra
- Aplicar quebra automática de texto em todas as células.
- Nenhum texto pode ultrapassar a borda da célula.
- Nenhum texto pode sobrepor outra célula.

ETAPA 3 - Adaptação automática
Se a tabela não couber:
1. Reduzir colunas descritivas.
2. Reduzir fonte gradualmente até 8 pt.
3. Aumentar altura das linhas.
4. Alterar a orientação da página para paisagem.
5. Se ainda não couber, dividir a tabela em múltiplas partes.

ETAPA 4 - Inspeção visual obrigatória
Após gerar o PDF:
- Renderizar TODAS as páginas em imagem.
- Verificar visualmente cada página.
- Identificar:
    * texto cortado
    * texto fora da página
    * sobreposição de células
    * tabelas ultrapassando margens
    * cabeçalhos ilegíveis

ETAPA 5 - Correção obrigatória
Caso qualquer problema seja encontrado:
- Regenerar automaticamente a tabela.
- Repetir a inspeção visual.
- Não entregar o PDF até que todas as tabelas estejam corretas.

ETAPA 6 - Critérios de aprovação
Somente entregar o PDF quando:
- Nenhum texto estiver fora da página
- Nenhuma célula estiver sobreposta
- Todas as colunas estiverem visíveis
- Todos os valores estiverem legíveis
- Cabeçalhos totalmente visíveis
- Tabelas respeitarem as margens

Se qualquer critério falhar, o PDF deve ser recriado automaticamente.

REGRA ESPECIAL PARA TABELAS COM MAIS DE 8 COLUNAS
- Gerar automaticamente em orientação paisagem.
- Utilizar fonte 8 pt.
- Utilizar colunas proporcionais.
- Priorizar largura para: Bosch Key, OE Number, Volume e Faturamento.
- Colunas descritivas devem quebrar linha automaticamente.
- É proibido compactar texto até ficar ilegível.
""";
        private bool suppressRegionRestrictionMessage;
        private bool isFilteringVehicleInputs;

                private string LoadViabilityPromptTemplate()
                {
                        try
                        {
                                var candidates = new[]
                                {
                                        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PromptTemplates", "viability_prompt.txt"),
                                        Path.Combine(AppContext.BaseDirectory, "PromptTemplates", "viability_prompt.txt"),
                                        Path.Combine(Environment.CurrentDirectory, "PromptTemplates", "viability_prompt.txt")
                                };

                                foreach (var candidate in candidates)
                                {
                                        if (File.Exists(candidate))
                                        {
                                                string loaded = File.ReadAllText(candidate, Encoding.UTF8);
                                                if (!string.IsNullOrWhiteSpace(loaded))
                                                        return loaded;
                                        }
                                }
                        }
                        catch
                        {
                                // Fall back to embedded prompt
                        }

                        return EmbeddedViabilityPrompt;
                }

        private static string ResolveDatabasePath()
        {
            foreach (var candidate in DbPathCandidates)
            {
                if (File.Exists(candidate))
                    return candidate;
            }

            foreach (var root in DbSearchRoots)
            {
                string? discovered = TryFindDatabaseInRoot(root);
                if (!string.IsNullOrWhiteSpace(discovered))
                    return discovered;
            }

            // If none is reachable now, keep UNC as default so connection errors point to the canonical path.
            return DbPathCandidates[0];
        }

        private static string? TryFindDatabaseInRoot(string root)
        {
            try
            {
                if (!Directory.Exists(root))
                    return null;

                var found = Directory
                    .EnumerateFiles(root, "PE_DB_v1.accdb", SearchOption.AllDirectories)
                    .FirstOrDefault();

                return string.IsNullOrWhiteSpace(found) ? null : found;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsOleDbProviderRegistered(string provider)
        {
            try
            {
                return Type.GetTypeFromProgID(provider, throwOnError: false) != null;
            }
            catch
            {
                return false;
            }
        }

        private static string ResolveOleDbProvider()
        {
            foreach (var provider in OleDbProviderCandidates)
            {
                if (IsOleDbProviderRegistered(provider))
                    return provider;
            }

            // Keep legacy default for backwards compatibility; load error will show actionable guidance.
            return OleDbProviderCandidates[1];
        }

        private static bool LooksLikeMissingAceProvider(Exception ex)
        {
            string message = ex.Message ?? string.Empty;
            return message.IndexOf("not registered", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("nao esta registrado", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("não está registrado", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string CreateStartupDiagnosticLogPath()
        {
            string baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string logFolder = Path.Combine(baseFolder, AppLogFolderName);
            Directory.CreateDirectory(logFolder);
            return Path.Combine(logFolder, $"startup_{DateTime.Now:yyyyMMdd_HHmmss}.log");
        }

        private static void AppendStartupDiagnostic(string logPath, string message)
        {
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now:O}] {message}{Environment.NewLine}", Encoding.UTF8);
            }
            catch
            {
                // Diagnostics must never crash the app.
            }
        }

        private string ResolveConnectionDatabasePath(string preferredSourcePath)
        {
            try
            {
                if (!File.Exists(preferredSourcePath))
                {
                    AppendStartupDiagnostic(startupDiagnosticLogPath, "Source DB path does not exist. Using source path directly.");
                    return preferredSourcePath;
                }

                string baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string cacheFolder = Path.Combine(baseFolder, "Pre_Launch_Tool", "Cache");
                Directory.CreateDirectory(cacheFolder);

                string localCopyPath = Path.Combine(cacheFolder, "PE_DB_v1_runtime.accdb");
                File.Copy(preferredSourcePath, localCopyPath, true);

                AppendStartupDiagnostic(startupDiagnosticLogPath, $"Database copied to local cache: {localCopyPath}");
                return localCopyPath;
            }
            catch (Exception ex)
            {
                AppendStartupDiagnostic(startupDiagnosticLogPath, $"Failed to create local DB cache copy. Fallback to source path. Error: {ex.Message}");
                return preferredSourcePath;
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            sourceDbPath = ResolveDatabasePath();
            activeOleDbProvider = ResolveOleDbProvider();
            startupDiagnosticLogPath = CreateStartupDiagnosticLogPath();
            dbPath = ResolveConnectionDatabasePath(sourceDbPath);
            // App performs read-only queries; opening in read mode avoids lock-file write issues on shared folders.
            connectionString = $@"Provider={activeOleDbProvider};Data Source={dbPath};Persist Security Info=False;Mode=Read;";
            AppendStartupDiagnostic(startupDiagnosticLogPath, "MainWindow initialized.");
            AppendStartupDiagnostic(startupDiagnosticLogPath, $"Process architecture: {(Environment.Is64BitProcess ? "x64" : "x86")}");
            AppendStartupDiagnostic(startupDiagnosticLogPath, $"Resolved source DB path: {sourceDbPath}");
            AppendStartupDiagnostic(startupDiagnosticLogPath, $"Source DB path exists: {File.Exists(sourceDbPath)}");
            AppendStartupDiagnostic(startupDiagnosticLogPath, $"Connection DB path: {dbPath}");
            AppendStartupDiagnostic(startupDiagnosticLogPath, $"Connection DB path exists: {File.Exists(dbPath)}");
            AppendStartupDiagnostic(startupDiagnosticLogPath, $"OLE DB provider: {activeOleDbProvider}");

            viabilityPrompt = LoadViabilityPromptTemplate();

            // Set username label and email address. Domain lookup may fail on some environments.
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.Current;

                    if (user != null)
                    {
                        LbUsernameMain.Content = $"{user.GivenName} {user.Surname}";
                        emailAddress = user.EmailAddress;
                    }
                }
            }
            catch
            {
                string fallbackUser = Environment.UserName;
                LbUsernameMain.Content = fallbackUser;
                emailAddress = string.Empty;
                AppendStartupDiagnostic(startupDiagnosticLogPath, "Domain user lookup failed. Fallback to local username.");
            }

            previewItems = new ObservableCollection<PreviewItem>();
            DgPreview.ItemsSource = previewItems;

            // Inicializar tabela de dados
            baseDataTable = new List<FleetDataRow>();

            // Disable all ComboBoxes except BU
            DisableComboBoxes();
            CbBU.IsEnabled = true;

            // Add event handlers
            CbBU.SelectionChanged += CbBU_SelectionChanged;
            CbProduto.SelectionChanged += CbProduto_SelectionChanged;
            CbRegiao.SelectionChanged += CbRegiao_SelectionChanged;
            CbMarca.SelectionChanged += CbMarca_SelectionChanged;
            CbModelo.SelectionChanged += CbModelo_SelectionChanged;
            CbAnoDe.SelectionChanged += CbAnoDe_SelectionChanged;
            CbAnoAte.SelectionChanged += CbAnoAte_SelectionChanged;
            CbTipoCombustivel.SelectionChanged += CbTipoCombustivel_SelectionChanged;
            CbExplanation.SelectionChanged += CbExplanation_SelectionChanged;
            CbCategoriaVeiculo.SelectionChanged += CbCategoriaVeiculo_SelectionChanged;

            // Competitor selection handlers to populate codes
            CbCompetitor.SelectionChanged += CbCompetitor_SelectionChanged;
            CbCompetitor_OE.SelectionChanged += CbCompetitor_OE_SelectionChanged;
            CbCompetitor_BK.SelectionChanged += CbCompetitor_BK_SelectionChanged;
            CbCompetitor.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler(CbCompetitor_TextChanged));
            CbCompetitor_OE.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler(CbCompetitor_OE_TextChanged));
            CbCompetitor_BK.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler(CbCompetitor_BK_TextChanged));

            // Ensure Office COM objects are released when window closes
            this.Closing += MainWindow_Closing;

            TxMarketShare.PreviewTextInput += TxMarketShare_PreviewTextInput;
            TxMarketShare.LostFocus += TxMarketShare_LostFocus;
            TxMarketShare.SelectionChanged += TxMarketShare_SelectionChanged;
            TxMarketShare.PreviewKeyDown += TxMarketShare_PreviewKeyDown;
            DataObject.AddPastingHandler(TxMarketShare, OnMarketSharePaste);
        }

        private void LbGotoWiki_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                string wikiUrl = "https://inside-share2.bosch.com/sites/01040098/_layouts/15/WopiFrame.aspx?sourcedoc=/sites/01040098/Shared%20Documents/Forms/AllItems.aspx&action=default";
                Process.Start(new ProcessStartInfo { FileName = wikiUrl, UseShellExecute = true });
            }
            catch { }
        }

        // Handler para digitação — aceita apenas dígitos (o "%" é gerenciado automaticamente)
        private void TxMarketShare_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char ch = e.Text.FirstOrDefault();
            if (!char.IsDigit(ch))
            {
                e.Handled = true;
                return;
            }

            var tb = (System.Windows.Controls.TextBox)sender;
            string currentNumber = tb.Text.Replace("%", "").Trim();
            int caretIndex = tb.CaretIndex;
            int selectionLength = tb.SelectionLength;

            int numberLength = currentNumber.Length;
            if (caretIndex > numberLength) caretIndex = numberLength;

            string selectedPart = (selectionLength > 0 && caretIndex + selectionLength <= numberLength)
                ? currentNumber.Substring(caretIndex, selectionLength)
                : string.Empty;

            string newNumber = currentNumber.Substring(0, caretIndex)
                             + e.Text
                             + currentNumber.Substring(caretIndex + selectedPart.Length);
 
            if (int.TryParse(newNumber, out int value))
            {
                if (value < 1 || value > 100)
                {
                    e.Handled = true;
                    return;
                }
            }
            else
            {
                e.Handled = true;
                return;
            }

            e.Handled = true;
            tb.Text = newNumber + "%";
            tb.CaretIndex = caretIndex + 1;
        }

        // Handler para colar texto
        private void OnMarketSharePaste(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = ((string)e.DataObject.GetData(typeof(string))).Replace("%", "").Trim();
                if (int.TryParse(text, out int value) && value >= 1 && value <= 100)
                {
                    var tb = (System.Windows.Controls.TextBox)sender;
                    tb.Text = value + "%";
                    tb.CaretIndex = tb.Text.Length - 1;
                }
            }
            e.CancelCommand();
        }

        // Handler para quando o campo perde foco — garante formato correto
        private void TxMarketShare_LostFocus(object sender, RoutedEventArgs e)
        {
            var tb = (System.Windows.Controls.TextBox)sender;
            string raw = tb.Text.Replace("%", "").Trim();

            if (string.IsNullOrEmpty(raw))
            {
                tb.Text = string.Empty;
                return;
            }

            if (int.TryParse(raw, out int value))
            {
                if (value < 1) value = 1;
                if (value > 100) value = 100;
                tb.Text = value + "%";
            }
            else
            {
                tb.Text = string.Empty;
            }
        }

        // Handler para impedir que o cursor fique depois do %
        private void TxMarketShare_SelectionChanged(object sender, RoutedEventArgs e)
        {
            var tb = (System.Windows.Controls.TextBox)sender;
            if (tb.Text.EndsWith("%") && tb.CaretIndex > tb.Text.Length - 1)
            {
                tb.CaretIndex = tb.Text.Length - 1;
            }
        }

        // Handler para tecla Delete/Backspace — gerencia remoção correta
        private void TxMarketShare_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var tb = (System.Windows.Controls.TextBox)sender;
            string currentNumber = tb.Text.Replace("%", "").Trim();

            if (e.Key == Key.Back)
            {
                e.Handled = true;
                if (currentNumber.Length > 0 && tb.CaretIndex > 0)
                {
                    int pos = Math.Min(tb.CaretIndex, currentNumber.Length);
                    string newNumber = currentNumber.Remove(pos - 1, 1);
                    if (string.IsNullOrEmpty(newNumber))
                    {
                        tb.Text = string.Empty;
                        tb.CaretIndex = 0;
                    }
                    else
                    {
                        tb.Text = newNumber + "%";
                        tb.CaretIndex = pos - 1;
                    }
                }
            }
            else if (e.Key == Key.Delete)
            {
                e.Handled = true;
                int pos = Math.Min(tb.CaretIndex, currentNumber.Length);
                if (pos < currentNumber.Length)
                {
                    string newNumber = currentNumber.Remove(pos, 1);
                    if (string.IsNullOrEmpty(newNumber))
                    {
                        tb.Text = string.Empty;
                        tb.CaretIndex = 0;
                    }
                    else
                    {
                        tb.Text = newNumber + "%";
                        tb.CaretIndex = pos;
                    }
                }
            }
        }

        private async Task FirstLoadInfos()
        {
            try
            {
                AppendStartupDiagnostic(startupDiagnosticLogPath, "FirstLoadInfos started.");

                // Usar uma estratégia de pré-carregamento para todos os dados relevantes
                await PreLoadAllData();

                AppendStartupDiagnostic(startupDiagnosticLogPath, $"PreLoadAllData completed. Fleet rows: {baseDataTable.Count}");

                if (baseDataTable.Count == 0 || ((List<string>)QueryCache.Instance.GetStaticData("BUs")).Count == 0)
                {
                    AppendStartupDiagnostic(startupDiagnosticLogPath, "Data loaded empty: one or more core datasets returned zero rows.");
                    MessageBox.Show(
                        "The application connected, but no data was loaded from Access.\n\n" +
                        $"Source DB path: {sourceDbPath}\n" +
                        $"Connection DB path: {dbPath}\n" +
                        $"OLE DB provider: {activeOleDbProvider}\n\n" +
                        "A diagnostic log was saved at:\n" +
                        startupDiagnosticLogPath,
                        "Access data not loaded",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }

                // Agora populamos todos os controles a partir dos dados pré-carregados
                PopulateControlsFromPreloadedData();
                AppendStartupDiagnostic(startupDiagnosticLogPath, "PopulateControlsFromPreloadedData completed.");
            }
            catch (Exception ex)
            {
                AppendStartupDiagnostic(startupDiagnosticLogPath, $"FirstLoadInfos failed: {ex}");
                string providerHint = LooksLikeMissingAceProvider(ex)
                    ? "\n\nPossivel causa: Microsoft Access Database Engine (ACE) nao instalado/registrado neste PC.\nPara este executavel (x64), instale o Access Database Engine 2016 x64."
                    : string.Empty;

                MessageBox.Show(
                    $"Erro ao carregar dados iniciais: {ex.Message}\n\nProvider OLE DB em uso:\n{activeOleDbProvider}\n\nCaminho da base de origem:\n{sourceDbPath}\n\nCaminho da base em uso na conexao:\n{dbPath}\n\nValide o acesso de rede ao caminho UNC:\n{DbPathCandidates[0]}{providerHint}\n\nLog de diagnostico:\n{startupDiagnosticLogPath}",
                    "Erro ao carregar bases",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task UpdateLoadingStatusAsync(string status, int? progressPercent = null)
        {
            try
            {
                LbLoadingStatus.Content = status;
                if (progressPercent.HasValue)
                {
                    int normalized = Math.Max(0, Math.Min(100, progressPercent.Value));
                    PbLoadingBases.Value = normalized;
                }
                await Dispatcher.Yield(DispatcherPriority.Background);
            }
            catch
            {
                // Keep loading flow resilient even if UI update fails.
            }
        }

        // Método para pré-carregar todos os dados necessários
        private async Task PreLoadAllData()
        {
            List<string> listBUs = new List<string>();
            List<string> listProdutos = new List<string>();
            List<string> listRegioes = new List<string>();
            List<EposDataItem> listEposDataItems = new List<EposDataItem>();
            AppendStartupDiagnostic(startupDiagnosticLogPath, "PreLoadAllData started.");

            await UpdateLoadingStatusAsync("Connecting to database...", 10);
            
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                await con.OpenAsync();

                // Carregar caminho do template de relatório
                await UpdateLoadingStatusAsync("Loading settings...", 20);
                using (OleDbCommand com = new OleDbCommand("SELECT TEMPLATE_PATH FROM SETTINGS", con))
                {
                    using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            templatePath = reader[0].ToString();
                        }
                    }
                }

                // Carregar BUs e Produtos da tabela epos_data
                await UpdateLoadingStatusAsync("Loading BU and products...", 35);
                using (OleDbCommand com = new OleDbCommand("SELECT DISTINCT BU, PRODUCT FROM epos_data", con))
                {
                    using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            string bu = reader[0].ToString();
                            string product = reader[1].ToString();

                            listEposDataItems.Add(new EposDataItem { BU = bu, Produto = product });

                            if (!string.IsNullOrEmpty(bu) && !listBUs.Contains(bu))
                                listBUs.Add(bu);
                        }
                    }
                }

                // Carregar Regiões da tabela epos_data
                await UpdateLoadingStatusAsync("Loading regions...", 50);
                using (OleDbCommand com = new OleDbCommand("SELECT DISTINCT COUNTRY FROM epos_data", con))
                {
                    using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            string country = reader[0].ToString();
                            if (!string.IsNullOrEmpty(country))
                                listRegioes.Add(country);
                        }
                    }
                }

                // A parte mais importante: pré-carregar TODOS os dados da base_frota_total para manipulação em memória
                await UpdateLoadingStatusAsync("Loading fleet database...", 75);
                using (OleDbCommand com = new OleDbCommand(
                    "SELECT SHORT, COUNTRY, FAS_POPULATION, BRAND, VEHICLE_TYPE, DATA, ENGINE_INFO, FUEL_TYPE, EXPLANATION, V_CLASS FROM base_frota_total", con))
                {
                    com.CommandTimeout = 120; // 2 minutos para esta operação crítica

                    using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            FleetDataRow row = new FleetDataRow
                            {
                                Short = reader.IsDBNull(0) ? null : reader.GetString(0), // SHORT might be null
                                Country = reader.IsDBNull(1) ? null : reader.GetString(1), // COUNTRY might be null
                                FasPopulation = reader.IsDBNull(2) ? 0 : reader.GetDouble(2), // Handle DBNull for population, default to 0
                                                                                             // If FAS_POPULATION is actually a long or double in Access:
                                                                                             // FasPopulation = reader.IsDBNull(2) ? 0L : reader.GetInt64(2), // For long integer
                                                                                             // FasPopulation = reader.IsDBNull(2) ? 0.0 : reader.GetDouble(2), // For double/single
                                Brand = reader.IsDBNull(3) ? null : reader.GetString(3),
                                VehicleType = reader.IsDBNull(4) ? null : reader.GetString(4),
                                Data = reader.IsDBNull(5) ? null : reader.GetString(5),
                                EngineInfo = reader.IsDBNull(6) ? null : reader.GetString(6),
                                FuelType = reader.IsDBNull(7) ? null : reader.GetString(7),
                                Explanation = reader.IsDBNull(8) ? null : reader.GetString(8),
                                VClass = reader.IsDBNull(9) ? null : reader.GetString(9)
                            };

                            // Parse anos
                            if (!string.IsNullOrEmpty(row.Data))
                            {
                                var anos = row.Data.Split(" -> ");
                                if (anos.Length == 2)
                                {
                                    row.AnoInicio = anos[0];
                                    row.AnoFim = anos[1];
                                }
                            }

                            baseDataTable.Add(row);
                        }
                    }
                }

                // Load competitor table cod_concorrente (MARKENBEZ -> TW -> TWNR_VERD)
                try
                {
                    await UpdateLoadingStatusAsync("Loading competitor database...", 90);
                    competitorTable.Clear();
                    using (OleDbCommand com = new OleDbCommand("SELECT MARKENBEZ, TW, TWNR_VERD FROM cod_concorrente", con))
                    using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            string? markenbez = reader.IsDBNull(0) ? null : reader.GetString(0);
                            string? tw = reader.IsDBNull(1) ? null : reader.GetString(1);
                            string? verd = reader.IsDBNull(2) ? null : reader.GetString(2);

                            if (string.IsNullOrWhiteSpace(markenbez))
                                continue;

                            competitorTable.Add(new CompetitorRow(markenbez.Trim(), (tw ?? string.Empty).Trim(), string.IsNullOrWhiteSpace(verd) ? null : verd.Trim()));
                        }
                    }
                }
                catch
                {
                    competitorTable.Clear();
                }
            }

            // Configurar os dados para uso posterior
            await UpdateLoadingStatusAsync("Finalizing load...", 98);
            QueryCache.Instance.SetStaticData("BUs", listBUs.OrderBy(b => b).ToList());
            QueryCache.Instance.SetStaticData("AllEposDataItems", listEposDataItems.OrderBy(item => item.BU).ThenBy(item => item.Produto).ToList());
            QueryCache.Instance.SetStaticData("Regioes", listRegioes);
            AppendStartupDiagnostic(startupDiagnosticLogPath, $"Data summary: BUs={listBUs.Count}, EPOS={listEposDataItems.Count}, Regions={listRegioes.Count}, Fleet={baseDataTable.Count}, Competitors={competitorTable.Count}");

            // Populate competitor combos after data load
            PopulateCompetitorCombos();
            await UpdateLoadingStatusAsync("Databases loaded", 100);
            AppendStartupDiagnostic(startupDiagnosticLogPath, "PreLoadAllData finished.");
        }

        // Método para popular os controles a partir dos dados pré-carregados
        private void PopulateControlsFromPreloadedData()
        {
            // Dados de epos_data
            CbBU.ItemsSource = (List<string>)QueryCache.Instance.GetStaticData("BUs");
            //CbProduto.ItemsSource = (List<string>)QueryCache.Instance.GetStaticData("Produtos");

            // Regiões
            var regioesCheckboxes = CreateCheckBoxItems((List<string>)QueryCache.Instance.GetStaticData("Regioes"), CbRegiao, true, false);
            CbRegiao.ItemsSource = regioesCheckboxes;

            // Extrair dados únicos dos dados pré-carregados
            var marcas = baseDataTable.Select(r => r.Brand).Distinct().OrderBy(m => m).ToList();
            var modelos = baseDataTable.Select(r => r.VehicleType).Distinct().OrderBy(m => m).ToList();
            var anosInicio = baseDataTable.Select(r => r.AnoInicio).Where(a => !string.IsNullOrEmpty(a)).Distinct().OrderBy(a => a).ToList();
            var anosFim = baseDataTable.Select(r => r.AnoFim).Where(a => !string.IsNullOrEmpty(a)).Distinct().OrderByDescending(a => a).ToList();
            var combustiveis = baseDataTable.Select(r => r.FuelType).Distinct().OrderBy(c => c).ToList();
            var explanations = baseDataTable.Select(r => r.Explanation).Distinct().OrderBy(e => e).ToList();
            var categorias = baseDataTable.Select(r => r.VClass).Distinct().OrderBy(c => c).ToList();

            // Criar CheckBoxes para os campos que precisam de múltipla seleção
            var marcasCheckboxes = CreateCheckBoxItems(marcas, CbMarca);
            var modelosCheckboxes = CreateCheckBoxItems(modelos, CbModelo);
            var combustiveisCheckboxes = CreateCheckBoxItems(combustiveis, CbTipoCombustivel);
            var explanationsCheckboxes = CreateCheckBoxItems(explanations, CbExplanation);
            var categoriasCheckboxes = CreateCheckBoxItems(categorias, CbCategoriaVeiculo);

            // Popular os ComboBoxes
            CbMarca.ItemsSource = marcasCheckboxes;
            CbModelo.ItemsSource = modelosCheckboxes;
            CbAnoDe.ItemsSource = anosInicio;
            CbAnoAte.ItemsSource = anosFim;
            CbTipoCombustivel.ItemsSource = combustiveisCheckboxes;
            CbExplanation.ItemsSource = explanationsCheckboxes;
            CbCategoriaVeiculo.ItemsSource = categoriasCheckboxes;
        }

        // Method to disable all ComboBoxes except BU
        private void DisableComboBoxes()
        {
            CbProduto.IsEnabled = false;
            CbRegiao.IsEnabled = false;
            CbMarca.IsEnabled = false;
            CbModelo.IsEnabled = false;
            CbAnoDe.IsEnabled = false;
            CbAnoAte.IsEnabled = false;
            CbTipoCombustivel.IsEnabled = false;
            CbExplanation.IsEnabled = false;
            CbCategoriaVeiculo.IsEnabled = false;
        }

        // Method to enable the next ComboBox in the sequence
        private void EnableNextComboBox(ComboBox currentComboBox)
        {
            if (currentComboBox == CbBU && CbBU.SelectedItem != null)
            {
                CbProduto.IsEnabled = true;
            }
            else if (currentComboBox == CbProduto && CbProduto.SelectedItem != null)
            {
                CbRegiao.IsEnabled = true;
            }
            else if (currentComboBox == CbRegiao && SpecialHandleCB(CbRegiao) != string.Empty)
            {
                CbMarca.IsEnabled = true;
            }
            else if (currentComboBox == CbMarca && SpecialHandleCB(CbMarca) != string.Empty)
            {
                CbModelo.IsEnabled = true;
            }
            else if (currentComboBox == CbModelo && SpecialHandleCB(CbModelo) != string.Empty)
            {
                CbAnoDe.IsEnabled = true;
            }
            else if (currentComboBox == CbAnoDe && CbAnoDe.SelectedItem != null)
            {
                CbAnoAte.IsEnabled = true;
            }
            else if (currentComboBox == CbAnoAte && CbAnoAte.SelectedItem != null)
            {
                CbTipoCombustivel.IsEnabled = true;
            }
            else if (currentComboBox == CbTipoCombustivel && SpecialHandleCB(CbTipoCombustivel) != string.Empty)
            {
                CbExplanation.IsEnabled = true;
            }
            else if (currentComboBox == CbExplanation && SpecialHandleCB(CbExplanation) != string.Empty)
            {
                CbCategoriaVeiculo.IsEnabled = true;
            }
        }

        // Event handlers for SelectionChanged events
        private void CbBU_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbBU.SelectedItem is string selectedBU && !string.IsNullOrEmpty(selectedBU))
            {
                List<EposDataItem> allEposDataItems = (List<EposDataItem>)QueryCache.Instance.GetStaticData("AllEposDataItems");

                var filteredProdutos = allEposDataItems.Where(item => item.BU == selectedBU).Select(item => item.Produto).Distinct().OrderBy(p => p).ToList();

                CbProduto.ItemsSource = filteredProdutos;
            }
            EnableNextComboBox(CbBU);
        }

        private void CbProduto_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbProduto);
        }

        private void CbRegiao_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbRegiao);
        }

        private void RegionCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (suppressRegionRestrictionMessage)
                return;

            if (sender is not CheckBox cb)
                return;

            if (IsSelectAllCheckBox(cb))
                return;

            string selectedRegion = cb.Content?.ToString()?.Trim() ?? string.Empty;
            if (string.Equals(selectedRegion, AllowedRegion, StringComparison.OrdinalIgnoreCase))
                return;

            try
            {
                suppressRegionRestrictionMessage = true;
                cb.IsChecked = false;
            }
            finally
            {
                suppressRegionRestrictionMessage = false;
            }

            MessageBox.Show(
                "Nesta versão só constam bases do Brazil.",
                "Região indisponível",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void CbMarca_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbMarca);
        }

        private void CbModelo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbModelo);
        }

        private void CbAnoDe_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbAnoDe);
            FilterAllFields();
        }

        private void CbAnoAte_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbAnoAte);
            FilterAllFields();
        }

        private void CbTipoCombustivel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbTipoCombustivel);
        }

        private void CbExplanation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbExplanation);
        }

        private void CbCategoriaVeiculo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EnableNextComboBox(CbCategoriaVeiculo);
        }

        // Método ultra-otimizado para filtrar todos os campos com base na marca e/ou modelo
        private async void FilterAllFields()
        {
            if (isFilteringVehicleInputs)
                return;

            isFilteringVehicleInputs = true;
            Mouse.SetCursor(Cursors.Wait);

            try
            {
                var selectedMarcas = GetSelectedValuesFromCheckBoxComboBox(CbMarca)
                    .OrderBy(s => s)
                    .ToList();

                var selectedModelosAnterior = GetSelectedValuesFromCheckBoxComboBox(CbModelo);
                var selectedCombustivelAnterior = GetSelectedValuesFromCheckBoxComboBox(CbTipoCombustivel);
                var selectedExplanationAnterior = GetSelectedValuesFromCheckBoxComboBox(CbExplanation);
                var selectedCategoriaAnterior = GetSelectedValuesFromCheckBoxComboBox(CbCategoriaVeiculo);
                string selectedAnoDeAnterior = CbAnoDe.SelectedItem?.ToString() ?? CbAnoDe.Text;
                string selectedAnoAteAnterior = CbAnoAte.SelectedItem?.ToString() ?? CbAnoAte.Text;

                if (selectedMarcas.Count == 0)
                {
                    CbModelo.ItemsSource = null;
                    CbAnoDe.ItemsSource = null;
                    CbAnoAte.ItemsSource = null;
                    CbTipoCombustivel.ItemsSource = null;
                    CbExplanation.ItemsSource = null;
                    CbCategoriaVeiculo.ItemsSource = null;
                    CbModelo.Text = string.Empty;
                    CbAnoDe.Text = string.Empty;
                    CbAnoAte.Text = string.Empty;
                    CbTipoCombustivel.Text = string.Empty;
                    CbExplanation.Text = string.Empty;
                    CbCategoriaVeiculo.Text = string.Empty;
                    return;
                }

                var byBrand = baseDataTable
                    .Where(row => selectedMarcas.Contains(row.Brand, StringComparer.OrdinalIgnoreCase))
                    .ToList();

                var modelosDisponiveis = byBrand
                    .Select(row => row.VehicleType)
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(v => v)
                    .ToList();

                CbModelo.ItemsSource = CreateCheckBoxItems(modelosDisponiveis, CbModelo);
                foreach (var cb in CbModelo.Items.OfType<CheckBox>())
                {
                    if (IsSelectAllCheckBox(cb)) continue;
                    cb.IsChecked = selectedModelosAnterior.Contains(cb.Content?.ToString() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
                }
                CbModelo.Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox(CbModelo));

                var selectedModelos = GetSelectedValuesFromCheckBoxComboBox(CbModelo);
                var byBrandAndModel = byBrand;
                if (selectedModelos.Count > 0)
                {
                    byBrandAndModel = byBrandAndModel
                        .Where(row => selectedModelos.Contains(row.VehicleType, StringComparer.OrdinalIgnoreCase))
                        .ToList();
                }

                var anosDeDisponiveis = byBrandAndModel
                    .Select(row => row.AnoInicio)
                    .Where(a => !string.IsNullOrWhiteSpace(a))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(a => a)
                    .ToList();

                var anosAteDisponiveis = byBrandAndModel
                    .Select(row => row.AnoFim)
                    .Where(a => !string.IsNullOrWhiteSpace(a))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(a => a)
                    .ToList();

                CbAnoDe.ItemsSource = anosDeDisponiveis;
                CbAnoAte.ItemsSource = anosAteDisponiveis;

                CbAnoDe.SelectedItem = anosDeDisponiveis.Contains(selectedAnoDeAnterior) ? selectedAnoDeAnterior : null;
                CbAnoAte.SelectedItem = anosAteDisponiveis.Contains(selectedAnoAteAnterior) ? selectedAnoAteAnterior : null;

                int.TryParse(CbAnoDe.SelectedItem?.ToString(), out int anoDeSelecionado);
                int.TryParse(CbAnoAte.SelectedItem?.ToString(), out int anoAteSelecionado);

                var byBrandModelAndYear = byBrandAndModel
                    .Where(row =>
                    {
                        bool okAnoDe = true;
                        bool okAnoAte = true;

                        if (anoDeSelecionado > 0)
                        {
                            okAnoDe = !string.IsNullOrWhiteSpace(row.AnoFim) && int.TryParse(row.AnoFim, out int anoFim) && anoFim >= anoDeSelecionado;
                        }

                        if (anoAteSelecionado > 0)
                        {
                            okAnoAte = !string.IsNullOrWhiteSpace(row.AnoInicio) && int.TryParse(row.AnoInicio, out int anoInicio) && anoInicio <= anoAteSelecionado;
                        }

                        return okAnoDe && okAnoAte;
                    })
                    .ToList();

                var combustiveisDisponiveis = byBrandModelAndYear
                    .Select(row => row.FuelType)
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(c => c)
                    .ToList();

                CbTipoCombustivel.ItemsSource = CreateCheckBoxItems(combustiveisDisponiveis, CbTipoCombustivel);
                foreach (var cb in CbTipoCombustivel.Items.OfType<CheckBox>())
                {
                    if (IsSelectAllCheckBox(cb)) continue;
                    cb.IsChecked = selectedCombustivelAnterior.Contains(cb.Content?.ToString() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
                }
                CbTipoCombustivel.Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox(CbTipoCombustivel));

                var selectedCombustivel = GetSelectedValuesFromCheckBoxComboBox(CbTipoCombustivel);
                var byBrandModelYearAndFuel = byBrandModelAndYear;
                if (selectedCombustivel.Count > 0)
                {
                    byBrandModelYearAndFuel = byBrandModelYearAndFuel
                        .Where(row => selectedCombustivel.Contains(row.FuelType, StringComparer.OrdinalIgnoreCase))
                        .ToList();
                }

                var explanationsDisponiveis = byBrandModelYearAndFuel
                    .Select(row => row.Explanation)
                    .Where(e => !string.IsNullOrWhiteSpace(e))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(e => e)
                    .ToList();

                CbExplanation.ItemsSource = CreateCheckBoxItems(explanationsDisponiveis, CbExplanation);
                foreach (var cb in CbExplanation.Items.OfType<CheckBox>())
                {
                    if (IsSelectAllCheckBox(cb)) continue;
                    cb.IsChecked = selectedExplanationAnterior.Contains(cb.Content?.ToString() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
                }
                CbExplanation.Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox(CbExplanation));

                var selectedExplanations = GetSelectedValuesFromCheckBoxComboBox(CbExplanation);
                var byBrandModelYearFuelAndExplanation = byBrandModelYearAndFuel;
                if (selectedExplanations.Count > 0)
                {
                    byBrandModelYearFuelAndExplanation = byBrandModelYearFuelAndExplanation
                        .Where(row => selectedExplanations.Contains(row.Explanation, StringComparer.OrdinalIgnoreCase))
                        .ToList();
                }

                var categoriasDisponiveis = byBrandModelYearFuelAndExplanation
                    .Select(row => row.VClass)
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(c => c)
                    .ToList();

                CbCategoriaVeiculo.ItemsSource = CreateCheckBoxItems(categoriasDisponiveis, CbCategoriaVeiculo);
                foreach (var cb in CbCategoriaVeiculo.Items.OfType<CheckBox>())
                {
                    if (IsSelectAllCheckBox(cb)) continue;
                    cb.IsChecked = selectedCategoriaAnterior.Contains(cb.Content?.ToString() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
                }
                CbCategoriaVeiculo.Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox(CbCategoriaVeiculo));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao filtrar campos: {ex.Message}");
            }
            finally
            {
                Mouse.SetCursor(null);
                isFilteringVehicleInputs = false;
            }
        }

        private async void BtPesquisar_Click(object sender, RoutedEventArgs e)
        {
            string mode = GetActiveSearchMode();

            // Se o usuário digitou BKs, filtrar apenas por eles (apenas no modo BK)
            string bksInput = TxSearchBK_BK.Text;
            var bksList = bksInput?.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(bk => !string.IsNullOrWhiteSpace(bk)).ToList() ?? new List<string>();

            // Se o usuário digitou OEs, buscar SHORTs correspondentes na base_OE (apenas no modo OE)
            string oesInput = TxSearchOENumber_OE.Text;
            var oesList = oesInput?.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(oe => !string.IsNullOrWhiteSpace(oe)).ToList() ?? new List<string>();

            ComboBox activeCodeCombo = CbCompetitorCode;
            if (mode == "OE_SEARCH")
                activeCodeCombo = CbCompetitorCode_OE;
            else if (mode == "BK_SEARCH")
                activeCodeCombo = CbCompetitorCode_BK;

            var competitorCodesForSearch = GetSelectedCompetitorCodes(activeCodeCombo);
            if (mode == "OE_SEARCH" && competitorCodesForSearch.Count > 0)
            {
                // Allow searching by multiple competitor codes selected/typed in Code field.
                oesList = oesList
                    .Concat(competitorCodesForSearch)
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            List<string> shortsFromOe = new List<string>();
            if (mode == "OE_SEARCH" && oesList.Count > 0)
            {
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    await con.OpenAsync();
                    string oeWhere = string.Join(",", oesList.Select(oe => "'" + oe.Replace("'", "''") + "'"));
                    // Buscar SHORTs na tabela base_OE usando a coluna OE_number
                    string query = $"SELECT DISTINCT SHORT FROM base_OE WHERE OE_number IN ({oeWhere})";
                    using (OleDbCommand com = new OleDbCommand(query, con))
                    {
                        using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                if (!reader.IsDBNull(0))
                                    shortsFromOe.Add(reader.GetString(0));
                            }
                        }
                    }
                }
            }

            // Obtenha o valor digitado pelo usuário para Market Share
            string userDefinedMS = TxMarketShare.Text.Replace("%", "").Trim();
            double? userMS = null;
            if (double.TryParse(userDefinedMS, NumberStyles.Any, CultureInfo.InvariantCulture, out double ms))
                userMS = ms / 100.0; // Sempre inteiro de 1-100, converter para fração

            // Required Information é comum para todas as buscas.
            if (CbBU.Text == string.Empty || CbProduto.Text == string.Empty || CbRegiao.Text == string.Empty)
            {
                MessageBox.Show("Um ou mais campos obrigatórios foram deixados em branco.\nPreencha BU, Product e Region antes de realizar a pesquisa.", "Campo(s) em branco!", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (mode == "VEHICLE_SEARCH" &&
                (CbMarca.Text == string.Empty || CbModelo.Text == string.Empty || CbAnoDe.Text == string.Empty || CbAnoAte.Text == string.Empty ||
                 CbTipoCombustivel.Text == string.Empty || CbExplanation.Text == string.Empty || CbCategoriaVeiculo.Text == string.Empty))
            {
                MessageBox.Show("Preencha os campos da aba VEHICLE SEARCH antes de realizar a pesquisa.", "Campo(s) em branco!", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (mode == "OE_SEARCH" && oesList.Count == 0)
            {
                MessageBox.Show("Informe ao menos um código OE na aba OE SEARCH.", "Campo(s) em branco!", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (mode == "BK_SEARCH" && bksList.Count == 0)
            {
                MessageBox.Show("Informe ao menos um código BK na aba BK SEARCH.", "Campo(s) em branco!", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var activeCompetitor = GetActiveCompetitorSelection(mode);

            Mouse.SetCursor(Cursors.Wait);
            previewItems.Clear();
            LbPreviewBU.Content = CbBU.Text;
            LbPreviewProduto.Content = CbProduto.Text;

            try
            {
                IEnumerable<FleetDataRow> filteredFleetData = baseDataTable;
                if (mode == "BK_SEARCH")
                {
                    filteredFleetData = baseDataTable.Where(dr => bksList.Contains(dr.Short));
                }
                else if (mode == "OE_SEARCH")
                {
                    if (shortsFromOe.Count == 0)
                    {
                        MessageBox.Show("Nenhum SHORT foi encontrado para os códigos OE informados.", "Sem resultados", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    filteredFleetData = baseDataTable.Where(dr => shortsFromOe.Contains(dr.Short));
                }
                else
                {
                    var selectedBrands = GetSelectedValuesFromCheckBoxComboBox(CbMarca);
                    var selectedModels = GetSelectedValuesFromCheckBoxComboBox(CbModelo);
                    var selectedFuelTypes = GetSelectedValuesFromCheckBoxComboBox(CbTipoCombustivel);
                    var selectedExplanations = GetSelectedValuesFromCheckBoxComboBox(CbExplanation);
                    var selectedVehicleCategories = GetSelectedValuesFromCheckBoxComboBox(CbCategoriaVeiculo);
                    string selectedYearFrom = CbAnoDe.SelectedItem?.ToString();
                    string selectedYearTo = CbAnoAte.SelectedItem?.ToString();

                    if (selectedBrands.Any()) filteredFleetData = filteredFleetData.Where(dr => selectedBrands.Contains(dr.Brand));
                    if (selectedModels.Any()) filteredFleetData = filteredFleetData.Where(dr => selectedModels.Contains(dr.VehicleType));
                    if (selectedFuelTypes.Any()) filteredFleetData = filteredFleetData.Where(dr => selectedFuelTypes.Contains(dr.FuelType));
                    if (selectedExplanations.Any()) filteredFleetData = filteredFleetData.Where(dr => selectedExplanations.Contains(dr.Explanation));
                    if (selectedVehicleCategories.Any()) filteredFleetData = filteredFleetData.Where(dr => selectedVehicleCategories.Contains(dr.VClass));
                    if (!string.IsNullOrEmpty(selectedYearFrom) && int.TryParse(selectedYearFrom, out int yf)) filteredFleetData = filteredFleetData.Where(dr => !string.IsNullOrEmpty(dr.AnoFim) && int.Parse(dr.AnoFim) >= yf);
                    if (!string.IsNullOrEmpty(selectedYearTo) && int.TryParse(selectedYearTo, out int yt)) filteredFleetData = filteredFleetData.Where(dr => !string.IsNullOrEmpty(dr.AnoInicio) && int.Parse(dr.AnoInicio) <= yt);
                }

                List<FleetDataRow> finalFilteredFleetList = filteredFleetData.ToList();
                // Store for report generation (contains motorização / EngineInfo)
                lastFilteredFleetList = finalFilteredFleetList;

                if (finalFilteredFleetList.Count == 0)
                {
                    MessageBox.Show("No fleet data found for the selected criteria.");
                    return;
                }

                HashSet<string> distinctVClassesFromFleet = new HashSet<string>(finalFilteredFleetList.Select(dr => dr.VClass).Where(s => !string.IsNullOrEmpty(s)));
                HashSet<string> distinctCountriesFromFleet = new HashSet<string>(finalFilteredFleetList.Select(dr => dr.Country).Where(s => !string.IsNullOrEmpty(s)));

                List<string> eposConditions = new List<string>();
                AddEqualsCondition(eposConditions, "BU", CbBU.SelectedItem?.ToString());
                AddEqualsCondition(eposConditions, "PRODUCT", CbProduto.SelectedItem?.ToString());
                var selectedRegionsUI = GetSelectedValuesFromCheckBoxComboBox(CbRegiao);
                var effectiveEposCountries = selectedRegionsUI.Any()
                    ? distinctCountriesFromFleet.Intersect(selectedRegionsUI, StringComparer.OrdinalIgnoreCase).ToList()
                    : distinctCountriesFromFleet.ToList();
                AddInCondition(eposConditions, "COUNTRY", effectiveEposCountries);
                var selectedCategoryUI = GetSelectedValuesFromCheckBoxComboBox(CbCategoriaVeiculo);
                var effectiveEposCategories = selectedCategoryUI.Any()
                    ? distinctVClassesFromFleet.Intersect(selectedCategoryUI, StringComparer.OrdinalIgnoreCase).ToList()
                    : distinctVClassesFromFleet.ToList();
                AddInCondition(eposConditions, "CATEGORY", effectiveEposCategories);
                string eposWhereClause = eposConditions.Any() ? string.Join(" AND ", eposConditions) : "1=1";
                string eposQuery = $"SELECT CATEGORY, COUNTRY, EXCHANGE_RATE_YEAR, UNIT_REPLACEMENT_SHARE, AVERAGE_REPLACEMENT_QUANTITY, MS FROM epos_data WHERE {eposWhereClause}";

                Dictionary<(string Category, string Country), EposDataItem> eposLookup = new Dictionary<(string Category, string Country), EposDataItem>();
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    await con.OpenAsync();
                    using (OleDbCommand com = new OleDbCommand(eposQuery, con))
                    {
                        using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                var item = new EposDataItem
                                {
                                    Category = reader.GetString(0),
                                    Country = reader.GetString(1),
                                    ExchangeRateYear = reader.GetDouble(2),
                                    UnitReplacementShare = reader.GetDouble(3),
                                    AverageReplacementQuantity = reader.GetDouble(4),
                                    MS = reader.GetDouble(5)
                                };
                                eposLookup.TryAdd((item.Category, item.Country), item);
                            }
                        }
                    }
                }

                Dictionary<string, string> shortToOeMap = new();

                var shortsToQuery = new List<string>();
                if (shortsFromOe.Count > 0)
                {
                    shortsToQuery.AddRange(shortsFromOe.Select(s => s.Trim()));
                }
                if (bksList.Count > 0)
                {
                    shortsToQuery.AddRange(bksList.Select(s => s.Trim()));
                }

                shortsToQuery = shortsToQuery.Where(s => !string.IsNullOrEmpty(s)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                if (shortsToQuery.Count > 0)
                {
                    using (OleDbConnection con = new OleDbConnection(connectionString))
                    {
                        await con.OpenAsync();
                        string shortWhere = string.Join(",", shortsToQuery.Select(s => "'" + s.Replace("'", "''") + "'"));
                        string mapQuery = $"SELECT SHORT, OE_number FROM base_OE WHERE SHORT IN ({shortWhere})";
                        using (OleDbCommand com = new OleDbCommand(mapQuery, con))
                        using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                string s = reader.IsDBNull(0) ? null : reader.GetString(0);
                                string oe = reader.IsDBNull(1) ? null : reader.GetString(1);
                                if (string.IsNullOrEmpty(s)) continue;
                                s = s.Trim();
                                if (string.IsNullOrEmpty(oe)) oe = string.Empty;

                                if (shortToOeMap.ContainsKey(s))
                                {
                                    var existing = shortToOeMap[s];
                                    var parts = existing.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(p => p.Trim()).Where(p => p.Length > 0).ToList();
                                    if (!parts.Contains(oe) && !string.IsNullOrEmpty(oe))
                                        shortToOeMap[s] = string.IsNullOrEmpty(existing) ? oe : existing + ";" + oe;
                                }
                                else
                                {
                                    shortToOeMap[s] = oe;
                                }
                            }
                        }
                    }
                }

                foreach (var fleetRow in finalFilteredFleetList)
                {
                    string oeValue = null;
                    if (shortToOeMap.TryGetValue(fleetRow.Short, out var oeFound))
                        oeValue = oeFound;
                    if (eposLookup.TryGetValue((fleetRow.VClass, fleetRow.Country), out EposDataItem eposData))
                    {
                        double marketShareToUse = userMS ?? eposData.MS;
                        double potential = Math.Round((double)fleetRow.FasPopulation * marketShareToUse * eposData.ExchangeRateYear *
                                               eposData.UnitReplacementShare * eposData.AverageReplacementQuantity, 0);

                        previewItems.Add(new PreviewItem
                        {
                            BK = fleetRow.Short,
                            OENumber = oeValue,
                            Manufacturer = fleetRow.Brand,
                            VehicleName = fleetRow.VehicleType,
                            ApplicationPeriod = fleetRow.Data,
                            FuelType = fleetRow.FuelType,
                            EngineType = fleetRow.EngineInfo,
                            Explanation = fleetRow.Explanation,
                            Fleet = fleetRow.FasPopulation.ToString(),
                            Region = fleetRow.Country,
                            Potential = potential.ToString("N2"),
                            UserDefinedMarketShare = userDefinedMS,
                            CompetitorName = activeCompetitor.Name,
                            CompetitorCode = activeCompetitor.Code,
                        });
                    }
                    else
                    {
                        previewItems.Add(new PreviewItem
                        {
                            BK = fleetRow.Short,
                            OENumber = oeValue,
                            Manufacturer = fleetRow.Brand,
                            VehicleName = fleetRow.VehicleType,
                            ApplicationPeriod = fleetRow.Data,
                            FuelType = fleetRow.FuelType,
                            EngineType = fleetRow.EngineInfo,
                            Explanation = fleetRow.Explanation,
                            Fleet = fleetRow.FasPopulation.ToString(),
                            Region = fleetRow.Country,
                            Potential = "N/A (No EPOS Data)",
                            UserDefinedMarketShare = userDefinedMS,
                            CompetitorName = activeCompetitor.Name,
                            CompetitorCode = activeCompetitor.Code,
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during calculation: {ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.SetCursor(null);
            }
            
            // Generate report
            try
            {
                string projectID = $"{CbBU.Text}_{CbProduto.Text}_{DateTime.Now:yyyyMMddHHmmss}";
                Mouse.SetCursor(Cursors.Wait);
                await GenerateWordReportAsync(projectID);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gerar o documento Word: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.SetCursor(null);
            }
        }

        private string GetActiveSearchMode()
        {
            try
            {
                if (TcSearchTabs != null)
                {
                    switch (TcSearchTabs.SelectedIndex)
                    {
                        case 0: return "VEHICLE_SEARCH";
                        case 1: return "OE_SEARCH";
                        case 2: return "BK_SEARCH";
                    }
                }
            }
            catch { }

            return "VEHICLE_SEARCH";
        }

        private (string Name, string Code) GetActiveCompetitorSelection(string mode)
        {
            ComboBox nameCombo = CbCompetitor;
            ComboBox codeCombo = CbCompetitorCode;

            if (mode == "OE_SEARCH")
            {
                nameCombo = CbCompetitor_OE;
                codeCombo = CbCompetitorCode_OE;
            }
            else if (mode == "BK_SEARCH")
            {
                nameCombo = CbCompetitor_BK;
                codeCombo = CbCompetitorCode_BK;
            }

            string competitorName = nameCombo?.Text?.Trim() ?? string.Empty;
            string competitorCode = GetCompetitorCodeForPrompt(codeCombo);

            return (competitorName, competitorCode);
        }

        private void BtResetAllInputs_Click(object sender, RoutedEventArgs e)
        {
            ResetAllInputs();
        }

        private void BtConfirmInfo_Click(object sender, RoutedEventArgs e)
        {
            DgPreview.IsReadOnly = true;
            BtConfirmInfo.IsEnabled = false;
            BtGenerateReport.IsEnabled = true;
        }

        private async void BtGenerateReport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
            {
                MessageBox.Show($"Template de Excel não encontrado. Path: '{templatePath ?? "(null)"}'", "Template ausente", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                Mouse.SetCursor(Cursors.Wait);

                int counter = 0, updatedCounter = counter, defaultRows = 12;
                string indexador = string.Empty, projectID = string.Empty;

                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    await con.OpenAsync();
                    using (OleDbCommand com = new OleDbCommand("SELECT COUNTER FROM SETTINGS", con))
                    {
                        using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync()) counter = Convert.ToInt32(reader[0]);
                        }
                    }
                    using (OleDbCommand com = new OleDbCommand("UPDATE [SETTINGS] SET [COUNTER] = @counter WHERE [ID] = @id", con))
                    {
                        updatedCounter = counter + 1;
                        OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("@counter", updatedCounter), new OleDbParameter("@id", 1) };
                        com.Parameters.AddRange(parameters);
                        com.Connection = con;
                        com.ExecuteNonQuery();
                    }
                }

                if (updatedCounter < 10) indexador = $"00000{updatedCounter}";
                else if (counter < 100) indexador = $"0000{updatedCounter}";
                else if (counter < 1000) indexador = $"000{updatedCounter}";
                else if (counter < 10000) indexador = $"00{updatedCounter}";
                else if (counter < 100000) indexador = $"0{updatedCounter}";
                else indexador = updatedCounter.ToString();

                projectID = $"{CbBU.Text}_{CbProduto.Text}_{DateTime.Today:yyyyMMdd}_{indexador}";

                // Use EPPlus to modify template and save as .xlsm/.xlsx
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var templateFile = new FileInfo(templatePath);
                using (var package = new OfficeOpenXml.ExcelPackage(templateFile))
                {
                    var worksheet = package.Workbook.Worksheets.First();

                    worksheet.Cells[4, 3].Value = emailAddress; // C4
                    worksheet.Cells[4, 8].Value = DateTime.Today.ToShortDateString(); // H4
                    worksheet.Cells[4, 10].Value = CbBU.Text; // J4
                    worksheet.Cells[4, 12].Value = CbProduto.Text; // L4
                    worksheet.Cells[4, 14].Value = CbRegiao.Text; // N4
                    worksheet.Cells[6, 3].Value = projectID; // C6
                    worksheet.Cells[6, 8].Value = string.Empty; // H6

                    // add more rows if needed at row 13
                    if (DgPreview.Items.Count > defaultRows)
                    {
                        for (int i = 0; i < DgPreview.Items.Count - defaultRows; i++)
                        {
                            worksheet.InsertRow(13, 1);
                            if (i % 2 == 0)
                            {
                                using (var range = worksheet.Cells[13, 2, 13, 22])
                                {
                                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                }
                            }
                        }
                    }

                    for (int i = 0; i < DgPreview.Items.Count; i++)
                    {
                        int currentRow = 11 + i;
                        var item = (PreviewItem)DgPreview.Items[i];
                        worksheet.Cells[currentRow, 2].Value = item.BK; // B
                        worksheet.Cells[currentRow, 3].Value = item.Manufacturer; // C
                        worksheet.Cells[currentRow, 4].Value = item.VehicleName; // D
                        worksheet.Cells[currentRow, 5].Value = item.ApplicationPeriod; // E
                        worksheet.Cells[currentRow, 6].Value = item.FuelType; // F
                        worksheet.Cells[currentRow, 7].Value = item.EngineType; // G
                        worksheet.Cells[currentRow, 8].Value = item.Explanation; // H
                        worksheet.Cells[currentRow, 9].Value = item.Fleet; // I
                        worksheet.Cells[currentRow, 10].Value = item.Region; // J
                        worksheet.Cells[currentRow, 11].Value = item.Potential; // K
                    }

                    // (Printer orientation left as template default) 

                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string outputPath = Path.Combine(desktopPath, projectID + Path.GetExtension(templatePath));
                    package.SaveAs(new FileInfo(outputPath));

                    MessageBox.Show($"Relatório gerado com sucesso: {outputPath}");
                }

                // Generate TXT report as well (in background)
                try { await GenerateTxtReportAsync(projectID); } catch { }

                GrReportPage.Visibility = Visibility.Hidden;
                BtConfirmInfo.IsEnabled = true;
                BtGenerateReport.IsEnabled = false;
                DgPreview.IsReadOnly = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"O seguinte erro ocorreu ao gerar Excel: {ex.Message}\nPath: {templatePath}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.SetCursor(null);
            }
        }

        private async Task GenerateWordReportAsync(string projectID)
        {
            // Generate CSV with structured sections instead of Word
            await GenerateCsvReportAsync(projectID);
        }

        private async Task GenerateCsvReportAsync(string projectID)
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                // Use TSV to avoid issues with semicolons inside field values when opened by different locales
                string path = Path.Combine(desktop, projectID + ".tsv");

                string Delim = "\t";

                string Sanitize(string s)
                {
                    if (s == null) return string.Empty;
                    // remove line breaks and tabs which would break TSV structure; trim whitespace
                    return s.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
                }

                var sb = new StringBuilder();

                string mode = GetActiveSearchMode();

                // VEHICLE_SEARCH output
                if (mode == "VEHICLE_SEARCH")
                {
                    sb.AppendLine("VEHICLE_SEARCH");

                    // Required Info table
                    sb.AppendLine("Required Info");
                    var reqCols = new[] { "1.1.País_de_Circulação", "1.2.Business_Unit", "1.3.Nome_do_produto", "1.4.Market_Share_estipulado_pelo_usuário" };
                    sb.AppendLine(string.Join(Delim, reqCols.Select(c => Sanitize(c))));
                    var reqRow = new[] { SpecialHandleCB(CbRegiao), CbBU?.Text ?? string.Empty, CbProduto?.Text ?? string.Empty, TxMarketShare?.Text ?? string.Empty };
                    sb.AppendLine(string.Join(Delim, reqRow.Select(v => Sanitize(v))));
                    sb.AppendLine();

                    // Inputs table — one row per brand with aligned model/engine/fuel/etc.
                    sb.AppendLine("Inputs");
                    var inCols = new[] { "1.5.Marca_do_Veículo", "1.6.Modelo_do_Veículo", "1.7.Ano_de_Fabricação", "1.8.Motorização", "1.9.Carroceria", "1.10.Tipo_de_Combustível" };
                    sb.AppendLine(string.Join(Delim, inCols.Select(c => Sanitize(c))));

                    var selectedBrandsList = GetSelectedValuesFromCheckBoxComboBox(CbMarca).ToList();
                    var anoFrom = CbAnoDe?.Text ?? string.Empty;
                    var anoTo = CbAnoAte?.Text ?? string.Empty;
                    string anoPeriod = anoFrom + (string.IsNullOrEmpty(anoTo) ? string.Empty : " -> " + anoTo);

                    if (selectedBrandsList.Any() && lastFilteredFleetList != null && lastFilteredFleetList.Count > 0)
                    {
                        // Emit one row per brand with its own models, engines, fuel types, etc.
                        foreach (var brand in selectedBrandsList)
                        {
                            if (string.IsNullOrEmpty(brand)) continue;

                            var rowsForBrand = lastFilteredFleetList
                                .Where(r => string.Equals(r.Brand, brand, StringComparison.OrdinalIgnoreCase))
                                .ToList();

                            if (rowsForBrand.Count == 0) continue;

                            var modelsForBrand = rowsForBrand
                                .Select(r => r.VehicleType)
                                .Where(m => !string.IsNullOrEmpty(m))
                                .Distinct()
                                .OrderBy(m => m)
                                .ToList();

                            var enginesForBrand = rowsForBrand
                                .Select(r => r.EngineInfo)
                                .Where(e => !string.IsNullOrEmpty(e))
                                .Distinct()
                                .OrderBy(e => e)
                                .ToList();

                            var fuelTypesForBrand = rowsForBrand
                                .Select(r => r.FuelType)
                                .Where(f => !string.IsNullOrEmpty(f))
                                .Distinct()
                                .OrderBy(f => f)
                                .ToList();

                            var vClassesForBrand = rowsForBrand
                                .Select(r => r.VClass)
                                .Where(v => !string.IsNullOrEmpty(v))
                                .Distinct()
                                .OrderBy(v => v)
                                .ToList();

                            string brandModels = string.Join(" | ", modelsForBrand);
                            string brandEngines = string.Join(" | ", enginesForBrand);
                            string brandFuels = string.Join(" | ", fuelTypesForBrand);
                            string brandVClasses = string.Join(" | ", vClassesForBrand);

                            var row = new[] { brand, brandModels, anoPeriod, brandEngines, brandVClasses, brandFuels };
                            sb.AppendLine(string.Join(Delim, row.Select(v => Sanitize(v))));
                        }
                    }
                    else
                    {
                        // Fallback: single row with all values joined
                        var brands = JoinList(selectedBrandsList);
                        var selectedModelsList = GetSelectedValuesFromCheckBoxComboBox(CbModelo).ToList();
                        var modelos = JoinList(selectedModelsList);
                        var carroceria = SpecialHandleCB(CbCategoriaVeiculo);
                        var combustivel = JoinList(GetSelectedValuesFromCheckBoxComboBox(CbTipoCombustivel));
                        var motores = lastFilteredFleetList != null
                            ? string.Join(" | ", lastFilteredFleetList.Select(r => r.EngineInfo).Where(e => !string.IsNullOrEmpty(e)).Distinct())
                            : string.Empty;
                        var inputRow = new[] { brands, modelos, anoPeriod, motores, carroceria, combustivel };
                        sb.AppendLine(string.Join(Delim, inputRow.Select(v => Sanitize(v))));
                    }
                    sb.AppendLine();

                    // Competitor table
                    sb.AppendLine("Competitor");
                    var compCols = new[] { "1.11.Nome_competidor", "1.12.Código_competidor" };
                    sb.AppendLine(string.Join(Delim, compCols.Select(c => Sanitize(c))));

                    var (cName, cCode) = GetActiveCompetitorSelection(mode);

                    var compRow = new[] { cName, cCode };
                    sb.AppendLine(string.Join(Delim, compRow.Select(v => Sanitize(v))));
                    sb.AppendLine();
                }

                // OE_SEARCH output
                if (mode == "OE_SEARCH")
                {
                    sb.AppendLine("OE_SEARCH");

                    // Required Info
                    sb.AppendLine("Required Info");
                    var reqCols = new[] { "1.1.País_de_Circulação", "1.2.Business_Unit", "1.3.Nome_do_produto", "1.4.Market_Share_estipulado_pelo_usuário" };
                    sb.AppendLine(string.Join(Delim, reqCols.Select(c => Sanitize(c))));
                    var reqRow = new[] { SpecialHandleCB(CbRegiao), CbBU?.Text ?? string.Empty, CbProduto?.Text ?? string.Empty, TxMarketShare?.Text ?? string.Empty };
                    sb.AppendLine(string.Join(Delim, reqRow.Select(v => Sanitize(v))));
                    sb.AppendLine();

                    // Competitor
                    sb.AppendLine("Competitor");
                    var compCols = new[] { "1.11.Nome_competidor", "1.12.Código_competidor" };
                    sb.AppendLine(string.Join(Delim, compCols.Select(c => Sanitize(c))));

                    var cName = CbCompetitor_OE?.Text ?? string.Empty;
                    var cCode = GetCompetitorCodeForPrompt(CbCompetitorCode_OE);
                    sb.AppendLine(string.Join(Delim, new[] { cName, cCode }.Select(v => Sanitize(v))));
                    sb.AppendLine();

                    // Código OE
                    sb.AppendLine("Código OE");
                    var oeCols = new[] { "1.13.Códigos_OE" };
                    sb.AppendLine(string.Join(Delim, oeCols.Select(c => Sanitize(c))));
                    var oeInput = TxSearchOENumber_OE?.Text ?? string.Empty;
                    sb.AppendLine(Sanitize(oeInput));
                    sb.AppendLine();

                    // SHORTs found for OE
                    sb.AppendLine("SHORT");
                    var shortCols = new[] { "1.13.1.SHORT_encontrados" };
                    sb.AppendLine(string.Join(Delim, shortCols.Select(c => Sanitize(c))));

                    var shortsFromOe = new List<string>();
                    var oesList = oeInput.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Where(oe => !string.IsNullOrWhiteSpace(oe)).ToList();
                    if (oesList.Count > 0)
                    {
                        using (OleDbConnection con = new OleDbConnection(connectionString))
                        {
                            await con.OpenAsync();
                            string oeWhere = string.Join(",", oesList.Select(oe => "'" + oe.Replace("'", "''") + "'"));
                            string query = $"SELECT DISTINCT SHORT FROM base_OE WHERE OE_number IN ({oeWhere})";
                            using (OleDbCommand com = new OleDbCommand(query, con))
                            using (OleDbDataReader reader = (OleDbDataReader)await com.ExecuteReaderAsync())
                            {
                                while (await reader.ReadAsync())
                                {
                                    if (!reader.IsDBNull(0))
                                        shortsFromOe.Add(reader.GetString(0));
                                }
                            }
                        }
                    }

                    sb.AppendLine(Sanitize(string.Join(";", shortsFromOe.Distinct(StringComparer.OrdinalIgnoreCase))));
                    sb.AppendLine();
                }

                // BK_SEARCH output
                if (mode == "BK_SEARCH")
                {
                    sb.AppendLine("BK_SEARCH");

                    // Required Info
                    sb.AppendLine("Required Info");
                    var reqCols = new[] { "1.1.País_de_Circulação", "1.2.Business_Unit", "1.3.Nome_do_produto", "1.4.Market_Share_estipulado_pelo usuário" };
                    sb.AppendLine(string.Join(Delim, reqCols.Select(c => Sanitize(c))));
                    var reqRow = new[] { SpecialHandleCB(CbRegiao), CbBU?.Text ?? string.Empty, CbProduto?.Text ?? string.Empty, TxMarketShare?.Text ?? string.Empty };
                    sb.AppendLine(string.Join(Delim, reqRow.Select(v => Sanitize(v))));
                    sb.AppendLine();

                    // Competitor
                    sb.AppendLine("Competitor");
                    var compCols = new[] { "1.11.Nome_competidor", "1.12.Código_competidor" };
                    sb.AppendLine(string.Join(Delim, compCols.Select(c => Sanitize(c))));

                    var cName = CbCompetitor_BK?.Text ?? string.Empty;
                    var cCode = GetCompetitorCodeForPrompt(CbCompetitorCode_BK);
                    sb.AppendLine(string.Join(Delim, new[] { cName, cCode }.Select(v => Sanitize(v))));
                    sb.AppendLine();

                    // Código BK
                    sb.AppendLine("Código BK");
                    var bkCols = new[] { "1.14.Bosch_Key (SHORT)" };
                    sb.AppendLine(string.Join(Delim, bkCols.Select(c => Sanitize(c))));
                    sb.AppendLine(Sanitize(TxSearchBK_BK?.Text ?? string.Empty));
                    sb.AppendLine();
                }

                // Append analysis context/instructions at end of TSV
                sb.AppendLine();

                // Merge with prompt template file (if available) to produce a single text
                string filePromptText = string.IsNullOrWhiteSpace(viabilityPrompt)
                    ? LoadViabilityPromptTemplate()
                    : viabilityPrompt;

                // Compose final text as: generated sb content + filePromptText (if any)
                string promptPart = (filePromptText ?? string.Empty).TrimEnd();
                string finalText = string.Empty;

                if (!string.IsNullOrEmpty(promptPart))
                    finalText = sb.ToString() + Environment.NewLine + promptPart;
                else
                    finalText = sb.ToString();

                // Save merged content
                await Task.Run(() => File.WriteAllText(path, finalText, new UTF8Encoding(true)));

                // Open agent chat and copy message for user to paste
                OpenM365AgentAndCopyInstruction(path);

                // Popup removido (não será mais utilizado)
                // MessageBox.Show(
                //     $"Relatório TSV gerado: {path}\n\nO link do agente foi aberto e as instruções foram copiadas para a área de transferência (Ctrl+V no chat).",
                //     "Relatório",
                //     MessageBoxButton.OK,
                //     MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gerar CSV: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void OpenM365AgentAndCopyInstruction(string tsvPath)
        {
            try
            {
                string tsvText = string.Empty;
                try
                {
                    tsvText = File.ReadAllText(tsvPath, Encoding.UTF8);
                }
                catch
                {
                    tsvText = $"(Nao foi possivel ler o TSV automaticamente. Caminho: {tsvPath})";
                }

                // Convert TSV to JSON for better agent processing
                string jsonMessage = ConvertTsvToJson(tsvText);

                // Abrir dentro do app (WebView2) e tentar auto-login/auto-send
                try
                {
                    var w = new M365ChatWindow(M365AgentUrl, jsonMessage);
                    w.Owner = this;
                    w.Show();
                }
                catch
                {
                    // fallback: open default browser
                    try
                    {
                        Process.Start(new ProcessStartInfo { FileName = M365AgentUrl, UseShellExecute = true });
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Falha ao abrir/enviar no Copilot pelo Playwright: {ex.Message}\n\nO texto foi copiado para a area de transferencia (Ctrl+V).",
                    "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string ConvertTsvToJson(string tsvText)
        {
            try
            {
                var lines = tsvText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                var jsonBuilder = new StringBuilder();
                jsonBuilder.AppendLine("{");

                int i = 0;
                string searchMode = "";
                var requiredInfo = new Dictionary<string, string>();
                var vehicleInputs = new List<Dictionary<string, string>>();
                var competitor = new Dictionary<string, string>();
                string oeNumbers = "";
                string bkCodes = "";
                string instructions = "";

                bool IsInstructionStart(string raw)
                {
                    if (string.IsNullOrWhiteSpace(raw)) return false;

                    string normalized = raw.TrimStart('\uFEFF').Trim();
                    normalized = normalized.TrimStart('"').Trim();

                    return normalized.StartsWith("# CONTEXTO", StringComparison.OrdinalIgnoreCase)
                           || normalized.StartsWith("# DADOS A CONSULTAR", StringComparison.OrdinalIgnoreCase)
                           || normalized.StartsWith("# ESTRUTURA OBRIGATÓRIA", StringComparison.OrdinalIgnoreCase)
                           || normalized.StartsWith("### REGRAS OBRIGATÓRIAS", StringComparison.OrdinalIgnoreCase);
                }

                // Parse sections
                while (i < lines.Length)
                {
                    string line = lines[i].Trim();

                    // Detect search mode
                    if (line == "VEHICLE_SEARCH" || line == "OE_SEARCH" || line == "BK_SEARCH")
                    {
                        searchMode = line;
                        i++;
                        continue;
                    }

                    // Parse Required Info
                    if (line == "Required Info")
                    {
                        i++;
                        if (i < lines.Length)
                        {
                            var headers = lines[i].Split('\t');
                            i++;
                            if (i < lines.Length)
                            {
                                var values = lines[i].Split('\t');
                                for (int j = 0; j < Math.Min(headers.Length, values.Length); j++)
                                {
                                    requiredInfo[headers[j]] = values[j];
                                }
                            }
                        }
                        i++;
                        continue;
                    }

                    // Parse Inputs (vehicle data)
                    if (line == "Inputs")
                    {
                        i++;
                        if (i < lines.Length)
                        {
                            var headers = lines[i].Split('\t');
                            i++;
                            while (i < lines.Length && !string.IsNullOrEmpty(lines[i].Trim()) && 
                                   !lines[i].StartsWith("Competitor") && !lines[i].StartsWith("#"))
                            {
                                var values = lines[i].Split('\t');
                                var vehicleData = new Dictionary<string, string>();
                                for (int j = 0; j < Math.Min(headers.Length, values.Length); j++)
                                {
                                    vehicleData[headers[j]] = values[j];
                                }
                                if (vehicleData.Count > 0)
                                    vehicleInputs.Add(vehicleData);
                                i++;
                            }
                        }
                        continue;
                    }

                    // Parse Competitor
                    if (line == "Competitor")
                    {
                        i++;
                        if (i < lines.Length)
                        {
                            var headers = lines[i].Split('\t');
                            i++;
                            if (i < lines.Length)
                            {
                                var values = lines[i].Split('\t');
                                for (int j = 0; j < Math.Min(headers.Length, values.Length); j++)
                                {
                                    competitor[headers[j]] = values[j];
                                }
                            }
                        }
                        i++;
                        continue;
                    }

                    // Parse OE Numbers
                    if (line == "OE Number")
                    {
                        i++;
                        if (i < lines.Length)
                        {
                            i++; // skip header
                            if (i < lines.Length)
                            {
                                oeNumbers = lines[i];
                            }
                        }
                        i++;
                        continue;
                    }

                    // Parse BK Codes
                    if (line == "Código BK")
                    {
                        i++;
                        if (i < lines.Length)
                        {
                            i++; // skip header
                            if (i < lines.Length)
                            {
                                bkCodes = lines[i];
                            }
                        }
                        i++;
                        continue;
                    }

                    // Capture instructions (everything after markdown prompt headers)
                    if (IsInstructionStart(lines[i]))
                    {
                        instructions = string.Join("\n", lines.Skip(i));
                        break;
                    }

                    i++;
                }

                // Fallback: if markdown header was not found, keep everything after the competitor section.
                if (string.IsNullOrWhiteSpace(instructions))
                {
                    int competitorIndex = Array.FindIndex(lines, l => string.Equals(l.Trim(), "Competitor", StringComparison.OrdinalIgnoreCase));
                    if (competitorIndex >= 0)
                    {
                        // Competitor block has at most: title, headers, values, blank line
                        int start = Math.Min(competitorIndex + 4, lines.Length);
                        var tail = lines.Skip(start).ToArray();
                        if (tail.Any(l => !string.IsNullOrWhiteSpace(l)))
                            instructions = string.Join("\n", tail).Trim();
                    }
                }

                if (string.IsNullOrWhiteSpace(instructions) && !string.IsNullOrWhiteSpace(viabilityPrompt))
                {
                    instructions = viabilityPrompt;
                }

                // Build JSON structure
                jsonBuilder.AppendLine($"  \"searchMode\": \"{EscapeJson(searchMode)}\",");
                
                // Required Info
                jsonBuilder.AppendLine("  \"requiredInfo\": {");
                var reqInfoItems = requiredInfo.Select(kvp => $"    \"{EscapeJson(kvp.Key)}\": \"{EscapeJson(kvp.Value)}\"");
                jsonBuilder.AppendLine(string.Join(",\n", reqInfoItems));
                jsonBuilder.AppendLine("  },");

                // Vehicle Inputs (or OE/BK data)
                if (searchMode == "VEHICLE_SEARCH" && vehicleInputs.Count > 0)
                {
                    jsonBuilder.AppendLine("  \"vehicles\": [");
                    for (int v = 0; v < vehicleInputs.Count; v++)
                    {
                        jsonBuilder.AppendLine("    {");
                        var vehicleItems = vehicleInputs[v].Select(kvp => $"      \"{EscapeJson(kvp.Key)}\": \"{EscapeJson(kvp.Value)}\"");
                        jsonBuilder.AppendLine(string.Join(",\n", vehicleItems));
                        jsonBuilder.Append("    }");
                        if (v < vehicleInputs.Count - 1) jsonBuilder.AppendLine(",");
                        else jsonBuilder.AppendLine();
                    }
                    jsonBuilder.AppendLine("  ],");
                }
                else if (searchMode == "OE_SEARCH")
                {
                    jsonBuilder.AppendLine($"  \"oeNumbers\": \"{EscapeJson(oeNumbers)}\",");
                }
                else if (searchMode == "BK_SEARCH")
                {
                    jsonBuilder.AppendLine($"  \"boschKeys\": \"{EscapeJson(bkCodes)}\",");
                }

                // Competitor
                jsonBuilder.AppendLine("  \"competitor\": {");
                var compItems = competitor.Select(kvp => $"    \"{EscapeJson(kvp.Key)}\": \"{EscapeJson(kvp.Value)}\"");
                jsonBuilder.AppendLine(string.Join(",\n", compItems));
                jsonBuilder.AppendLine("  },");

                // Instructions
                jsonBuilder.AppendLine($"  \"analysisInstructions\": \"{EscapeJson(instructions)}\"");

                jsonBuilder.AppendLine("}");

                return jsonBuilder.ToString();
            }
            catch (Exception ex)
            {
                // If conversion fails, return original TSV with error note
                return $"{{\"error\": \"Failed to convert TSV to JSON: {EscapeJson(ex.Message)}\", \"originalTsv\": \"{EscapeJson(tsvText)}\"}}";
            }
        }

        private string EscapeJson(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t");
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            DisposeOfficeApps();
        }

        private void DisposeOfficeApps()
        {
            try
            {
                // No COM objects to release anymore
            }
            catch { }
        }

        private List<string> GetSelectedValuesFromCheckBoxComboBox(ComboBox cb)
        {
            if (cb?.Items == null) return new List<string>();

            var selectedValues = cb.Items.OfType<CheckBox>()
                .Where(i => i.IsChecked == true && !IsSelectAllCheckBox(i))
                .Select(i => i.Content?.ToString() ?? string.Empty)
                .Where(s => s.Length > 0)
                .ToList();

            var typedValues = ParseMultiValues(cb.Text);

            return selectedValues
                .Concat(typedValues)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // Return selected competitor codes from a ComboBox whose items are CheckBoxes
        private List<string> GetSelectedCompetitorCodes(ComboBox codeCombo)
        {
            if (codeCombo == null) return new List<string>();

            var selectedCodes = codeCombo.Items.OfType<CheckBox>()
                .Where(cb => cb.IsChecked == true && !IsSelectAllCheckBox(cb))
                .Select(cb => cb.Content?.ToString() ?? string.Empty)
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            var typedCodes = ParseMultiValues(codeCombo.Text);

            return selectedCodes
                .Concat(typedCodes)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> ParseMultiCodes(string raw)
        {
            return ParseMultiValues(raw);
        }

        private List<string> ParseMultiValues(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new List<string>();

            return raw
                .Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(s => !string.IsNullOrWhiteSpace(s) && !string.Equals(s.Trim(), SelectAllOption, StringComparison.OrdinalIgnoreCase))
                .Select(s => s.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void UpdateCodeComboText(ComboBox codeCombo)
        {
            try
            {
                var selectedCodes = GetSelectedCompetitorCodes(codeCombo);
                Dispatcher.Invoke(() => {
                    codeCombo.Text = string.Join("; ", selectedCodes);
                });
            }
            catch { }
        }

        private bool IsSelectAllCheckBox(CheckBox cb)
        {
            return string.Equals(cb?.Content?.ToString(), SelectAllOption, StringComparison.OrdinalIgnoreCase);
        }

        private List<CheckBox> CreateCheckBoxItems(IEnumerable<string> values, ComboBox ownerCombo, bool attachRegionRestriction = false, bool includeSelectAll = true)
        {
            var list = new List<CheckBox>();

            if (includeSelectAll)
            {
                var selectAll = new CheckBox() { Content = SelectAllOption };
                selectAll.Tag = ownerCombo;
                selectAll.Checked += CheckBoxItem_Checked;
                selectAll.Unchecked += CheckBoxItem_Unchecked;
                list.Add(selectAll);
            }

            foreach (var value in values.Where(v => !string.IsNullOrWhiteSpace(v)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v => v))
            {
                var cb = new CheckBox() { Content = value };
                cb.Tag = ownerCombo;
                cb.Checked += CheckBoxItem_Checked;
                cb.Unchecked += CheckBoxItem_Unchecked;
                if (attachRegionRestriction)
                {
                    cb.Checked += RegionCheckBox_Checked;
                }
                list.Add(cb);
            }

            return list;
        }

        private void CheckBoxItem_Checked(object sender, RoutedEventArgs e)
        {
            HandleCheckBoxStateChange(sender as CheckBox);
        }

        private void CheckBoxItem_Unchecked(object sender, RoutedEventArgs e)
        {
            HandleCheckBoxStateChange(sender as CheckBox);
        }

        private void HandleCheckBoxStateChange(CheckBox? changed)
        {
            if (changed == null || isBulkChecking)
                return;

            var owner = FindOwningComboBox(changed);
            if (owner == null)
                return;

            try
            {
                isBulkChecking = true;

                if (IsSelectAllCheckBox(changed))
                {
                    bool shouldCheck = changed.IsChecked == true;
                    bool suppressRestrictionBackup = suppressRegionRestrictionMessage;
                    if (owner == CbRegiao)
                    {
                        suppressRegionRestrictionMessage = true;
                    }

                    foreach (var cb in owner.Items.OfType<CheckBox>().Where(i => !IsSelectAllCheckBox(i)))
                    {
                        if (owner == CbRegiao && shouldCheck)
                        {
                            string region = cb.Content?.ToString()?.Trim() ?? string.Empty;
                            cb.IsChecked = string.Equals(region, AllowedRegion, StringComparison.OrdinalIgnoreCase);
                        }
                        else
                        {
                            cb.IsChecked = shouldCheck;
                        }
                    }

                    if (owner == CbRegiao)
                    {
                        suppressRegionRestrictionMessage = suppressRestrictionBackup;
                    }
                }
                else
                {
                    var normalItems = owner.Items.OfType<CheckBox>().Where(i => !IsSelectAllCheckBox(i)).ToList();
                    var selectAll = owner.Items.OfType<CheckBox>().FirstOrDefault(IsSelectAllCheckBox);
                    if (selectAll != null)
                    {
                        selectAll.IsChecked = normalItems.Count > 0 && normalItems.All(i => i.IsChecked == true);
                    }
                }
            }
            finally
            {
                isBulkChecking = false;
            }

            owner.Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox(owner));
            EnableNextComboBox(owner);

            if (!isFilteringVehicleInputs && ShouldTriggerRealtimeCascade(owner))
            {
                FilterAllFields();
            }
        }

        private bool ShouldTriggerRealtimeCascade(ComboBox owner)
        {
            return owner == CbMarca
                || owner == CbModelo
                || owner == CbTipoCombustivel
                || owner == CbExplanation
                || owner == CbCategoriaVeiculo;
        }

        private ComboBox? FindOwningComboBox(CheckBox? checkbox)
        {
            if (checkbox == null)
                return null;

            if (checkbox.Tag is ComboBox taggedCombo)
                return taggedCombo;

            var parent = System.Windows.Media.VisualTreeHelper.GetParent(checkbox);
            while (parent != null)
            {
                if (parent is ComboBox combo)
                    return combo;

                parent = System.Windows.Media.VisualTreeHelper.GetParent(parent);
            }

            return null;
        }

        private string GetCompetitorCodeForPrompt(ComboBox codeCombo)
        {
            var selectedCodes = GetSelectedCompetitorCodes(codeCombo);
            if (selectedCodes.Count > 0)
                return string.Join(", ", selectedCodes);

            return codeCombo?.Text?.Trim() ?? string.Empty;
        }

        // Join a list of strings into a readable single string (used for fallback TSV rows)
        private string JoinList(List<string> list)
        {
            if (list == null || list.Count == 0) return string.Empty;
            return string.Join(" ; ", list);
        }

        private void AddInCondition(List<string> conditions, string columnName, List<string> selectedValues)
        {
            if (selectedValues == null || selectedValues.Count == 0) return;
            string formattedValues = string.Join(",", selectedValues.Select(v => "'" + v.Replace("'", "''") + "'"));
            conditions.Add($"[{columnName}] IN ({formattedValues})");
        }

        private void AddEqualsCondition(List<string> conditions, string columnName, string selectedValue)
        {
            if (string.IsNullOrEmpty(selectedValue)) return;
            conditions.Add($"[{columnName}] = '{selectedValue.Replace("'", "''")}'");
        }

        private string SpecialHandleCB(ComboBox cb)
        {
            if (cb?.Items == null) return string.Empty;
            return string.Join(", ", GetSelectedValuesFromCheckBoxComboBox(cb));
        }

        private void ResetAllInputs()
        {
            try
            {
                foreach (var cb in CbMarca.Items.OfType<CheckBox>()) cb.IsChecked = false;
                foreach (var cb in CbRegiao.Items.OfType<CheckBox>()) cb.IsChecked = false;
                foreach (var cb in CbModelo.Items.OfType<CheckBox>()) cb.IsChecked = false;
                foreach (var cb in CbTipoCombustivel.Items.OfType<CheckBox>()) cb.IsChecked = false;
                foreach (var cb in CbExplanation.Items.OfType<CheckBox>()) cb.IsChecked = false;
                foreach (var cb in CbCategoriaVeiculo.Items.OfType<CheckBox>()) cb.IsChecked = false;
            }
            catch { }
            CbBU.Text = CbProduto.Text = CbRegiao.Text = CbMarca.Text = CbModelo.Text = string.Empty;
            CbAnoAte.Text = CbAnoDe.Text = CbTipoCombustivel.Text = CbExplanation.Text = CbCategoriaVeiculo.Text = string.Empty;
            TxSearchBK_BK.Text = TxSearchOENumber_OE.Text = TxMarketShare.Text = string.Empty;

            // Clear competitor fields from all tabs (VEHICLE, OE, BK)
            try { CbCompetitor.Text = string.Empty; } catch { }
            try { CbCompetitorCode.Text = string.Empty; } catch { }
            try { CbCompetitor_OE.Text = string.Empty; } catch { }
            try { CbCompetitorCode_OE.Text = string.Empty; } catch { }
            try { CbCompetitor_BK.Text = string.Empty; } catch { }
            try { CbCompetitorCode_BK.Text = string.Empty; } catch { }

            // Re-disable cascading ComboBoxes
            DisableComboBoxes();
            CbBU.IsEnabled = true;

            previewItems.Clear();
        }

        private async Task GenerateTxtReportAsync(string projectID)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine($"Projeto: {projectID}");
                sb.AppendLine($"Gerado em: {DateTime.Now}");
                sb.AppendLine();
                sb.AppendLine("BK\tVehicleName\tPotential");
                foreach (var p in previewItems)
                {
                    sb.AppendLine($"{p.BK}\t{p.VehicleName}\t{p.Potential}");
                }
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string path = Path.Combine(desktop, projectID + ".txt");
                await Task.Run(() => File.WriteAllText(path, sb.ToString(), Encoding.UTF8));
            }
            catch { }
        }

        // Simple in-memory query cache singleton used by UI
        public class QueryCache
        {
            private static readonly Lazy<QueryCache> instance = new Lazy<QueryCache>(() => new QueryCache());
            public static QueryCache Instance => instance.Value;
            private readonly Dictionary<string, Dictionary<string, object>> cache = new();
            private readonly Dictionary<string, object> staticData = new();
            private QueryCache() { }
            public bool HasKey(string k) => cache.ContainsKey(k);
            public Dictionary<string, object> Get(string k) => cache[k];
            public void Set(string k, Dictionary<string, object> v) => cache[k] = v;
            public void SetStaticData(string k, object v) => staticData[k] = v;
            public object GetStaticData(string k) => staticData[k];
        }

        // Types for data rows
        public class FleetDataRow
        {
            public string Short { get; set; }
            public string Country { get; set; }
            public double FasPopulation { get; set; }
            public string Brand { get; set; }
            public string VehicleType { get; set; }
            public string Data { get; set; }
            public string EngineInfo { get; set; }
            public string FuelType { get; set; }
            public string Explanation { get; set; }
            public string VClass { get; set; }

            // Parsed year properties
            public string AnoInicio { get; set; }
            public string AnoFim { get; set; }
        }

        public class EposDataItem
        {
            public string BU { get; set; }
            public string Produto { get; set; }
            public string Category { get; set; }
            public string Country { get; set; }
            public double ExchangeRateYear { get; set; }
            public double UnitReplacementShare { get; set; }
            public double AverageReplacementQuantity { get; set; }
            public double MS { get; set; }
        }

        public class PreviewItem
        {
            public string BK { get; set; }
            public string OENumber { get; set; }
            public string Manufacturer { get; set; }
            public string VehicleName { get; set; }
            public string ApplicationPeriod { get; set; }
            public string FuelType { get; set; }
            public string EngineType { get; set; }
            public string Explanation { get; set; }
            public string Fleet { get; set; }
            public string Region { get; set; }
            public string Potential { get; set; }
            public string UserDefinedMarketShare { get; set; }
            public string CompetitorName { get; set; }
            public string CompetitorCode { get; set; }
        }

        private async void Window_ContentRendered(object sender, EventArgs e)
        {
            try
            {
                GrLoadingCover.Visibility = Visibility.Visible;
                PbLoadingBases.IsIndeterminate = false;
                PbLoadingBases.Value = 0;
                LbLoadingStatus.Content = "Starting database load...";
            }
            catch { }

            Mouse.SetCursor(Cursors.Wait);
            await FirstLoadInfos();
            Mouse.SetCursor(null);
            try { GrLoadingCover.Visibility = Visibility.Hidden; } catch { }
        }

        private void CbRegiao_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                ((ComboBox)sender).Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox((ComboBox)sender));
                EnableNextComboBox(CbRegiao);
            }
            catch { }
        }

        private void CbMarca_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                ((ComboBox)sender).Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox((ComboBox)sender));
                EnableNextComboBox(CbMarca);
                FilterAllFields();
            }
            catch { }
        }

        private void CbModelo_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                ((ComboBox)sender).Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox((ComboBox)sender));
                EnableNextComboBox(CbModelo);
                FilterAllFields();
            }
            catch { }
        }

        private void CbTipoCombustivel_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                ((ComboBox)sender).Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox((ComboBox)sender));
                EnableNextComboBox(CbTipoCombustivel);
                FilterAllFields();
            }
            catch { }
        }

        private void CbExplanation_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                ((ComboBox)sender).Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox((ComboBox)sender));
                EnableNextComboBox(CbExplanation);
                FilterAllFields();
            }
            catch { }
        }

        private void CbCategoriaVeiculo_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                ((ComboBox)sender).Text = string.Join("; ", GetSelectedValuesFromCheckBoxComboBox((ComboBox)sender));
                EnableNextComboBox(CbCategoriaVeiculo);
            }
            catch { }
        }

        private void BtReset_Click(object sender, RoutedEventArgs e)
        {
            ResetAllInputs();
        }

        private void BtGerarRelatorio_Click(object sender, RoutedEventArgs e)
        {
            Mouse.SetCursor(Cursors.Wait);
            string? projectID = Microsoft.VisualBasic.Interaction.InputBox("Digite o nome do projeto:", "Nome do Projeto", "Project_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            if (!string.IsNullOrWhiteSpace(projectID))
            {
                GenerateCsvReportAsync(projectID).ConfigureAwait(false);
            }
            Mouse.SetCursor(null);
        }

        private void PopulateCompetitorCombos()
        {
            try
            {
                // Use competitorTable preloaded in PreLoadAllData (cod_concorrente)
                var brands = competitorTable.Select(c => c.Markenbez).Where(b => !string.IsNullOrEmpty(b)).Distinct().OrderBy(b => b).ToList();

                // Populate on UI thread
                Dispatcher.Invoke(() => {
                    if (CbCompetitor != null) CbCompetitor.ItemsSource = brands;
                    if (CbCompetitor_OE != null) CbCompetitor_OE.ItemsSource = brands;
                    if (CbCompetitor_BK != null) CbCompetitor_BK.ItemsSource = brands;
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao popular concorrentes: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CbCompetitor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string brand = CbCompetitor?.Text ?? string.Empty;
                PopulateCompetitorCodesForCombo(brand, CbCompetitorCode);
            }
            catch { }
        }

        private void CbCompetitor_OE_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string brand = CbCompetitor_OE?.Text ?? string.Empty;
                PopulateCompetitorCodesForCombo(brand, CbCompetitorCode_OE);
            }
            catch { }
        }

        private void CbCompetitor_BK_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string brand = CbCompetitor_BK?.Text ?? string.Empty;
                PopulateCompetitorCodesForCombo(brand, CbCompetitorCode_BK);
            }
            catch { }
        }

        private void CbCompetitor_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                string brand = CbCompetitor?.Text ?? string.Empty;
                PopulateCompetitorCodesForCombo(brand, CbCompetitorCode);
            }
            catch { }
        }

        private void CbCompetitor_OE_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                string brand = CbCompetitor_OE?.Text ?? string.Empty;
                PopulateCompetitorCodesForCombo(brand, CbCompetitorCode_OE);
            }
            catch { }
        }

        private void CbCompetitor_BK_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                string brand = CbCompetitor_BK?.Text ?? string.Empty;
                PopulateCompetitorCodesForCombo(brand, CbCompetitorCode_BK);
            }
            catch { }
        }

        private void PopulateCompetitorCodesForCombo(string brand, ComboBox codeCombo)
        {
            try
            {
                string brandFilter = (brand ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(brandFilter))
                {
                    Dispatcher.Invoke(() => codeCombo.ItemsSource = null);
                    return;
                }

                var exactCodes = competitorTable
                    .Where(c => string.Equals(c.Markenbez, brandFilter, StringComparison.OrdinalIgnoreCase))
                    .Select(c => c.TwnrVerd)
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Select(c => c!.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(c => c)
                    .ToList();

                var codes = exactCodes;

                // Fallback for typed filtering: allow partial match on MARKENBEZ.
                if (codes.Count == 0)
                {
                    codes = competitorTable
                        .Where(c => !string.IsNullOrWhiteSpace(c.Markenbez) && c.Markenbez.Contains(brandFilter, StringComparison.OrdinalIgnoreCase))
                        .Select(c => c.TwnrVerd)
                        .Where(c => !string.IsNullOrWhiteSpace(c))
                        .Select(c => c!.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(c => c)
                        .ToList();
                }

                var checkboxes = CreateCheckBoxItems(codes, codeCombo);

                Dispatcher.Invoke(() => codeCombo.ItemsSource = checkboxes);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao popular códigos de concorrente: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}