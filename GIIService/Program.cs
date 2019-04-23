using Microsoft.Azure.ServiceBus;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Sandboxable.Microsoft.WindowsAzure.Storage;
using Sandboxable.Microsoft.WindowsAzure.Storage.File;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace GIIService
{
    class Program
    {
        static IQueueClient queueClient;
        static bool processing = false;

        static void Main(string[] args)
        {
            Console.WriteLine("Started Process..");
            queueClient = new QueueClient(ConfigurationManager.AppSettings.Get("ServiceBusConnectionString"), ConfigurationManager.AppSettings.Get("QueueName"));

            // Register QueueClient's MessageHandler and receive messages in a loop
            RegisterOnMessageHandlerAndReceiveMessages();

            //Keep the app running to recieve service bus messages
            while (true)
                Thread.Sleep(3000);
        }

        static void RegisterOnMessageHandlerAndReceiveMessages()
        {
            var messageHandlerOptions = new MessageHandlerOptions(ExceptionReceivedHandler)
            {
                MaxConcurrentCalls = 1,
                AutoComplete = true
            };

            queueClient.RegisterMessageHandler(ProcessMessagesAsync, messageHandlerOptions);
        }

        private static async Task ProcessMessagesAsync(Message message, CancellationToken token)
        {
            Console.WriteLine($"Received message: SequenceNumber:{message.SystemProperties.SequenceNumber} Body:{Encoding.UTF8.GetString(message.Body)}");

            dynamic jsonObject = JsonConvert.DeserializeObject(Encoding.UTF8.GetString(message.Body));
            string year = jsonObject.year;
            string reportType = jsonObject.reportType;

            //Ignore concurrent messages recieved from the queue when reports are being generated
            if (!processing)
            {
                var thread = new Thread(new ParameterizedThreadStart(param =>
                {
                    processing = true;
                    KillAllExcelFileProcesses();
                    ClearImages();
                    GenerateReport(year, reportType);
                    UploadCountryBriefReports(year, reportType);
                    KillAllExcelFileProcesses();
                    Console.WriteLine("Process Completed!");
                    processing = false;

                }));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            }

        }

        static Task ExceptionReceivedHandler(ExceptionReceivedEventArgs exceptionReceivedEventArgs)
        {
            return Task.CompletedTask;
        }
        static void GenerateReport(string year, string reportType)
        {

            //Ranking data excel file details
            string wbkRanking = ConfigurationManager.AppSettings.Get("wbkRanking");
            Console.WriteLine("Processing: " + wbkRanking);
            string incomeGroupRegionRankingWorkSheet = "Table 1";
            //string rankingWorkSheetInputIndex = "Table 2";
            //string rankingWorkSheetOutputIndex = "Table 3";
            int incomeGroupRegionCountryColumn = 1;
            int incomeGroupRegionCountryRowStart = 5;
            int incomeGroupRegionCountryRowEnd = 130;
            int incomeGroupCount = 4;
            int regionCount = 7;

            var incomeGroupRegionRankJsonData = @"[{'key': '#ig#', valueType: 'string', 'column' : 5}, {'key': '#r#', valueType: 'string', 'column' : 7}]";
            var superScriptIncomeGroupRegionRankJsonData = @"[{'key': '#ixg#', valueType: 'string', 'column' : 5}, {'key': '#rx#', valueType: 'string', 'column' : 7}]";

            Excel._Application xlAppRanking = new Excel.Application();
            Excel.Workbook workbookRanking = xlAppRanking.Workbooks.Open(wbkRanking);
            Excel.Worksheet worksheetIncomeGroupRegionRanking = workbookRanking.Sheets[incomeGroupRegionRankingWorkSheet];
            //Excel.Worksheet worksheetRankingInputIndex = workbookRanking.Sheets[rankingWorkSheetInputIndex];
            // Excel.Worksheet worksheetRankingOutputIndex = workbookRanking.Sheets[rankingWorkSheetOutputIndex];

            //strength weakness and graph1 (bar chart) data excel file details
            string wbkGII = ConfigurationManager.AppSettings.Get("wbkGII");
            Console.WriteLine("Processing: " + wbkGII);
            string workSheetNameGII2018 = "GII 2018";
            string workSheetNameStrengthWeakness = "App I";
            //Data in App I worksheet
            int countryNameRow = 1;
            int countryNameColumn = 6;
            int StrengthWeaknessColumn = 11;
            int subPillarListColumn = 7;
            int pillarSubpillarIndicatorColumn = 5;
            //int pillarSubpillarRankColumn = 10;
            int[] pillarRowList = new int[7] { 17, 28, 44, 58, 71, 90, 108 };
            int[] subPillarStartRowList = new int[7] { 18, 29, 45, 59, 72, 91, 109 };
            int[] subPillarLastRowList = new int[7] { 27, 43, 57, 70, 89, 107, 124 };

            //Data in GII 2018 worksheet
            //word document text and excel sheet row (score and rank value) one to one map json array 
            var GIIInputOutputRankIGroupRegionJsonData = @"[{'key': '#country#', valueType: 'string', 'column' : 3}, {'key': '#gr#', valueType: 'string', 'column' : 38}, {'key': '#er#', valueType: 'string', 'column' : 39}, {'key': '#ir#', valueType: 'string', 'column' : 40}, {'key': '#or#', valueType: 'string', 'column' : 41}, {'key': '#incomegroup#', valueType: 'string', 'column' : 8}, {'key': '#region#', valueType: 'string', 'column' : 11 }]";
            var superScriptInputOutputRankIGroupRegionJsonData = @"[{'key': '#gxr#', valueType: 'supper', 'column' : 38}, {'key': '#exr#', valueType: 'supper', 'column' : 39}, {'key': '#ixr#', valueType: 'supper', 'column' : 40}, {'key': '#oxr#', valueType: 'supper', 'column' : 41}]";

            int GIICountryListColumn = 3;
            int GIICountryListRowStart = 15;
            int GIICountryListRowEnd = 140;
            //int inputColumn = 17;
            //int outputColumn = 19;

            int incomeGroupColumn = 8;
            int regionColumn = 11;

            //To get data to create graph1(bar chart)
            int pillarRankColumn = 10;
            int giiRankRow = 12;
            int[] graph1PillarRowList = new int[8] { 17, 28, 44, 58, 71, 90, 108, 12 };
            string[] graph1PillarList = new String[8];

            Excel._Application xlAppGII = new Excel.Application();
            Excel.Workbook workbookGII = xlAppGII.Workbooks.Open(wbkGII);
            Excel.Worksheet worksheetGII2018 = workbookGII.Sheets[workSheetNameGII2018];
            Excel.Worksheet worksheetStrengthWeakness = workbookGII.Sheets[workSheetNameStrengthWeakness];


            //region,income group, graph1, graph2 data excel file details
            string wbkIncomeRegion = ConfigurationManager.AppSettings.Get("wbkIncomeRegion");
            Console.WriteLine("Processing: " + wbkIncomeRegion);
            string workSheetNameIncome = "Income group";
            string workSheetNameRegion = "region";
            string workSheetNameGraph1 = "graph 1";
            string workSheetNameGraph2 = "graph 2";
            string workSheetNameTop10 = "Top 10";
            //get pillars row column (region worksheet)
            int pillarListRow = 4;
            int pillarListStartColumn = 2;
            int pillarListLastColumn = 8;
            int pilarCount = pillarListLastColumn - pillarListStartColumn + 1;
            string[] pillarList = new String[pilarCount];
            //To get income group average values(income worksheet)
            string[] incomeHighAverageList = new String[pilarCount];
            string[] incomeUppermiddleAverageList = new String[pilarCount];
            string[] incomeLowermiddleAverageList = new String[pilarCount];
            string[] incomeLowAverageList = new String[pilarCount];
            //get income group Average rows (income group worksheet)
            int incomeHighRow = 56;
            int incomeHighSartColumn = 3;
            int incomeUppermiddleRow = 43;
            int incomeUppermiddleSartColumn = 12;
            int incomeLowermiddleRow = 39;
            int incomeLowermiddleSartColumn = 21;
            int incomeLowRow = 24;
            int incomeLowSartColumn = 30;
            int incomeCountryStartRow = 9;
            //To get region average values(region worksheet)
            string[] regionNAAverageList = new String[pilarCount];
            string[] regionCSAAverageList = new String[pilarCount];
            string[] regionEAverageList = new String[pilarCount];
            string[] regionLACAverageList = new String[pilarCount];
            string[] regionNAWAAverageList = new String[pilarCount];
            string[] regionSEAEAOAverageList = new String[pilarCount];
            string[] regionSSAAverageList = new String[pilarCount];
            //get region Average rows (region worksheet)
            int regionNARow = 11;
            int regionNAStartColumn = 2;
            int regionCSARow = 18;
            int regionCSAStartColumn = 11;
            int regionERow = 48;
            int regionEStartColumn = 20;
            int regionLACRow = 27;
            int regionLACStartColumn = 29;
            int regionNAWARow = 28;
            int regionNAWAStartColumn = 38;
            int regionSEAEAORow = 24;
            int regionSEAEAOStartColumn = 47;
            int regionSSARow = 33;
            int regionSSAStartColumn = 56;
            //To get Top10 average values
            // string[] Top10AverageList = new String[pilarCount];
            //get Top10 Average rows (Top 10 worksheet)
            int top10Row = 18;
            int top10StartColumn = 2;

            //To get graph2 data(graph2 worksheet)
            int graph2Table2Row = 21;
            int graph2Table2CountryColumn = 2;
            int graph2Table2IncomeColumn = 3;
            int graph2Table2RegionColumn = 4;
            int graph2Table2Top10Column = 5;

            //To get graph1 data(graph1 worksheet)
            int graph1TableRowStart = 6;
            int graph1TablePillarNameColumn = 1;
            int graph1TablePillarRankColumn = 2;
            int graph1TableRowCount = 8;

            Excel._Application xlAppIncomeRegion = new Excel.Application();
            Excel.Workbook workbookIncomeRegion = xlAppIncomeRegion.Workbooks.Open(wbkIncomeRegion);
            Excel.Worksheet worksheetIncome = workbookIncomeRegion.Sheets[workSheetNameIncome];
            Excel.Worksheet worksheetRegion = workbookIncomeRegion.Sheets[workSheetNameRegion];
            Excel.Worksheet worksheetGraph1 = workbookIncomeRegion.Sheets[workSheetNameGraph1];
            Excel.Worksheet worksheetGraph2 = workbookIncomeRegion.Sheets[workSheetNameGraph2];
            Excel.Worksheet worksheetTop10 = workbookIncomeRegion.Sheets[workSheetNameTop10];


            //bubble chart and GDP data excel file details
            string wbkInovationPerformance = ConfigurationManager.AppSettings.Get("wbkInovationPerformance");
            Console.WriteLine("Processing: " + wbkInovationPerformance);
            string workSheetNameGDPStatus = "F5 Bubble Data";
            //To get GDP column row(F5 Bubble Data worksheet)
            int GDPCountryListColumn = 1;
            int GDPStatusColumn = 7;
            int GDPStartRow = 6;
            int GDPLastRow = 131;

            Excel._Application xlAppInovationPerformance = new Excel.Application();
            Excel.Workbook workbookInovationPerformance = xlAppInovationPerformance.Workbooks.Open(wbkInovationPerformance);
            Excel.Worksheet worksheetGDPStatus = workbookInovationPerformance.Sheets[workSheetNameGDPStatus];

            //missing and outdated data excel file details
            string wbkMissingOutdated = ConfigurationManager.AppSettings.Get("wbkMissingOutdated");
            Console.WriteLine("Processing: " + wbkMissingOutdated);
            string workSheetNameMissingOutdatedFiltters = "App I";
            string workSheetNameMissingData = "Missing values";
            string workSheetNameOutdatedData = "Outdated data";
            //To fillter data by (missing data)n/a and (outdated data)$ (worksheet Ap I)
            int missingOutdatedCountryColumn = 6;
            int missingOutdatedCountryRow = 1;
            int filtterRowStart = 17;
            int filtterRowLast = 124;
            int missingDataFiltterColumn = 9;
            int outdatedDataFiltterColumn = 8;
            int filterIndicatorCodeColumn = 6;
            int filterPilarColumn = 7;

            //To get missing value row columns(Missing values worksheet)
            int missingValueIndicatorCodeColumn = 2;
            int missingValuePilarColumn = 3;
            int missingValueStartRow = 5;

            //To get outdated data row columns(Outdated data worksheet)
            int outdatedValueIndicatorCodeColumn = 2;
            int outdatedValuePilarColumn = 3;
            int outdatedValueStartRow = 5;


            Excel._Application xlAppMissingOutdated = new Excel.Application();
            Excel.Workbook workbookMissingOutdated = xlAppMissingOutdated.Workbooks.Open(wbkMissingOutdated);
            Excel.Worksheet worksheetMissingOutdatedFiltters = workbookMissingOutdated.Sheets[workSheetNameMissingOutdatedFiltters];
            Excel.Worksheet worksheetMissingData = workbookMissingOutdated.Sheets[workSheetNameMissingData];
            Excel.Worksheet worksheetOutdtedData = workbookMissingOutdated.Sheets[workSheetNameOutdatedData];

            Console.WriteLine("Creating pillar list from worksheet");
            //create pillar list from worksheet name : (region worksheet)
            for (int i = 0; i < pilarCount; i++)
            {
                pillarList[i] = (worksheetRegion.Cells[pillarListRow, i + 2] as Excel.Range).Value2.ToString();
                //To get graph1 (bar chart) y axis data
                graph1PillarList[i] = (worksheetRegion.Cells[pillarListRow, i + 2] as Excel.Range).Value2.ToString();

                //income group Average values
                incomeHighAverageList[i] = (worksheetIncome.Cells[incomeHighRow, i + incomeHighSartColumn] as Excel.Range).Value2.ToString();
                incomeUppermiddleAverageList[i] = (worksheetIncome.Cells[incomeUppermiddleRow, i + incomeUppermiddleSartColumn] as Excel.Range).Value2.ToString();
                incomeLowermiddleAverageList[i] = (worksheetIncome.Cells[incomeLowermiddleRow, i + incomeLowermiddleSartColumn] as Excel.Range).Value2.ToString();
                incomeLowAverageList[i] = (worksheetIncome.Cells[incomeLowRow, i + incomeLowSartColumn] as Excel.Range).Value2.ToString();

                //region average values
                regionNAAverageList[i] = (worksheetRegion.Cells[regionNARow, i + regionNAStartColumn] as Excel.Range).Value2.ToString();
                regionCSAAverageList[i] = (worksheetRegion.Cells[regionCSARow, i + regionCSAStartColumn] as Excel.Range).Value2.ToString();
                regionEAverageList[i] = (worksheetRegion.Cells[regionERow, i + regionEStartColumn] as Excel.Range).Value2.ToString();
                regionLACAverageList[i] = (worksheetRegion.Cells[regionLACRow, i + regionLACStartColumn] as Excel.Range).Value2.ToString();
                regionNAWAAverageList[i] = (worksheetRegion.Cells[regionNAWARow, i + regionNAWAStartColumn] as Excel.Range).Value2.ToString();
                regionSEAEAOAverageList[i] = (worksheetRegion.Cells[regionSEAEAORow, i + regionSEAEAOStartColumn] as Excel.Range).Value2.ToString();
                regionSSAAverageList[i] = (worksheetRegion.Cells[regionSSARow, i + regionSSAStartColumn] as Excel.Range).Value2.ToString();

                // Top10AverageList[i] = (worksheetTop10.Cells[top10Row, i + top10StartColumn] as Excel.Range).Value2.ToString();

                //set graph2 Top10 values
                worksheetGraph2.Cells[graph2Table2Row + 1, graph2Table2Top10Column] = (worksheetTop10.Cells[top10Row, i + top10StartColumn] as Excel.Range).Value2.ToString();
            }
            //To get graph1 (bar chart) y axis data
            graph1PillarList[7] = "Global Innovation Index 2018";

            Console.WriteLine("Generating reports for all countries");
            //generate report for all countries
            for (int i = GIICountryListRowStart; i <= GIICountryListRowEnd; i++)
            {
                // Path to a loadable document.
                string wordTemplatePath = ConfigurationManager.AppSettings.Get("wordTemplatePath");
                // Load a document intoDocumentCore.
                DocumentCore dc = DocumentCore.Load(wordTemplatePath);

                string country = (worksheetGII2018.Cells[i, GIICountryListColumn] as Excel.Range).Value2.ToString();
                Console.WriteLine("Processing - " + country + "CountryBriefReport.docx");
                worksheetStrengthWeakness.Cells[countryNameRow, countryNameColumn] = country; // change country name cell value for each country

                worksheetMissingOutdatedFiltters.Cells[missingOutdatedCountryRow, missingOutdatedCountryColumn] = country;

                string missingValuePilarsString = "";
                string missingValueIndicatorCodeString = "";
                string outdatedValuePilarsString = "";
                string outdatedValueIndicatorCodeString = "";
                for (int x = filtterRowStart; x <= filtterRowLast; x++)
                {
                    if ((worksheetMissingOutdatedFiltters.Cells[x, missingDataFiltterColumn] as Excel.Range).Value2.ToString() == "n/a")
                    {
                        string missingValuePilars = (worksheetMissingOutdatedFiltters.Cells[x, filterPilarColumn] as Excel.Range).Value2.ToString();
                        string missingValueIndicatorCode = (worksheetMissingOutdatedFiltters.Cells[x, filterIndicatorCodeColumn] as Excel.Range).Value2.ToString();

                        missingValuePilarsString = String.Concat(missingValuePilarsString, missingValuePilars + "?");
                        missingValueIndicatorCodeString = String.Concat(missingValueIndicatorCodeString, missingValueIndicatorCode + "?");
                    }
                    if ((worksheetMissingOutdatedFiltters.Cells[x, outdatedDataFiltterColumn] as Excel.Range).Value2.ToString() == "\u00A7".ToString())
                    {
                        string outdatedValuePilars = (worksheetMissingOutdatedFiltters.Cells[x, filterPilarColumn] as Excel.Range).Value2.ToString();
                        string outdatedValueIndicatorCode = (worksheetMissingOutdatedFiltters.Cells[x, filterIndicatorCodeColumn] as Excel.Range).Value2.ToString();

                        outdatedValuePilarsString = String.Concat(outdatedValuePilarsString, outdatedValuePilars + "?");
                        outdatedValueIndicatorCodeString = String.Concat(outdatedValueIndicatorCodeString, outdatedValueIndicatorCode + "?");
                    }
                }
                string[] missingValuePillarList = missingValuePilarsString.Split('?');
                string[] missingValueIndicatorCodeList = missingValueIndicatorCodeString.Split('?');
                string[] outdatedValuePillarList = outdatedValuePilarsString.Split('?');
                string[] outdatedValueIndicatorCodeList = outdatedValueIndicatorCodeString.Split('?');

                //missing values table as image
                for (int y = 0; y < missingValueIndicatorCodeList.Count() - 1; y++)
                {
                    worksheetMissingData.Cells[y + missingValueStartRow, missingValuePilarColumn] = missingValuePillarList[y];
                    worksheetMissingData.Cells[y + missingValueStartRow, missingValueIndicatorCodeColumn] = missingValueIndicatorCodeList[y];
                }
                int missingDataTableEndRange = missingValuePillarList.Count() + 3;
                Excel.Range rMissing = worksheetMissingData.Range["B4:F" + missingDataTableEndRange];
                rMissing.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                Bitmap missingDataImage = new Bitmap(System.Windows.Forms.Clipboard.GetImage());
                missingDataImage.Save(ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "2018missingDataImage.png");
                string picturePathMissingData = ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "2018missingDataImage.png";

                replaceImage("#missingDataTable#", dc, picturePathMissingData, missingDataImage, "full");

                //outdated Values table as image
                for (int y = 0; y < outdatedValueIndicatorCodeList.Count() - 1; y++)
                {
                    worksheetOutdtedData.Cells[y + outdatedValueStartRow, outdatedValuePilarColumn] = outdatedValuePillarList[y];
                    worksheetOutdtedData.Cells[y + outdatedValueStartRow, outdatedValueIndicatorCodeColumn] = outdatedValueIndicatorCodeList[y];
                }
                int outdatedDataTableEndRange = outdatedValuePillarList.Count() + 3;
                Excel.Range rOutdated = worksheetOutdtedData.Range["B4:F" + outdatedDataTableEndRange];
                rOutdated.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                Bitmap outdatedDataImage = new Bitmap(System.Windows.Forms.Clipboard.GetImage());
                outdatedDataImage.Save(ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "2018outdatedDataImage.png");
                string picturePathOutdatedData = ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "2018outdatedDataImage.png";

                replaceImage("#outdatedDataTable#", dc, picturePathOutdatedData, outdatedDataImage, "full");

                //Replace income group rank, region rank and there supper script values
                getValuesFromJson(GIIInputOutputRankIGroupRegionJsonData, i, dc, worksheetGII2018);
                FindAndReplaceSuperScriptValue(superScriptInputOutputRankIGroupRegionJsonData, i, dc, worksheetGII2018);

                //set header data for create graph2 (country, income group, region)
                string incomeGroupValue = (worksheetGII2018.Cells[i, incomeGroupColumn] as Excel.Range).Value2.ToString();
                string regionValue = (worksheetGII2018.Cells[i, regionColumn] as Excel.Range).Value2.ToString();
                worksheetGraph2.Cells[graph2Table2Row, graph2Table2CountryColumn] = country;
                worksheetGraph2.Cells[graph2Table2Row, graph2Table2IncomeColumn] = incomeGroupValue;
                worksheetGraph2.Cells[graph2Table2Row, graph2Table2RegionColumn] = regionValue;
                //set Income Group data for create graph2
                switch (incomeGroupValue)
                {
                    case "High":
                        //To set income group column (average) values 
                        for (int x = 0; x < incomeHighAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2IncomeColumn] = incomeHighAverageList[x];
                        }
                        //To set country column values 
                        int incomeHighScoreCount = 0;
                        string incomeHighScorePilars = "";
                        for (int z = incomeCountryStartRow; z < incomeHighRow; z++)
                        {
                            if ((worksheetIncome.Cells[z, incomeHighSartColumn - 1] as Excel.Range).Value2.ToString() == country)
                            {
                                for (int x = 0; x < pilarCount; x++)
                                {
                                    worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] = worksheetIncome.Cells[z, incomeHighSartColumn + x];
                                    if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(incomeHighAverageList[x]))
                                    {
                                        incomeHighScoreCount++;
                                        incomeHighScorePilars = incomeHighScorePilars + pillarList[x] + ", ";
                                    }
                                }
                            }
                        }
                        FindAndReplace("#noIGhighscorepillars#", incomeHighScoreCount.ToString(), dc);
                        FindAndReplace("#IGhighScorePillars#", incomeHighScorePilars, dc);
                        FindAndReplace("#noIncomegroupCountry#", (incomeHighRow - incomeCountryStartRow).ToString(), dc);
                        break;
                    case "Upper-middle":
                        for (int x = 0; x < incomeUppermiddleAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2IncomeColumn] = incomeUppermiddleAverageList[x];
                        }
                        incomeHighScoreCount = 0;
                        incomeHighScorePilars = "";
                        for (int z = incomeCountryStartRow; z < incomeUppermiddleRow; z++)
                        {
                            if ((worksheetIncome.Cells[z, incomeUppermiddleSartColumn - 1] as Excel.Range).Value2.ToString() == country)
                            {
                                for (int x = 0; x < pilarCount; x++)
                                {
                                    worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] = worksheetIncome.Cells[z, incomeUppermiddleSartColumn + x];
                                    if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(incomeUppermiddleAverageList[x]))
                                    {
                                        incomeHighScoreCount++;
                                        incomeHighScorePilars = incomeHighScorePilars + pillarList[x] + ", ";
                                    }
                                }
                            }
                        }
                        FindAndReplace("#noIGhighscorepillars#", incomeHighScoreCount.ToString(), dc);
                        FindAndReplace("#IGhighScorePillars#", incomeHighScorePilars, dc);
                        FindAndReplace("#noIncomegroupCountry#", (incomeUppermiddleRow - incomeCountryStartRow).ToString(), dc);
                        break;
                    case "Lower-middle":
                        for (int x = 0; x < incomeLowermiddleAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2IncomeColumn] = incomeLowermiddleAverageList[x];
                        }
                        incomeHighScoreCount = 0;
                        incomeHighScorePilars = "";
                        for (int z = incomeCountryStartRow; z < incomeLowermiddleRow; z++)
                        {
                            if ((worksheetIncome.Cells[z, incomeLowermiddleSartColumn - 1] as Excel.Range).Value2.ToString() == country)
                            {

                                for (int x = 0; x < pilarCount; x++)
                                {
                                    worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] = worksheetIncome.Cells[z, incomeLowermiddleSartColumn + x];
                                    if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(incomeLowermiddleAverageList[x]))
                                    {
                                        incomeHighScoreCount++;
                                        incomeHighScorePilars = incomeHighScorePilars + pillarList[x] + ", ";
                                    }
                                }

                            }
                        }
                        FindAndReplace("#noIGhighscorepillars#", incomeHighScoreCount.ToString(), dc);
                        FindAndReplace("#IGhighScorePillars#", incomeHighScorePilars, dc);
                        FindAndReplace("#noIncomegroupCountry#", (incomeLowermiddleRow - incomeCountryStartRow).ToString(), dc);
                        break;
                    case "Low":
                        for (int x = 0; x < incomeLowAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2IncomeColumn] = incomeLowAverageList[x];
                        }
                        incomeHighScoreCount = 0;
                        incomeHighScorePilars = "";
                        for (int z = incomeCountryStartRow; z < incomeLowRow; z++)
                        {
                            if ((worksheetIncome.Cells[z, incomeLowSartColumn - 1] as Excel.Range).Value2.ToString() == country)
                            {

                                for (int x = 0; x < pilarCount; x++)
                                {
                                    worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] = worksheetIncome.Cells[z, incomeLowSartColumn + x];
                                    if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(incomeLowAverageList[x]))
                                    {
                                        incomeHighScoreCount++;
                                        incomeHighScorePilars = incomeHighScorePilars + pillarList[x] + ", ";
                                    }
                                }

                            }
                        }
                        FindAndReplace("#noIGhighscorepillars#", incomeHighScoreCount.ToString(), dc);
                        FindAndReplace("#IGhighScorePillars#", incomeHighScorePilars, dc);
                        FindAndReplace("#noIncomegroupCountry#", (incomeLowRow - incomeCountryStartRow).ToString(), dc);
                        break;
                }
                //set Region data
                switch (regionValue)
                {
                    //To set region column (average) values
                    case "Northern America":
                        for (int x = 0; x < regionNAAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2RegionColumn] = regionNAAverageList[x];
                        }
                        string regionAboveAvgPillars = "";
                        string regionBelowAvgPillars = "";
                        for (int x = 0; x < pilarCount; x++)
                        {
                            if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(regionNAAverageList[x]))
                            {
                                regionAboveAvgPillars = regionAboveAvgPillars + pillarList[x] + ", ";
                            }
                            else
                            {
                                regionBelowAvgPillars = regionBelowAvgPillars + pillarList[x] + ", ";
                            }
                        }
                        setRegionAverageAboveBelowText(regionAboveAvgPillars, regionBelowAvgPillars, dc);
                        break;
                    case "Central and Southern Asia":
                        for (int x = 0; x < regionCSAAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2RegionColumn] = regionCSAAverageList[x];
                        }
                        regionAboveAvgPillars = "";
                        regionBelowAvgPillars = "";
                        for (int x = 0; x < pilarCount; x++)
                        {
                            if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(regionCSAAverageList[x]))
                            {
                                regionAboveAvgPillars = regionAboveAvgPillars + pillarList[x] + ", ";
                            }
                            else
                            {
                                regionBelowAvgPillars = regionBelowAvgPillars + pillarList[x] + ", ";
                            }
                        }
                        setRegionAverageAboveBelowText(regionAboveAvgPillars, regionBelowAvgPillars, dc);
                        break;
                    case "Europe":
                        for (int x = 0; x < regionEAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2RegionColumn] = regionEAverageList[x];
                        }
                        regionAboveAvgPillars = "";
                        regionBelowAvgPillars = "";
                        for (int x = 0; x < pilarCount; x++)
                        {
                            if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(regionEAverageList[x]))
                            {
                                regionAboveAvgPillars = regionAboveAvgPillars + pillarList[x] + ", ";
                            }
                            else
                            {
                                regionBelowAvgPillars = regionBelowAvgPillars + pillarList[x] + ", ";
                            }
                        }
                        setRegionAverageAboveBelowText(regionAboveAvgPillars, regionBelowAvgPillars, dc);
                        break;
                    case "Latin America and the Caribbean":
                        for (int x = 0; x < regionLACAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2RegionColumn] = regionLACAverageList[x];
                        }
                        regionAboveAvgPillars = "";
                        regionBelowAvgPillars = "";
                        for (int x = 0; x < pilarCount; x++)
                        {
                            if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(regionLACAverageList[x]))
                            {
                                regionAboveAvgPillars = regionAboveAvgPillars + pillarList[x] + ", ";
                            }
                            else
                            {
                                regionBelowAvgPillars = regionBelowAvgPillars + pillarList[x] + ", ";
                            }
                        }
                        setRegionAverageAboveBelowText(regionAboveAvgPillars, regionBelowAvgPillars, dc);
                        break;
                    case "Northern Africa and Western Asia":
                        for (int x = 0; x < regionNAWAAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2RegionColumn] = regionNAWAAverageList[x];
                        }
                        regionAboveAvgPillars = "";
                        regionBelowAvgPillars = "";
                        for (int x = 0; x < pilarCount; x++)
                        {
                            if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(regionNAWAAverageList[x]))
                            {
                                regionAboveAvgPillars = regionAboveAvgPillars + pillarList[x] + ", ";
                            }
                            else
                            {
                                regionBelowAvgPillars = regionBelowAvgPillars + pillarList[x] + ", ";
                            }
                        }
                        setRegionAverageAboveBelowText(regionAboveAvgPillars, regionBelowAvgPillars, dc);
                        break;
                    case "South East Asia, East Asia, and Oceania":
                        for (int x = 0; x < regionSEAEAOAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2RegionColumn] = regionSEAEAOAverageList[x];
                        }
                        regionAboveAvgPillars = "";
                        regionBelowAvgPillars = "";
                        for (int x = 0; x < pilarCount; x++)
                        {
                            if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(regionSEAEAOAverageList[x]))
                            {
                                regionAboveAvgPillars = regionAboveAvgPillars + pillarList[x] + ", ";
                            }
                            else
                            {
                                regionBelowAvgPillars = regionBelowAvgPillars + pillarList[x] + ", ";
                            }
                        }
                        setRegionAverageAboveBelowText(regionAboveAvgPillars, regionBelowAvgPillars, dc);
                        break;
                    case "Sub-Saharan Africa":
                        for (int x = 0; x < regionSSAAverageList.Count(); x++)
                        {
                            worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2RegionColumn] = regionSSAAverageList[x];
                        }
                        regionAboveAvgPillars = "";
                        regionBelowAvgPillars = "";
                        for (int x = 0; x < pilarCount; x++)
                        {
                            if (Double.Parse((worksheetGraph2.Cells[graph2Table2Row + 1 + x, graph2Table2CountryColumn] as Excel.Range).Value2.ToString()) > Double.Parse(regionSSAAverageList[x]))
                            {
                                regionAboveAvgPillars = regionAboveAvgPillars + pillarList[x] + ", ";
                            }
                            else
                            {
                                regionBelowAvgPillars = regionBelowAvgPillars + pillarList[x] + ", ";
                            }
                        }
                        setRegionAverageAboveBelowText(regionAboveAvgPillars, regionBelowAvgPillars, dc);
                        break;
                }

                //graph2 as image
                Excel.Range rGraph2 = worksheetGraph2.Range["A33:H63"];
                rGraph2.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                Bitmap graph2Image = new Bitmap(System.Windows.Forms.Clipboard.GetImage());
                graph2Image.Save(ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "graph2Image.png");
                string picturePathGraph2 = ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "graph2Image.png";

                replaceImage("#graph2#", dc, picturePathGraph2, graph2Image, "graphQuater");

                //To get GDP level
                string GDPStatusLevel = "";
                for (int x = GDPStartRow; x <= GDPLastRow; x++)
                {
                    if ((worksheetGDPStatus.Cells[x, GDPCountryListColumn] as Excel.Range).Value2.ToString() == country)
                    {
                        string GDPValue = (worksheetGDPStatus.Cells[x, GDPStatusColumn] as Excel.Range).Value2.ToString();
                        switch (GDPValue)
                        {
                            case "Innovation Achievers":
                                GDPStatusLevel = "above";
                                break;
                            case "Leader":
                                GDPStatusLevel = "above";
                                break;
                            case "Performing at development":
                                GDPStatusLevel = "at";
                                break;
                            case "Performing below development":
                                GDPStatusLevel = "below";
                                break;
                        }
                    }
                }
                //generate strength and weakness bullet list
                int strengthCount = 0;
                string strengthText = "";
                int weaknessCount = 0;
                string weaknessText = "";

                for (int j = 0; j <= pillarListLastColumn - pillarListStartColumn; j++)
                {
                    //generate strength bullet list
                    string strengthSubPillarListString = "";
                    string strengthPillarSubpillarIndicatorListString = "";
                    if ((worksheetStrengthWeakness.Cells[pillarRowList[j], StrengthWeaknessColumn] as Excel.Range).Value2.ToString() == "@")
                    {
                        strengthCount++;
                    }
                    int subPillarStrengthCount = 0;
                    for (int k = subPillarStartRowList[j]; k <= subPillarLastRowList[j]; k++)
                    {
                        if ((worksheetStrengthWeakness.Cells[k, StrengthWeaknessColumn] as Excel.Range).Value2.ToString() == "@")
                        {
                            subPillarStrengthCount++;
                            string strengthPillarSubpillarIndicator = (worksheetStrengthWeakness.Cells[k, pillarSubpillarIndicatorColumn] as Excel.Range).Value2.ToString();
                            string strengthPillarSubpillar = (worksheetStrengthWeakness.Cells[k, subPillarListColumn] as Excel.Range).Value2.ToString();

                            strengthSubPillarListString = String.Concat(strengthSubPillarListString, strengthPillarSubpillar + "?");
                            strengthPillarSubpillarIndicatorListString = String.Concat(strengthPillarSubpillarIndicatorListString, strengthPillarSubpillarIndicator + "?");

                        }
                    }
                    if (subPillarStrengthCount > 0)
                    {
                        string[] strengthSubPillarList = strengthSubPillarListString.Substring(0, strengthSubPillarListString.Length - 1).Split('?');
                        string[] strengthPillarSubpillarIndicatorList = strengthPillarSubpillarIndicatorListString.Substring(0, strengthPillarSubpillarIndicatorListString.Length - 1).Split('?');

                        strengthText = strengthText + "\u23FA".ToString() + " In " + pillarList[j] + ", " + country + " exhibits strength in the " + string.Join(", ", strengthSubPillarList) + " and in " + string.Join(", ", strengthPillarSubpillarIndicatorList) + "\n";
                    }
                    //generate weakness bullet list
                    string weaknessSubPillarListString = "";
                    string weaknessPillarSubpillarIndicatorListString = "";
                    if ((worksheetStrengthWeakness.Cells[pillarRowList[j], StrengthWeaknessColumn] as Excel.Range).Value2.ToString() == "!")
                    {
                        weaknessCount++;
                    }
                    int subPillarWeaknessCount = 0;
                    for (int k = subPillarStartRowList[j]; k <= subPillarLastRowList[j]; k++)
                    {
                        if ((worksheetStrengthWeakness.Cells[k, StrengthWeaknessColumn] as Excel.Range).Value2.ToString() == "!")
                        {
                            subPillarWeaknessCount++;
                            string weaknesshPillarSubpillarIndicator = (worksheetStrengthWeakness.Cells[k, pillarSubpillarIndicatorColumn] as Excel.Range).Value2.ToString();
                            string weaknessPillarSubpillar = (worksheetStrengthWeakness.Cells[k, subPillarListColumn] as Excel.Range).Value2.ToString();

                            weaknessSubPillarListString = String.Concat(weaknessSubPillarListString, weaknessPillarSubpillar + "?");
                            weaknessPillarSubpillarIndicatorListString = String.Concat(weaknessPillarSubpillarIndicatorListString, weaknesshPillarSubpillarIndicator + "?");
                        }
                    }
                    if (subPillarWeaknessCount > 0)
                    {
                        string[] weaknessSubPillarList = weaknessSubPillarListString.Substring(0, weaknessSubPillarListString.Length - 1).Split('?');
                        string[] weaknessPillarSubpillarIndicatorList = weaknessPillarSubpillarIndicatorListString.Substring(0, weaknessPillarSubpillarIndicatorListString.Length - 1).Split('?');
                        weaknessText = weaknessText + "\u23FA".ToString() + " In " + pillarList[j] + ", " + country + " exhibits weakness in the " + string.Join(", ", weaknessSubPillarList) + " and in " + string.Join(", ", weaknessPillarSubpillarIndicatorList) + "\n";
                    }
                }

                for (int y = incomeGroupRegionCountryRowStart; y <= incomeGroupRegionCountryRowEnd; y++)
                {
                    if ((worksheetIncomeGroupRegionRanking.Cells[y, incomeGroupRegionCountryColumn] as Excel.Range).Value2.ToString() == country)
                    {
                        getValuesFromJson(incomeGroupRegionRankJsonData, y, dc, worksheetIncomeGroupRegionRanking);
                        FindAndReplaceSuperScriptValue(superScriptIncomeGroupRegionRankJsonData, y, dc, worksheetIncomeGroupRegionRanking);
                    }
                }

                //To get graph1 (bar chart) data
                int[] pillarRankList = new int[8];
                int[] indexRankList = new int[8];
                for (int y = 0; y < pilarCount; y++)
                {
                    pillarRankList[y] = Convert.ToInt32((worksheetStrengthWeakness.Cells[graph1PillarRowList[y], pillarRankColumn] as Excel.Range).Value2);
                    indexRankList[y] = y;
                }
                pillarRankList[7] = Convert.ToInt32((worksheetStrengthWeakness.Cells[giiRankRow, pillarRankColumn] as Excel.Range).Value2);
                indexRankList[7] = 7;
                //To set graph1 (bar chart) data
                int p, q, tempPillarRank, tempPillarIndex;
                for (p = 0; p < graph1TableRowCount; p++)
                {
                    for (q = p + 1; q < graph1TableRowCount; q++)
                    {
                        if (pillarRankList[p] < pillarRankList[q])
                        {
                            tempPillarRank = pillarRankList[p];
                            pillarRankList[p] = pillarRankList[q];
                            pillarRankList[q] = tempPillarRank;

                            tempPillarIndex = indexRankList[p];
                            indexRankList[p] = indexRankList[q];
                            indexRankList[q] = tempPillarIndex;
                        }
                    }
                }
                for (p = 0; p < graph1TableRowCount; p++)
                {
                    worksheetGraph1.Cells[graph1TableRowStart + p, graph1TablePillarNameColumn] = graph1PillarList[indexRankList[p]];
                    worksheetGraph1.Cells[graph1TableRowStart + p, graph1TablePillarRankColumn] = pillarRankList[p];
                }
                //graph1 (bar chart) as image
                Excel.Range rGraph1 = worksheetGraph1.Range["F2:N22"];
                rGraph1.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                Bitmap graph1Image = new Bitmap(System.Windows.Forms.Clipboard.GetImage());
                graph1Image.Save(ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "graph1Image.png");
                string picturePathGraph1 = ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "graph1Image.png";

                replaceImage("#graph1#", dc, picturePathGraph1, graph1Image, "graphHalf");
                //getValuesFromJson(incomeRegionJsonData, i + 10, dc, worksheetGII);


                FindAndReplace("#noIGroup#", incomeGroupCount.ToString(), dc);
                FindAndReplace("#noR#", regionCount.ToString(), dc);
                FindAndReplace("#noPStrengths#", strengthCount.ToString(), dc);
                FindAndReplace("#noPWeakness#", weaknessCount.ToString(), dc);
                FindAndReplace("#strengthText#", strengthText, dc);
                FindAndReplace("#weaknessText#", weaknessText, dc);
                FindAndReplace("#GDPabove/below/at#", GDPStatusLevel, dc);

                // Save our document into docx format.
                string savePath = ConfigurationManager.AppSettings.Get("GIICountryReportsPath") + country + "CountryBriefReport.docx";
                dc.Save(savePath);

                // Open the result for demonstation purposes.
                //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
                //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
            }

            Console.WriteLine("Document generation completed!");
        }
        public static void replaceImage(string findText, DocumentCore dc, string pictPath, Bitmap image, string imageDisplaySize)
        {
            Regex regex = new Regex(@findText, RegexOptions.IgnoreCase);
            Picture picture;
            int width;
            int height;
            if (imageDisplaySize == "full")
            {
                width = 550;
                height = image.Height;
                picture = new Picture(dc, InlineLayout.Inline(new SautinSoft.Document.Drawing.Size(width, height)), pictPath);
            }
            else if (imageDisplaySize == "graphQuater")
            {
                width = image.Width / 4;
                height = image.Height / 4;
                picture = new Picture(dc, pictPath);
                picture.Layout = FloatingLayout.Floating(
                    new HorizontalPosition(10, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                    new VerticalPosition(40, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                    new SautinSoft.Document.Drawing.Size(width, height));
                (picture.Layout as FloatingLayout).WrappingStyle = WrappingStyle.Square;
            }
            else
            {
                width = image.Width / 2;
                height = image.Height / 2;
                picture = new Picture(dc, InlineLayout.Inline(new SautinSoft.Document.Drawing.Size(width, height)), pictPath);
            }

            foreach (ContentRange item in dc.Content.Find(regex).Reverse())
            {
                item.Replace(picture.Content);
            }
        }
        public static void FindAndReplace(string findValue, string newValue, DocumentCore dc)
        {
            Regex regexValue = new Regex(@findValue, RegexOptions.IgnoreCase);

            foreach (ContentRange itemValue in dc.Content.Find(regexValue).Reverse())
            {
                itemValue.Replace(newValue);
            }
        }
        public static void FindAndReplaceSuperScriptValue(String jsonData, int row, DocumentCore dc, Excel.Worksheet worksheet)
        {
            var detail = JArray.Parse(jsonData);
            JArray array = detail;
            foreach (var item in array)
            {
                string rank = (worksheet.Cells[row, (int)item["column"]] as Excel.Range).Value2.ToString();
                int rankValue = Int32.Parse(rank);
                if (rankValue >= 11 && rankValue <= 13)
                {
                    FindAndReplace(item["key"].ToString(), "th", dc);
                }
                else
                {
                    switch (rank.Substring(rank.Length - 1, 1))
                    {
                        case "1":
                            FindAndReplace(item["key"].ToString(), "st", dc);
                            break;
                        case "2":
                            FindAndReplace(item["key"].ToString(), "nd", dc);
                            break;
                        case "3":
                            FindAndReplace(item["key"].ToString(), "rd", dc);
                            break;
                        default:
                            FindAndReplace(item["key"].ToString(), "th", dc);
                            break;
                    }
                }
            }
        }
        public static void getValuesFromJson(String jsonData, int row, DocumentCore dc, Excel.Worksheet worksheet)
        {
            var detail = JArray.Parse(jsonData);
            JArray array = detail;
            foreach (var item in array)
            {
                FindAndReplace(item["key"].ToString(), item["valueType"].ToString(), row, (int)item["column"], dc, worksheet);
            }
        }
        public static void FindAndReplace(string findValue, string valueType, int excelRow, int excelColumn, DocumentCore dc, Excel.Worksheet worksheet)
        {
            Regex regexValue = new Regex(@findValue, RegexOptions.IgnoreCase);

            foreach (ContentRange itemValue in dc.Content.Find(regexValue).Reverse())
            {
                if (valueType == "string")
                {
                    itemValue.Replace((worksheet.Cells[excelRow, excelColumn] as Excel.Range).Value2.ToString());
                }
                else
                {
                    itemValue.Replace((worksheet.Cells[excelRow, excelColumn] as Excel.Range).Value2.ToString("0.0"));
                }

            }
        }
        public static void setRegionAverageAboveBelowText(string averageAbovePillarText, string averageBelowpillarText, DocumentCore dc)
        {
            if (averageAbovePillarText != "")
            {
                FindAndReplace("#aboveAvgText#", "above average in", dc);
            }
            if (averageAbovePillarText == "" && averageBelowpillarText != "")
            {
                FindAndReplace("#aboveAvgText#", "", dc);
                FindAndReplace("#belowAvgText#", "below average in", dc);
            }
            if (averageAbovePillarText != "" && averageBelowpillarText != "")
            {
                FindAndReplace("#belowAvgText#", "and below average in", dc);
            }
            if (averageBelowpillarText == "")
            {
                FindAndReplace("#belowAvgText#", "", dc);
            }

            FindAndReplace("#regionAboveAvgPillars#", (averageAbovePillarText != "" ? averageAbovePillarText.Substring(0, averageAbovePillarText.Length - 2) : averageAbovePillarText), dc);
            FindAndReplace("#regionBelowAvgPillars#", (averageBelowpillarText != "" ? averageBelowpillarText.Substring(0, averageBelowpillarText.Length - 2) : averageBelowpillarText), dc);
        }
        public static void UploadCountryBriefReports(string year, string reportType)
        {
            Console.WriteLine("Uploading CountryBriefReports...");
            // Parse the connection string and return a reference to the storage account.
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings.Get("StorageConnectionString"));

            // Create a CloudFileClient object for credentialed access to Azure Files.
            CloudFileClient fileClient = storageAccount.CreateCloudFileClient();

            // Get a reference to the file share we created previously.
            CloudFileShare share = fileClient.GetShareReference(ConfigurationManager.AppSettings.Get("FileShareReference"));

            // Ensure that the share exists.
            if (share.Exists())
            {
                // Get a reference to the root directory for the share.
                CloudFileDirectory rootDir = share.GetRootDirectoryReference();

                // Get a reference to the directory we created previously.
                CloudFileDirectory sampleDir = rootDir.GetDirectoryReference("FileStorage\\Reports\\" + reportType + "\\" + year);

                // Ensure that the directory exists.
                if (sampleDir.Exists())
                {
                    // Get a reference to the file we created previously.

                    string sourceDirectory = @"C:\GIICountryReports";

                    var reportFiles = Directory.EnumerateFiles(sourceDirectory, "*.docx", SearchOption.TopDirectoryOnly);

                    foreach (string currentFile in reportFiles)
                    {
                        string fileName = currentFile.Substring(sourceDirectory.Length + 1);
                        Console.WriteLine("Uploading " + fileName);
                        byte[] fileByteArray = File.ReadAllBytes(currentFile);

                        // Create a new pdf file in the root directory.
                        CloudFile sourceFileImage = sampleDir.GetFileReference(fileName);
                        sourceFileImage.UploadFromByteArray(fileByteArray, 0, fileByteArray.Count<byte>());

                        File.Delete(currentFile);

                    }
                }
            }
            Console.WriteLine("Uploading Completed!");
        }

        private static void KillAllExcelFileProcesses()
        {
            Console.WriteLine("Killing all EXCEL processes");
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
                process.Kill();
        }

        private static void ClearImages()
        {
            var iamgeFiles = Directory.EnumerateFiles(ConfigurationManager.AppSettings.Get("GIICountryReportsPath"), "*.png", SearchOption.TopDirectoryOnly);

            foreach (string currentFile in iamgeFiles)
                File.Delete(currentFile);
        }
    }
}
