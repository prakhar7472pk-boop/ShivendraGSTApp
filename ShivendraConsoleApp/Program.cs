using ClosedXML.Excel;
using Microsoft.Playwright;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ShivendraConsoleApp;

static class Program
{
    static Program()
    {
        DefaultFileSuffix = InputPath = OutputFileName = string.Empty;
        ConfigReader.UpdateConfig();

        // Column names
        foreach (var column in ColumnNum)
        {
            Sheet.Cell(1, column.Value).Value = column.Key;
        }
    }

    internal static string InputPath;
    internal static string OutputFileName;
    internal static string DefaultFileSuffix;
    internal static string? OutputPath = null;
    internal static readonly string[] SupportedOutputExcelFormats = [".xlsx", ".xlsm", ".xltx"];
    internal static readonly Dictionary<string, int> ColumnNum = new();
    internal static int typing_delay = 50;

    //private static int _row = 2;

    // Normalize: Set fixed width and height
    internal static double FixedColumnWidth = 25;
    internal static double FixedRowHeight = 15;

    #region Constants

    // site constants
    private const string SiteUrl = "https://services.gst.gov.in/services/searchtp";
    private const string InputGstid = "input[name='for_gstin']";
    private const string CaptchaInput = "input[name='cap']";
    private const string gap = " ";

    // Column constants
    private const string GstinUin = "GSTIN/UIN";
    private const string AdministrativeOffice = "Administrative Office";
    private const string OtherOffice = "Other Office";
    private const string MainOffice = "Center / State";
    private const string Central_ = "Central ";
    private const string Zone = "Zone";
    private const string Commissionerate = "Commissionerate";
    private const string Division = "Division";
    private const string Range = "Range";
    private const string Jurisdiction = "JURISDICTION";
    private const string Center = "CENTER";
    private const string State = "State";
    private const string Charge = "Charge";
    private const string Circle = "Circle";
    private const string Ward = "Ward";
    private const string Sector = "Sector";
    private const string Unit = "Unit";
    private const string District = "District";
    private const string Headquarter = "Headquarter";
    private const string AC_or_CTO_Ward = "AC / CTO Ward";
    private const string LOCAL_GST_Office = "LOCAL GST Office";
    private const string Goods = "Goods";
    private const string Services = "Services";
    //private const string Zone = "Central Zone";
    //private const string Commissionerate = "Central Commissionerate";
    //private const string Division = "Central Division";
    //private const string Range = "Central Range";

    #endregion

    // Column lists
    private static readonly string[] Zone_Commissionerate = new[] { Zone, Commissionerate };

    private static readonly string[] Division_level = new[] { Division };

    private static readonly string[] Sub_division = new[]
        { Range, Circle, Ward, Unit, Charge, Sector, District, Headquarter, LOCAL_GST_Office, AC_or_CTO_Ward };

    private static readonly XLWorkbook Workbook = new XLWorkbook();
    private static readonly IXLWorksheet Sheet = Workbook.Worksheets.Add("Parsed HTML");

    public static async Task Main()
    {
        Console.Write("Enter file path - ");
        string? path = Console.ReadLine();

        if (string.IsNullOrEmpty(path))
        {
            path = InputPath;
            if (string.IsNullOrEmpty(path))
            {
                Console.WriteLine("Input path not valid");
                return;
            }

            path.Trim('"');
        }
        else
        {
            path = path.Trim('"');
            OutputFileName = path.Split("\\").Last().Split('.').First() + DefaultFileSuffix;
        }

        using var playwright = await Playwright.CreateAsync();

        var chromePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";

        var browser = await playwright.Chromium.LaunchAsync(new()
        {
            Headless = false,
            ExecutablePath = chromePath
        });

        var page = await browser.NewPageAsync();

        string[] gstIds = await ReadWriteOperations.GetGstIdsAsync(path);
        IdIterator.Configure(gstIds);

        var pageLoadtsk = page.GotoAsync(SiteUrl);

        CancellationTokenSource cts = new();
        page.Load += async (_, _) =>
        {
            cts.Cancel();
            cts = new();
            var token = cts.Token;

            int? idx = IdIterator.GetCurrentIdx();

            if (idx is null)
            {
                Environment.Exit(0);
                return;
            }

            string id = gstIds[idx.Value];
            var input = id.Trim();
            if (string.IsNullOrEmpty(input))
            {
                IdIterator.Complete(token);
                if (!token.IsCancellationRequested) 
                    await page.ReloadAsync();
                return;
            }

            try
            {
                int waitForSiteOpen = 0;
                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        await pageLoadtsk;
                        await page.FocusAsync(InputGstid);
#pragma warning disable CS0612 // Type or member is obsolete
                        await page.FillAsync(InputGstid, ""); // This will replace existing text
                        await page.TypeAsync(InputGstid, input, new() { Delay = typing_delay });

                        await page.Keyboard.PressAsync("Tab"); // Simulates global tab key press
                        await GetDataInXml(page, input, idx.Value + 2, token);
#pragma warning restore CS0612 // Type or member is obsolete
                        break;
                    }
                    catch
                    {
                        if (++waitForSiteOpen >= 10)
                        {
                            Console.WriteLine("Website took too long to load");
                            break;
                        }

                        try
                        {
                            await Task.Delay(100, token);
                        }
                        catch
                        {

                        }
                    }
                }

                ReadWriteOperations.HandleFileUsedByProcessException(Workbook, token);
                IdIterator.Complete(token);
                if (!token.IsCancellationRequested) 
                    await page.ReloadAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unable to process data for GSTID - {input}");
                Console.WriteLine($"Error - {e.Message}");
                throw;
            }
        };

        Console.ReadKey();
        await browser.CloseAsync();
    }

    internal static bool PageLoadSuccess;

    private static async Task GetDataInXml(IPage page, string originalId, int _row, CancellationToken token)
    {
        PageLoadSuccess = false;
        var cts = new CancellationTokenSource();

        GSTPageContentLoader.alreadyPromptedError = false;

        var tsk = Task.Run(async () => await GSTPageContentLoader.LoadPageContents(page, cts.Token));
        var tsk2 = Task.Run(async () => await GSTPageContentLoader.InvalidGstIdHandler(page, originalId, cts.Token));

        await Task.WhenAny(tsk, tsk2);
        cts.Cancel();

        if (!PageLoadSuccess)
        {
            Sheet.Cell(_row, ColumnNum[GstinUin]).Value = originalId;
            return;
        }

        string gstId = await page.InnerTextAsync("div.col-sm-6 > h4");
        gstId = gstId.Split(":").Last().Trim();
        
        if (string.IsNullOrEmpty(gstId)) gstId = originalId;

        var strongElements = await page.QuerySelectorAllAsync("strong");

        var data = new Dictionary<string, string>();
        
        foreach (var column in ColumnNum.Keys) data[column] = string.Empty;
        data[GstinUin] = gstId;

        int col = 2;
        foreach (var element in strongElements)
        {
            string value = await element.InnerTextAsync();

            // Get the parent <p> of <strong>
            var parentP = await element.EvaluateHandleAsync("el => el.parentElement");
            var nextP = await parentP.EvaluateHandleAsync("el => el.nextElementSibling");

            try
            {
                var jsHandle = await nextP.EvaluateHandleAsync(@"el => {
                    // Adjust selector as needed (e.g., 'li', 'div', 'tr td', etc.)
                    return Array.from(el.querySelectorAll(':scope > *')).map(child => child.textContent.trim());
                }");

                StringBuilder sb = new();

                string[] list = await jsHandle.JsonValueAsync<string[]>();

                if (list.Length > 0)
                {
                    if (value.Equals(AdministrativeOffice) || value.Equals(OtherOffice))
                    {
                        string[] strs = list[0].Split('(', '-', ')').Where(s => !s.Equals(string.Empty)).ToArray();

                        if (value.Equals(AdministrativeOffice))
                        {
                            data[MainOffice] = list.Where(entry 
                                    => string.Equals(Jurisdiction, entry.Split('(', '-', ')')
                                                    .First(s => !s.Equals(string.Empty)).Trim(), 
                                                    StringComparison.OrdinalIgnoreCase))
                                .Select(s 
                                    => s.Split('(', '-', ')').Last(s => !s.Equals(string.Empty)))
                                .First()?.Trim()!;
                        }

                        if (strs[0].Trim().Equals(Jurisdiction) && strs[^1].Trim().Equals(Center))
                        {
                            foreach (var str in list)
                            {
                                string? title = str.Split('(', '-', ')').FirstOrDefault(s => !s.Equals(string.Empty))?.Trim();

                                if (title is null) continue;

                                if (string.Equals(Zone, title, StringComparison.OrdinalIgnoreCase))
                                {
                                    data[Central_ + Zone] = str.Substring(7);
                                }
                                else if (string.Equals(Commissionerate, title, StringComparison.OrdinalIgnoreCase))
                                {
                                    data[Central_ + Commissionerate] = str.Substring(17);
                                }
                                else if (string.Equals(Division, title, StringComparison.OrdinalIgnoreCase))
                                {
                                    data[Central_ + Division] = str.Substring(11);
                                }
                                else if (string.Equals(Range, title, StringComparison.OrdinalIgnoreCase))
                                {
                                    data[Central_ + Range] = str.Substring(8);
                                }
                            }
                        }
                        else
                        {
                            foreach (var item in list)
                            {
                                string str = item.Trim();
                                string? val = Helper.GetFieldValue(str, State);
                                if (val is not null)
                                {
                                    data[State] = data[State] + Environment.NewLine + val;
                                    continue;
                                }

                                val = Helper.GetFieldValue(str, Zone_Commissionerate);
                                if (val is not null)
                                {
                                    data[State + gap + Zone] = data[State + gap + Zone] + Environment.NewLine + val;
                                    continue;
                                }

                                val = Helper.GetFieldValue(str, Division_level);
                                if (val is not null)
                                {
                                    data[State + gap + Division] = data[State + gap + Division] + Environment.NewLine + val;
                                    continue;
                                }

                                val = Helper.GetFieldValue(str, Sub_division);
                                if (val is not null)
                                {
                                    data[State + gap + Charge] = data[State + gap + Charge] + Environment.NewLine + val;
                                }
                            }
                        }
                    }

                    foreach (var item in list)
                    {
                        sb.AppendLine(item);
                    }

                    var value2 = sb.ToString();
                    if (!string.IsNullOrEmpty(value2))
                    {
                        data[value] = value2;
                    }
                }
                else
                {
                    string value2 = "";
                    if (nextP is IElementHandle elementHandle)
                    {
                        value2 = await elementHandle.InnerTextAsync();
                        data[value] = value2;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception for Title - {value}. Error - {ex.Message}");
            }
        }

        var element2 = await page.QuerySelectorAsync("div[ng-if='!goodServErrMsg']");
        if (element2 is null) return;

        var table = await element2.QuerySelectorAsync("table");

        if (table == null)
        {
            CommitDataToSheet(data, _row);
            Console.WriteLine("Table not found.");
            if (!tsk.IsCompleted) await tsk;
            else await tsk2;
            return;
        }

        // Get all rows (both thead and tbody)
        var rowsQuery = await table.QuerySelectorAllAsync("tr");

        StringBuilder goods = new(), services = new();
        for (int i = 0; i < rowsQuery.Count; i++)
        {
            if (i <= 1) continue;

            var rowQuery = rowsQuery[i];
            var cells = await rowQuery.QuerySelectorAllAsync("th, td"); // handle both header and data cells

            int colIdx = col;
            bool mergeTwoCol = false, isGoods = true;
            string colVal = "";
            foreach (var cell in cells)
            {
                var text = await cell.InnerTextAsync();

                if (mergeTwoCol)    
                {
                    if (isGoods)
                    {
                        if (!string.IsNullOrEmpty(colVal + text))
                            goods.AppendLine(colVal + " : " + text);
                        isGoods = false;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(colVal + text))
                            services.AppendLine(colVal + " : " + text);
                        isGoods = true;
                    }
                    mergeTwoCol = false;
                }
                else
                {
                    colVal = text;
                    mergeTwoCol = true;
                }
            }
        }

        data[Goods] = goods.ToString();
        data[Services] = services.ToString();

        CommitDataToSheet(data, _row);

        if (!tsk.IsCompleted) await tsk;
        else await tsk2;
    }

    private static void CommitDataToSheet(Dictionary<string, string> data, int _row)
    {
        // update column values for different gstin/uin
        foreach (var dataPair in ColumnNum)
        {
            int currentCol = dataPair.Value;

            Sheet.Cell(_row, currentCol).Value = data[dataPair.Key].Trim(' ', '-');
        }

        // Apply to used range only
        foreach (var column2 in Sheet.ColumnsUsed())
            column2.Width = FixedColumnWidth;

        foreach (var row2 in Sheet.RowsUsed())
            row2.Height = FixedRowHeight;
    }
}