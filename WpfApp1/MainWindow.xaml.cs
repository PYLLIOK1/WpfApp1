﻿using Microsoft.Win32;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Windows;
using System.IO;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Net;
using HtmlAgilityPack;
using System.Text;
using CefSharp;
using System.Linq;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        public const string main = "https://vk.com";
        //private List<Parse> Liste = new List<Parse>();
        public List<Parse> Pars(OpenFileDialog dialog)
        {
            HSSFWorkbook hssfwb;
            List<Parse> Liste = new List<Parse>();
            using (FileStream file = new FileStream(dialog.FileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheetAt(hssfwb.ActiveSheetIndex);
            if (sheet != null)
            {
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row != null && row.Cells.Count == 5 && row.Cells[4].CellType == CellType.String && !string.IsNullOrWhiteSpace(row.Cells[4].StringCellValue))
                    {
                        Liste.Add(new Parse { Name = row.Cells[1].StringCellValue, Price = row.Cells[4].StringCellValue.Replace("-", ",") });
                    }
                }
            }
            return Liste;
        }// парсинг таблицы
        private void AddFile(object sender, RoutedEventArgs e)
        {
            List<Parse> Liste = new List<Parse>();
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == true)
            {
                Liste = Pars(dialog);
                Add(Liste);
            }
            
        }//выбор таблицы

        public void Add(List<Parse> Liste)
        {
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            String spreadsheetId = "1IUrDULP4pzI8kioVAcl_5f1-B8mFdaahjuj8qwVe5dA";
            string AppName = "WpfApp1";
            
            using (var stream =
                 new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                UserCredential credential;
                string credPath = Environment.GetFolderPath(
                    Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = AppName
                });
                SpreadsheetsResource.GetRequest request = service.Spreadsheets.Get(spreadsheetId);
                var response = request.Execute();
                var sheet = response.Sheets.FirstOrDefault(x => Convert.ToBoolean(x.Properties.GridProperties.RowCount));
                List<Request> requests = new List<Request>();
                BatchUpdateSpreadsheetRequest busrer = new BatchUpdateSpreadsheetRequest
                {
                    Requests = requests
                };
                for (int i = 4; i < sheet.Properties.GridProperties.RowCount; i++)
                {
                    List<CellData> Valurrr = new List<CellData>
                    {
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = "" } },
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = "" } },
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = "" } }
                    };
                    requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = 0, RowIndex = i, ColumnIndex = 0 }, Rows = new List<RowData> { new RowData { Values = Valurrr } }, Fields = "userEnteredValue" } });
                }
                for (int i = 4; i < Liste.Count; i++)
                {
                    List<CellData> Value = new List<CellData>
                    {
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = Liste[i].Name, } },
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = Liste[i].Price } },
                        new CellData { UserEnteredValue = new ExtendedValue { FormulaValue = "=SUM(D" + (i + 5) + ":" + (i + 5) + ")" } }
                    };
                    requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = 0, RowIndex = i, ColumnIndex = 0 }, Rows = new List<RowData> { new RowData { Values = Value } }, Fields = "userEnteredValue" } });
                }
                List<CellData> Cash = new List<CellData>();
                for (char c = 'D'; c <= 'Z'; c++)
                {
                    Cash.Add(new CellData { UserEnteredValue = new ExtendedValue { FormulaValue = "=ROUNDUP(SUMPRODUCT($B5:$B;" + c + "5:" + c + ");1)" } });
                }
                requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = 0, RowIndex = 1, ColumnIndex = 3 }, Rows = new List<RowData> { new RowData { Values = Cash } }, Fields = "userEnteredValue" } });
                List<CellData> Dates = new List<CellData>
            {
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = DateTime.Today.ToShortDateString() } }
            };
                requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = 0, RowIndex = 0, ColumnIndex = 0 }, Rows = new List<RowData> { new RowData { Values = Dates } }, Fields = "userEnteredValue" } });
                List<CellData> Sum = new List<CellData>
            {
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = "Сумма" } },
            };
                requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = 0, RowIndex = 1, ColumnIndex = 1 }, Rows = new List<RowData> { new RowData { Values = Sum } }, Fields = "userEnteredValue" } });
                List<CellData> Info = new List<CellData>
            {
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = "Дополнительная информация" } },
            };
                requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = 0, RowIndex = 2, ColumnIndex = 0 }, Rows = new List<RowData> { new RowData { Values = Info } }, Fields = "userEnteredValue" } });
                List<CellData> Name = new List<CellData>
            {
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = "Наименование" } },
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = "Цена" } },
            };
                requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = 0, RowIndex = 3, ColumnIndex = 0 }, Rows = new List<RowData> { new RowData { Values = Name } }, Fields = "userEnteredValue" } });
                List<CellData> SumFormula = new List<CellData>
            {
                new CellData { UserEnteredValue = new ExtendedValue { FormulaValue = "=SUM(D2:2)" } }
            };
                service.Spreadsheets.BatchUpdate(busrer, spreadsheetId).Execute();
                MessageBox.Show("Выполнено");
            }
        }

        private void DownloadMenu(object sender, RoutedEventArgs e)
        {
            bool flag = false;
            string xls = "xls";
            WebClient client = new WebClient() { Encoding = Encoding.UTF8 };
            string s = client.DownloadString("https://vk.com/lunch_vesta");
            client.Dispose();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(s);
            HtmlNodeCollection c = doc.DocumentNode.SelectNodes("//div[@class='medias_row']");
            if (c != null)
            {
                foreach (HtmlNode n in c)
                {
                    if ((n.InnerHtml.Contains(xls)) && (n.InnerText.Contains(DateTime.Today.ToShortDateString())))
                    {
                        string text = n.InnerHtml;
                        HtmlDocument docc = new HtmlDocument();
                        docc.LoadHtml(text);
                        HtmlNode a = docc.DocumentNode.SelectSingleNode("//a[@class='mr_label medias_link']");
                        {
                            if (a.Attributes["href"] != null)
                            {
                                s = main + a.Attributes["href"].Value;
                                browser.FrameLoadEnd += Browser_FrameLoadEnd;
                                browser.Address = s;
                                flag = true;
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("не удалось найти");
            }
            if (!flag)
            {
                MessageBox.Show("не удалось найти");
            }

        } // парсинг вк группы 

        private void Browser_FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        {
            if (e.Frame.IsMain)
            {
                browser.DownloadHandler = new DownloadHandler();
                browser.ExecuteScriptAsync("saveDoc();");
                //Таблицу сохраняет где находится исполняемый файл
            }
        }
    }
    public class DownloadHandler : IDownloadHandler
    {

        public event EventHandler<DownloadItem> OnBeforeDownloadFired;

        public event EventHandler<DownloadItem> OnDownloadUpdatedFired;


        public void OnBeforeDownload(IBrowser browser, DownloadItem downloadItem, IBeforeDownloadCallback callback)
        {
            OnBeforeDownloadFired?.Invoke(this, downloadItem);

            if (!callback.IsDisposed)
            {
                using (callback)
                {
                    callback.Continue(downloadItem.SuggestedFileName, showDialog: false);
                }
            }
        }

        public void OnDownloadUpdated(IBrowser browser, DownloadItem downloadItem, IDownloadItemCallback callback)
        {
            OnDownloadUpdatedFired?.Invoke(this, downloadItem);
            if (downloadItem.IsComplete)
            {
                List<Parse> Liste = new List<Parse>();
                OpenFileDialog file = new OpenFileDialog
                {
                    FileName = downloadItem.FullPath
                };
                MainWindow main = new MainWindow();
                Liste = main.Pars(file);
                main.Add(Liste);
                File.Delete(downloadItem.FullPath);
            }
        }
    }
    public class Parse
    {
        public string Name { get; set; }
        public string Price { get; set; }
    }
}