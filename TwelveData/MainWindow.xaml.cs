using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Windows;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TwelveData
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private int count = 1;
        private readonly System.Timers.Timer timer = new System.Timers.Timer();
        private decimal firstNimber1;
        private decimal firstNimber2;
        private decimal firstNimber3;
        private decimal firstNimber4;
        private decimal firstNimber5;

        private decimal lastNumber1;
        private decimal lastNumber2;
        private decimal lastNumber3;
        private decimal lastNumber4;
        private decimal lastNumber5;


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Start_OnClick(object sender, RoutedEventArgs e)
        {
            CreateTable(null, null);
            timer.Interval = 60000;
            timer.Elapsed += CreateTable;
            timer.Start();
        }


        private void CreateTable(object sender, System.Timers.ElapsedEventArgs e)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://api.twelvedata.com/time_series?symbol=EUR/USD,USD/JPY,XPT/USD,JPM,ETH/USD,&interval=5min&apikey=d1f08feaaf364c9c813e284e69af04e0");

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream stream = response.GetResponseStream();
            StreamReader sr = new StreamReader(stream);
            string sReadData = sr.ReadToEnd();
            response.Close();

            JObject jObject = JObject.Parse(sReadData);
            Variables.symbol1 = (string)jObject["EUR/USD"]["meta"]["symbol"];
            Variables.symbol2 = (string)jObject["USD/JPY"]["meta"]["symbol"];
            Variables.symbol3 = (string)jObject["XPT/USD"]["meta"]["symbol"];
            Variables.symbol4 = (string)jObject["JPM"]["meta"]["symbol"];
            Variables.symbol5= (string)jObject["ETH/USD"]["meta"]["symbol"];

            Variables.currency1 = (string)jObject["EUR/USD"]["meta"]["currency_base"] + "/" + (string)jObject["EUR/USD"]["meta"]["currency_quote"];
            Variables.currency2 = (string)jObject["USD/JPY"]["meta"]["currency_base"] + "/" + (string)jObject["USD/JPY"]["meta"]["currency_quote"];
            Variables.currency3 = (string)jObject["XPT/USD"]["meta"]["currency_base"] + "/" + (string)jObject["XPT/USD"]["meta"]["currency_quote"];
            Variables.currency4 = (string)jObject["JPM"]["meta"]["currency"];
            Variables.currency5 = (string)jObject["ETH/USD"]["meta"]["currency_base"] + "/" + (string)jObject["ETH/USD"]["meta"]["currency_quote"];

            JArray jArray1 = (JArray)jObject["EUR/USD"]["values"];
            JArray jArray2 = (JArray)jObject["USD/JPY"]["values"];
            JArray jArray3 = (JArray)jObject["XPT/USD"]["values"];
            JArray jArray4 = (JArray)jObject["JPM"]["values"];
            JArray jArray5 = (JArray)jObject["ETH/USD"]["values"];

            Variables.value1 = (string)jObject["EUR/USD"]["values"][0]["close"];
            Variables.value2 = (string)jObject["USD/JPY"]["values"][0]["close"];
            Variables.value3 = (string)jObject["XPT/USD"]["values"][0]["close"];
            Variables.value4 = (string)jObject["JPM"]["values"][0]["close"];
            Variables.value5 = (string)jObject["ETH/USD"]["values"][0]["close"];

            var open1 = (string)jObject["EUR/USD"]["values"][0]["open"];
            var open2 = (string)jObject["USD/JPY"]["values"][0]["open"];
            var open3 = (string)jObject["XPT/USD"]["values"][0]["open"];
            var open4 = (string)jObject["JPM"]["values"][0]["open"];
            var open5 = (string)jObject["ETH/USD"]["values"][0]["open"];

            Decimal.TryParse(Variables.value1, NumberStyles.Any, new CultureInfo("en-US"), out firstNimber1);
            Decimal.TryParse(open1, NumberStyles.Any, new CultureInfo("en-US"), out lastNumber1);
            Variables.percent1 = (1 - lastNumber1 / firstNimber1) * 100;
            Decimal.TryParse(Variables.value2, NumberStyles.Any, new CultureInfo("en-US"), out firstNimber2);
            Decimal.TryParse(open2, NumberStyles.Any, new CultureInfo("en-US"), out lastNumber2);
            Variables.percent2 = (1 - lastNumber2 / firstNimber2) * 100;
            Decimal.TryParse(Variables.value3, NumberStyles.Any, new CultureInfo("en-US"), out firstNimber3);
            Decimal.TryParse(open3, NumberStyles.Any, new CultureInfo("en-US"), out lastNumber3);
            Variables.percent3 = (1 - lastNumber3 / firstNimber3) * 100;

            Decimal.TryParse(Variables.value4, NumberStyles.Any, new CultureInfo("en-US"), out firstNimber4);
            Decimal.TryParse(open4, NumberStyles.Any, new CultureInfo("en-US"), out lastNumber4);
            Variables.percent4 = (1 - lastNumber4 / firstNimber4) * 100;

            Decimal.TryParse(Variables.value5, NumberStyles.Any, new CultureInfo("en-US"), out firstNimber5);
            Decimal.TryParse(open5, NumberStyles.Any, new CultureInfo("en-US"), out lastNumber5);
            Variables.percent5 = (1 - lastNumber5 / firstNimber5) * 100;

            Variables.datetime1 = DateTime.Now.ToString();
            Variables.datetime2 = DateTime.Now.ToString();
            Variables.datetime3 = DateTime.Now.ToString();
            Variables.datetime4= DateTime.Now.ToString();
            Variables.datetime5 = DateTime.Now.ToString();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"D:\TwelveData.xlsx");
            IfFileExists(file);

        }

        private void Stop_OnClick(object sender, RoutedEventArgs e)
        {
            timer.Stop();
        }

        private async void IfFileExists(FileInfo file)
        {
            if (file.Exists)
            {
                var path = @"D:\TwelveData" + count + ".xlsx";
                file = new FileInfo(path);
                using var package = new ExcelPackage(file);
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Data");
                var range = ws.Cells["A1"];


                ws.Cells.AutoFitColumns(100, 100);
                ws.Cells["A1"].Value = "Twelve Data";
                ws.Cells["A1:C1"].Merge = true;
                ws.Row(1).Style.Font.Size = 24;
                ws.Row(1).Style.Font.Color.SetColor(Color.Blue);
                ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["A2"].Value = "Код";
                ws.Cells["B2"].Value = "Название";
                ws.Cells["C2"].Value = "Значение";
                ws.Cells["D2"].Value = "Изменение в процентах";
                ws.Cells["E2"].Value = "Время обновления";
                ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Row(2).Style.Font.Bold = true;

                ws.Column(2).Width = 50;
                ws.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(3).Width = 50;
                ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(4).Width = 50;
                ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(5).Width = 50;
                ws.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(6).Width = 50;
                ws.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(7).Width = 50;
                ws.Column(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["A3"].Value = Variables.symbol1;
                ws.Cells["B3"].Value = Variables.currency1;
                ws.Cells["C3"].Value = Variables.value1;
                ws.Cells["D3"].Value = Variables.percent1;
                ws.Cells["E3"].Value = Variables.datetime1;

                ws.Cells["A4"].Value = Variables.symbol2;
                ws.Cells["B4"].Value = Variables.currency2;
                ws.Cells["C4"].Value = Variables.value2;
                ws.Cells["D4"].Value = Variables.percent2;
                ws.Cells["E4"].Value = Variables.datetime2;

                ws.Cells["A5"].Value = Variables.symbol3;
                ws.Cells["B5"].Value = Variables.currency3;
                ws.Cells["C5"].Value = Variables.value3;
                ws.Cells["D5"].Value = Variables.percent3;
                ws.Cells["E5"].Value = Variables.datetime3;

                ws.Cells["A6"].Value = Variables.symbol4;
                ws.Cells["B6"].Value = Variables.currency4;
                ws.Cells["C6"].Value = Variables.value4;
                ws.Cells["D6"].Value = Variables.percent4;
                ws.Cells["E6"].Value = Variables.datetime4;

                ws.Cells["A7"].Value = Variables.symbol5;
                ws.Cells["B7"].Value = Variables.currency5;
                ws.Cells["C7"].Value = Variables.value5;
                ws.Cells["D7"].Value = Variables.percent5;
                ws.Cells["E7"].Value = Variables.datetime5;

                await package.SaveAsync();
                count++;
            }
            else
            {
                file = new FileInfo(@"D:\TwelveData.xlsx");
                using var package = new ExcelPackage(file);
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Data");
                var range = ws.Cells["A1"];

                ws.Cells["A1"].Value = "Twelve Data";
                ws.Cells["A1:E1"].Merge = true;
                ws.Row(1).Style.Font.Size = 24;
                ws.Row(1).Style.Font.Color.SetColor(Color.Blue);
                ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["A2"].Value = "Код";
                ws.Cells["B2"].Value = "Название";
                ws.Cells["C2"].Value = "Значение";
                ws.Cells["D2"].Value = "Изменение в процентах";
                ws.Cells["E2"].Value = "Время обновления";
                ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Row(2).Style.Font.Bold = true;

                ws.Column(2).Width = 50;
                ws.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(3).Width = 50;
                ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(4).Width = 50;
                ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(5).Width = 50;
                ws.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(6).Width = 50;
                ws.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(7).Width = 50;
                ws.Column(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["A3"].Value = Variables.symbol1;
                ws.Cells["B3"].Value = Variables.currency1;
                ws.Cells["C3"].Value = Variables.value1;
                ws.Cells["D3"].Value = Variables.percent1;
                ws.Cells["E3"].Value = Variables.datetime1;

                ws.Cells["A4"].Value = Variables.symbol2;
                ws.Cells["B4"].Value = Variables.currency2;
                ws.Cells["C4"].Value = Variables.value2;
                ws.Cells["D4"].Value = Variables.percent2;
                ws.Cells["E4"].Value = Variables.datetime2;

                ws.Cells["A5"].Value = Variables.symbol3;
                ws.Cells["B5"].Value = Variables.currency3;
                ws.Cells["C5"].Value = Variables.value3;
                ws.Cells["D5"].Value = Variables.percent3;
                ws.Cells["E5"].Value = Variables.datetime3;

                ws.Cells["A6"].Value = Variables.symbol4;
                ws.Cells["B6"].Value = Variables.currency4;
                ws.Cells["C6"].Value = Variables.value4;
                ws.Cells["D6"].Value = Variables.percent4;
                ws.Cells["E6"].Value = Variables.datetime4;

                ws.Cells["A7"].Value = Variables.symbol5;
                ws.Cells["B7"].Value = Variables.currency5;
                ws.Cells["C7"].Value = Variables.value5;
                ws.Cells["D7"].Value = Variables.percent5;
                ws.Cells["E7"].Value = Variables.datetime5;
                await package.SaveAsync();
            }
        }
    }
}
