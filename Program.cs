using System;
using HtmlAgilityPack;
using System.IO;
using System.Net;
using System.Text;
using ExcelDataReader;
using System.Data;
using System.Globalization;

class Program
{
    static void Main(string[] args)
    {
        string url = "https://www.abs.gov.au/statistics/labour/employment-and-unemployment/labour-force-australia";
        string prefix = "https://www.abs.gov.au";
        HtmlWeb web = new HtmlWeb();
        HtmlDocument doc = web.Load(url);

        var latestReleaseNode = doc.DocumentNode.SelectSingleNode("//h2[contains(text(), 'Latest release')]/following::a[1]");

        if (latestReleaseNode != null)
        {
            string latestReleaseLink = latestReleaseNode.GetAttributeValue("href", "");
            Console.WriteLine("Latest release link: " + latestReleaseLink);
            string url2 = prefix + latestReleaseLink;

            HtmlWeb web2 = new HtmlWeb();
            HtmlDocument doc2 = web2.Load(url2);

            HtmlNode xlsxNode = doc2.DocumentNode.SelectSingleNode("//h3[text()='Labour Force status']/following::a[contains(@href, '.xlsx')][1]");

            if (xlsxNode != null)
            {
                string xlsxUrl = prefix + xlsxNode.GetAttributeValue("href", "");
                Console.WriteLine("Found the first .xlsx URL after 'Labour Force status': " + xlsxUrl);

                WebClient client = new WebClient();
                //Had an issue of System.Net.WebException: The remote server returned an error: (403) Forbidden.
                //The website probably not allow web automation scraping
                client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3");

                string downloadsFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Downloads");
                Directory.CreateDirectory(downloadsFolderPath);

                string fileName = Path.Combine(downloadsFolderPath, "data.xlsx");
                client.DownloadFile(xlsxUrl, fileName);

                Console.WriteLine("Excel file downloaded to: " + fileName);
                string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Downloads", "data.xlsx");

                using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    //Explicitly set the encoding to UTF-8
                    //Had an issue of System.NotSupportedException
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                    var encoding = Encoding.GetEncoding(1252);

                    using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = encoding }))
                    {
                        DataSet result = reader.AsDataSet();

                        DataTable data = null;
                        foreach (DataTable table in result.Tables)
                        {
                            if (table.TableName == "Data1")
                            {
                                data = table;
                                break;
                            }
                        }

                        if (data == null)
                        {
                            Console.WriteLine("Worksheet 'Data1' not found in the Excel file.");
                            return;
                        }

                        //The below foreach loop is to ensure the date to be MMM-yyyy
                        foreach (DataRow row in data.Rows)
                        {
                            if (DateTime.TryParse(row[0].ToString(), out DateTime dateValue))
                            {
                                row[0] = dateValue.ToString("MMM-yyyy", CultureInfo.InvariantCulture);
                            }
                        }

                        int seriesIdRow = -1;

                        for (int i = 0; i < data.Rows.Count; i++)
                        {
                            if (data.Rows[i][0].ToString() == "Series ID")
                            {
                                seriesIdRow = i;
                                break;
                            }
                        }

                        if (seriesIdRow == -1)
                        {
                            Console.WriteLine("Series ID not found in the Excel sheet.");
                            return;
                        }

                        StringBuilder csvData = new StringBuilder();
                        for (int j = 0; j < data.Columns.Count; j++)
                        {
                            for (int i = seriesIdRow; i < data.Rows.Count; i++)
                            {
                                csvData.Append(data.Rows[i][j]);
                                if (i < data.Rows.Count - 1)
                                {
                                    csvData.Append(",");
                                }
                            }
                            csvData.AppendLine();
                        }

                        string csvFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Downloads", "result.csv");
                        File.WriteAllText(csvFilePath, csvData.ToString());

                        Console.WriteLine("Data transposed and saved to " + csvFilePath);
                    }
                }
            }
            else
            {
                Console.WriteLine("No .xlsx URL found after 'Labour Force status'.");
            }
        }
        else
        {
            Console.WriteLine("Latest release link not found.");
        }
    }
}