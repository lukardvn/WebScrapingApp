using CsvHelper;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Configuration;
using System.Text.RegularExpressions;

namespace WebScrapingApp
{
    public class CurrencyExchangeRecord
    {
        public string CurrencyName { get; set; }
        public string BuyingRate { get; set; }
        public string CashBuyingRate { get; set; }
        public string SellingRate { get; set; }
        public string CashSellingRate { get; set; }
        public string MiddleRate { get; set; }
        public string PubTime { get; set; }

        public override string ToString()
        {
            return "Row: " + CurrencyName + ", " + BuyingRate + ", " + CashBuyingRate + ", " 
                + SellingRate + ", " + CashSellingRate + ", " + MiddleRate + ", " + PubTime;
        }
    }



    class Program
    {
        static HttpWebRequest RequestBuilder(string url, string method = "GET", string data = null, int timeout = 3)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = method.ToUpper();

            if (request.Method == "POST" && data != null) 
            {
                byte[] byteArray = Encoding.UTF8.GetBytes(data);
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = byteArray.Length;
                using (Stream dataStream = request.GetRequestStream())
                    dataStream.Write(byteArray, 0, byteArray.Length);
            }

            request.Timeout = timeout * 1000;

            Console.WriteLine("Created request: " + request.Method + " " + url + (data != null ? data : ""));

            return request;
        }

        static HttpWebResponse ProcessRequest(HttpWebRequest request)
        {
            int numberOfTries = 1;

            while (numberOfTries > 0)
            {
                try
                {
                    return (HttpWebResponse)request.GetResponse();
                } catch (WebException e)
                {
                    Console.WriteLine("Request timeout: " + request.Method + " " + request.RequestUri.ToString());
                    numberOfTries -= 1;
                }
            }

            return new HttpWebResponse();
            
        }

        static HtmlDocument GetHtmlResponse(HttpWebResponse response)
        {
            HtmlDocument document = new HtmlDocument();
            using (Stream dataStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                HtmlWeb web = new HtmlWeb();
                document.LoadHtml(responseFromServer);
            }
            response.Close();
            return document;
        }

        static void WriteToCsv(List<CurrencyExchangeRecord> rows, string fileName)
        {
            //load path for file output from App.config
            DirectoryInfo di = new DirectoryInfo(ConfigurationManager.AppSettings.Get("OutputDirectory"));
            //combine path from App.config with customed file name
            var file = Path.Combine(di.FullName, fileName + ".csv");

            using (var writer = new StreamWriter(file))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            //invariant culture-separating values with comma
            {
                Console.WriteLine("writing to file: " + file);
                csv.WriteRecords(rows);
            }
        }

        static List<string> GetAvailableCurrencies()
        {
            const string url = "https://srh.bankofchina.com/search/whpj/searchen.jsp";

            List<string> currencies = new List<string>();
            //first, request web page so we could initialize list of all available currencies from select options
            HttpWebRequest request = RequestBuilder(url);
            HttpWebResponse response = ProcessRequest(request);
            //Stream dataStream;
            HtmlDocument document = GetHtmlResponse(response);


            foreach (HtmlNode node in document.DocumentNode.SelectNodes("//select[@id='pjname']//option[position()>1]"))
                currencies.Add(node.Attributes["value"].Value);

            return currencies;
        }

        static string GenerateFileName(string currency, int nDaysAgo)
        {
            string endDate = DateTime.Now.ToString("yyyy-MM-dd");
            string startDate = DateTime.Parse(endDate).AddDays(-nDaysAgo).ToString("yyyy-MM-dd");
            return currency + "_" + startDate + "_" + endDate;
        }

        static List<CurrencyExchangeRecord> ParseDataFromTable(HtmlDocument document)
        {
            var table = document.DocumentNode.SelectSingleNode("//table[2]");
            //if table doesn't contain any data, skip current iteration and do not export that file
            if (table.SelectNodes("tr[position() > 1]") == null)
                return null;

            //if table does contain data, go through each row and add it to the list of rows
            List<CurrencyExchangeRecord> rows = new List<CurrencyExchangeRecord>();
            foreach (HtmlNode row in table.SelectNodes("tr[position() > 1]"))
            {
                var cells = row.SelectNodes("td");
                rows.Add(new CurrencyExchangeRecord
                {
                    CurrencyName = cells[0].InnerText.Trim(),
                    BuyingRate = cells[1].InnerText.Trim(),
                    CashBuyingRate = cells[2].InnerText.Trim(),
                    SellingRate = cells[3].InnerText.Trim(),
                    CashSellingRate = cells[4].InnerText.Trim(),
                    MiddleRate = cells[5].InnerText.Trim(),
                    PubTime = cells[6].InnerText.Trim()
                });
            }

            return rows;
        }

        static List<CurrencyExchangeRecord> LoadSinglePage(string currency, int pageNumber, int previousNDays)
        {
            const string url = "https://srh.bankofchina.com/search/whpj/searchen.jsp";
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string nDaysAgo = DateTime.Parse(today).AddDays(-previousNDays).ToString("yyyy-MM-dd");

            //send POST request containing from-date, to-date and currency 
            string postData = "erectDate=" + nDaysAgo + "&nothing=" + today + "&pjname=" + currency + "&page=" + pageNumber;
            var request = RequestBuilder(url, "POST", postData);

            //after sending post request with dates and currency, get the data from table and export it
            //response = (HttpWebResponse)request.GetResponse();
            var response = ProcessRequest(request);
            var document = GetHtmlResponse(response);

            var rows = ParseDataFromTable(document);
            return rows;
        }

        static List<CurrencyExchangeRecord> LoadAllPages(string currency, int previousNDays)
        {
            int numberOfPages = GetNumberOfPages(currency, previousNDays);
            if (numberOfPages < 1)
            {
                Console.WriteLine("Error loading pages for " + currency);
                return null;
            }
            Console.WriteLine("Found " + numberOfPages + " pages available for " + currency);

            List<CurrencyExchangeRecord> records = new List<CurrencyExchangeRecord>();
            for (int i = 0; i < numberOfPages; i++)
            {
                var pageNumber = i + 1;
                var rows = LoadSinglePage(currency, pageNumber, previousNDays);
                if (rows == null)
                {
                    Console.WriteLine("Unable to load rows on page " + pageNumber + " for " + currency);
                    continue;
                }
                Console.WriteLine("Downloaded " + rows.Count + " records from " + pageNumber + " page for " + currency);
                records.AddRange(rows);
            }

            Console.WriteLine("Total downloaded records for currency " + currency + ": " + records.Count);
            return records;
        }

        static int GetNumberOfPages(string currency, int previousNDays)
        {
            const string url = "https://srh.bankofchina.com/search/whpj/searchen.jsp";
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string nDaysAgo = DateTime.Parse(today).AddDays(-previousNDays).ToString("yyyy-MM-dd");

            //send POST request containing from-date, to-date and currency 
            string postData = "erectDate=" + nDaysAgo + "&nothing=" + today + "&pjname=" + currency;
            var request = RequestBuilder(url, "POST", postData);
            HttpWebResponse response = ProcessRequest(request);
            //Stream dataStream;
            //HtmlDocument document = GetHtmlResponse(response);

            // var scriptCode = document.DocumentNode.SelectSingleNode("/html/body/script[4]").InnerText;
            var pageSizeVariableName = "m_nPageSize";
            var recordCountVarName = "m_nRecordCount";
            /*var pattern = pageSizeVariableName + @"\s?=\s?(\d*?);";
            var matches = Regex.Matches(scriptCode, pattern);
            if (matches.Count > 0 && matches[0].Groups.Count > 1)
            {
                Console.WriteLine("usao");
                Console.WriteLine(matches[0].Groups[0].Value.ToString());
                Console.WriteLine(matches[0].Groups[1].Value.ToString());
            }*/

            //var partContainingPageNumber = scriptCode.Split(pageSizeVariableName)[1];

            string responseData = "";
            using (Stream dataStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(dataStream);
                responseData = reader.ReadToEnd();
            }

            var splitted = responseData.Split(pageSizeVariableName);
            if(splitted.Length < 2)
            {
                return -1;
            }
            var partContainingPageNumber = splitted[1];
            var indexOfSemicolon = partContainingPageNumber.IndexOf(';');
            var number = partContainingPageNumber.Substring(0, indexOfSemicolon);
            var numberClean = number.Replace("=", "").Trim();
            int pageSize = -1;
            if (int.TryParse(numberClean, out int ps))
            {
                pageSize = ps;
            } else
            {
                return -1;
            }


            splitted = splitted[0].Split(recordCountVarName);
            if (splitted.Length < 2)
            {
                return -1;
            }

            var partContainingRecordCount = splitted[1];
            indexOfSemicolon = partContainingRecordCount.IndexOf(';');
            number = partContainingRecordCount.Substring(0, indexOfSemicolon);
            numberClean = number.Replace("=", "").Trim();
            if (int.TryParse(numberClean, out int recordCount))
            {
                return (int)Math.Ceiling(recordCount / pageSize * 1.0);
            }
            return -1;
        }


        static void Main(string[] args)
        {

            int daysAgo = 2;
            List<string> currencies = GetAvailableCurrencies();
            //now, in local variable 'currencies' there's list of strings with all available currencies to select

            //loop through currencies and send post request for every currency, get the result and export it to csv
            foreach (var currency in currencies)
            {
                Console.WriteLine("Started scrapping " + currency);
                var data = LoadAllPages(currency, daysAgo);
                if (data == null)
                    continue;

                WriteToCsv(data, GenerateFileName(currency, daysAgo));
            }
        }
    }
}