using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using HtmlAgilityPack;
using OfficeOpenXml;

namespace elliottjobscraper
{
    class Scrape
    {

        public static void ScrapeAll()
        {
            // Define the URL 
            string url = "https://www.elliottelectric.com/Kiosk/main.aspx";

            // Create an HttpClient object to make the web request
            var httpClient = new HttpClient();

            // Send the request and get the response
            var response = httpClient.GetAsync(url).Result;

            // Read the content of the response as a string
            var content = response.Content.ReadAsStringAsync().Result;

            // Parse the content with HtmlAgilityPack
            var doc = new HtmlDocument();
            doc.LoadHtml(content);

            // Find the table on the page that contains the job listings... its the first one
            var table = doc.DocumentNode.Descendants("table").FirstOrDefault(t => t.GetAttributeValue("id", "") == "jobs2");

            if (table != null)
            {
                // Create a new Excel workbook
                var excelPackage = new ExcelPackage();

                // Add a new worksheet to the workbook
                var worksheet = excelPackage.Workbook.Worksheets.Add("Job Openings");

                // Loop through the rows in the table
                var rows = table.Descendants("tr");
                for (int i = 0; i < rows.Count(); i++)
                {
                    var row = rows.ElementAt(i);

                    // Loop through the cells in the row
                    var cells = row.Descendants("td");
                    for (int j = 0; j < cells.Count(); j++)
                    {
                        var cell = cells.ElementAt(j);

                        // Write the cell value to the corresponding cell
                        worksheet.Cells[i + 1, j + 1].Value = cell.InnerText.Trim();
                    }
                }

                // Save the workbook as job-openings-all.xlsx
                var fileInfo = new System.IO.FileInfo("job-openings-all.xlsx");
                excelPackage.SaveAs(fileInfo);

                // Display the path where the Excel file was saved
                Console.WriteLine("Excel file saved to: " + fileInfo.FullName);
            }
            else
            {
                Console.WriteLine("Error: Jobs table not found on page.");
            }
        }
        public static void ScrapeLinks()
        {
            // Define the URL 
            string url = "https://www.elliottelectric.com/Kiosk/main.aspx";

            // Create an HttpClient object to make the web request
            var httpClient = new HttpClient();

            // Send the request and get the response
            var response = httpClient.GetAsync(url).Result;

            // Read the content of the response as a string
            var content = response.Content.ReadAsStringAsync().Result;

            // Parse the content with HtmlAgilityPack
            var doc = new HtmlDocument();
            doc.LoadHtml(content);

            // Find the table on the page that contains the job links... its the first one
            var table = doc.DocumentNode.Descendants("table").FirstOrDefault(t => t.GetAttributeValue("id", "") == "jobs2");

            if (table != null)
            {
                // Find all the job listings in the table
                var listings = table.Descendants("tr").Skip(2);

                // Create a list to store the job links
                var jobLinks = new List<string>();

                // Loop through the job listings and extract the link for each job
                foreach (var listing in listings)
                {
                    //three links in each row... skip first one to get desired
                    var relativeLink = listing.Descendants("a").Skip(1).FirstOrDefault()?.GetAttributeValue("href", "");

                    if (!string.IsNullOrEmpty(relativeLink))
                    {
                        var absoluteLink = new Uri(new Uri(url), relativeLink).ToString();
                        jobLinks.Add(absoluteLink);
                    }
                    else
                    {
                        Console.WriteLine("Error: Job listing does not have a link.");
                    }
                }

                if (jobLinks.Count > 0)
                {
                    
                    var excelPackage = new ExcelPackage();

                    
                    var worksheet = excelPackage.Workbook.Worksheets.Add("Job Openings");

                    
                    for (int i = 0; i < jobLinks.Count; i++)
                    {
                        worksheet.Cells[i + 1, 1].Value = jobLinks[i];
                    }

                    // Save the workbook to job-openings-links.xlsx
                    var fileInfo = new System.IO.FileInfo("job-openings-links.xlsx");
                    excelPackage.SaveAs(fileInfo);

                    
                    Console.WriteLine("Excel file saved to: " + fileInfo.FullName);
                }
                else
                {
                    Console.WriteLine("Error: No job links found on page.");
                }
            }
            else
            {
                Console.WriteLine("Error: Jobs table not found on page.");
            }
        }
    }
}