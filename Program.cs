using System;
using static elliottjobscraper.scrape;

namespace elliottjobscraper
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Elliot Electric job scraper");
            Console.WriteLine("Please select an option:");
            Console.WriteLine("1: Scrape the entire job table and write to Excel");
            Console.WriteLine("2: Scrape just the job links and write to Excel");
            Console.WriteLine("Enter a number and press Enter:");

            int userChoice;
            while (true)
            {
                if (int.TryParse(Console.ReadLine(), out userChoice))
                {
                    if (userChoice == 1)
                    {
                        ScrapeAll();
                        break;
                    }
                    else if (userChoice == 2)
                    {
                        ScrapeLinks();
                        break;
                    }
                }

                Console.WriteLine("Invalid choice. Please enter 1 or 2 and press Enter:");
            }

            Console.WriteLine("Done! Press any key to exit.");
            Console.ReadKey();
        }
    }
}