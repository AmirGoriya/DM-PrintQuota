using System.Collections.Generic;
using System.Diagnostics;
namespace PrintQuote
{
    class Program
    {
        static void Main(string[] args)
        {
            Quote quote = new Quote("Test Quote Document");
            quote.sections = new List<Section>();

            for (int i = 0; i < 5; i++)
            {
                quote.sections.Add(new Section
                {
                    title = $"Job #{i}",
                    materialTypes = new List<string> { "A", "B", "C", "D" },
                    quantity = new List<int> { 1, 2, 3, 4 },
                    materialCosts = new List<double> { 10, 20, 30, 40 },
                    materialUnitCosts = new List<double> { 10, 20, 30, 40 },
                    labourUnitCosts = new List<double> { 5, 10, 15, 20 },
                    labourCosts = new List<double> { 5, 10, 15, 20 }
                });
                quote.sections[i].calc_sectionTotals();
            }
            quote.costDedeductions = 10;
            quote.calcTotals();


            /* Amir's addition*****************/
            string filePath = "quote.pdf"; // Replace with the location of where the PDF will be saved.

            // Initialize a QuotePrinter object.
            var quotePrinter = new QuotePrinter(quote);

            // Calls the quotePrinter's print method.
            // This method creates a document using the data sent in.
            quotePrinter.PrintDocument(filePath);

            // Opens the PDF document using the default PDF viewer.           
            Process.Start(filePath);

            /*********************************/

        }
    }
}
