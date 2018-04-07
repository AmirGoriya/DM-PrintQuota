using System;
using System.Diagnostics;
using MigraDoc.Rendering;

namespace PrintQuota
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create garbage input data which will be passed into the document printer.

                var summaryData = new[]
                {
                    // Creates some garbage data on the fly.
                    new { JobName = "Above Ground DMV", LabourHours = 20, LabourCost = 400, MaterialCost = 500},
                    new { JobName = "Blah", LabourHours = 24, LabourCost = 480, MaterialCost = 600},
                    new { JobName = "Garbage Stuff", LabourHours = 48, LabourCost = 960, MaterialCost = 900},
                    new { JobName = "Some Other Stuff", LabourHours = 91, LabourCost = 182, MaterialCost = 1800},
                    new { JobName = "And this and that", LabourHours = 62, LabourCost = 124, MaterialCost = 1100},
                    new { JobName = "I don't know what else", LabourHours = 55, LabourCost = 110, MaterialCost = 1000},
                    new { JobName = "This is getting boring", LabourHours = 44, LabourCost = 88, MaterialCost = 700},
                    new { JobName = "Someone help me please", LabourHours = 51, LabourCost = 102, MaterialCost = 1050},
                    new { JobName = "I'm really dying here", LabourHours = 57, LabourCost = 114, MaterialCost = 1125},
                    new { JobName = "Give me data", LabourHours = 35, LabourCost = 70, MaterialCost = 635},
                    new { JobName = "Anything will do", LabourHours = 12, LabourCost = 24, MaterialCost = 210},
                    new { JobName = "Okay sleepy time", LabourHours = 8, LabourCost = 16, MaterialCost = 22}
                };

                var sectionData = new[]
                {
                    // Creates some garbage data on the fly.
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500},
                    new { LabourHours = 10, MaterialDesc = "ABC", LabourUnitPrice = 1, LabourPrice = 400, MaterialUnitPrice = 5, MaterialPrice = 500}
                };

                var allSectionsData = new object[summaryData.Length];

                for (int i = 0; i < summaryData.Length; i++)
                {
                    allSectionsData[i] = sectionData;
                }

                // Initialize a quota object.
                var quota = new QuotaDocument();

                // Create the document using MigraDoc.
                var document = quota.CreateDocument();
                document.UseCmykColor = true;




                // Create a renderer for PDF that uses Unicode font encoding.
                var pdfRenderer = new PdfDocumentRenderer(true);

                // Set the MigraDoc document.
                pdfRenderer.Document = document;

                // Create the PDF document.
                pdfRenderer.RenderDocument();

                // Save the PDF document...
                var filePath = "Quota.pdf";
                pdfRenderer.Save(filePath);
                // ...and start a PDF viewer.
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }
    }
}
