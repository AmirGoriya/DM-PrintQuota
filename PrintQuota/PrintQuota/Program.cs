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
