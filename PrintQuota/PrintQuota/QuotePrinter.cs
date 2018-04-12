using System;
using System.Diagnostics;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;

namespace PrintQuote
{
    class QuotePrinter
    {
        /// <summary>
        /// The MigraDoc document that represents the Quote.
        /// </summary>
        Document _document;

        /// <summary>
        /// The text frame of the MigraDoc document that contains the quote name.
        /// </summary>
        TextFrame _quoteName;

        /// <summary>
        /// The table of the MigraDoc document that contains the Quote items.
        /// </summary>
        Table _table;

        /// <summary>
        /// The quote object which will hold all the data being passed to QuotePrinter.
        /// </summary>
        Quote _quote;

        // Some pre-defined colors in RGB.
        readonly static Color TableBorder = new Color(41, 111, 81);
        readonly static Color TableGreen = new Color(240, 249, 240);
        readonly static Color TableGray = new Color(242, 242, 242);

        /// <summary>
        /// Initializes an object which allows printing a quote to a PDF document.
        /// </summary>
        /// <param name="quote">An initialized Quote object.</param>
        public QuotePrinter(Quote quote)
        {
            // Assigns the passed quote object to quote field.
            _quote = quote;
        }

        /// <summary>
        /// Creates and saves a formatted PDF document containing 
        /// </summary>
        /// <param name="filePath"></param>
        public void PrintDocument(string filePath)
        {
            try
            {
                // Create the document using MigraDoc.
                var document = CreateDocument();
                document.UseCmykColor = true;

                // Create a renderer for PDF that uses Unicode font encoding.
                var pdfRenderer = new PdfDocumentRenderer(true);

                // Set the MigraDoc document.
                pdfRenderer.Document = document;

                // Create the PDF document.
                pdfRenderer.RenderDocument();

                // Save the PDF document...
                pdfRenderer.Save(filePath);
                // ...and start a PDF viewer.
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Creates the Quote printout document.
        /// </summary>
        public Document CreateDocument()
        {
            // Create a new MigraDoc document.
            _document = new Document();
            _document.Info.Title = "Quote Document";
            _document.Info.Subject = _quote.title;
            _document.Info.Author = "Dell Mechanical";


            DefineStyles();

            // Create and populate the summary page.
            CreateSummaryPage();
            PopulateSummaryPage();

            // Create and populate all section pages.
            for (int intSectionNumber = 0; intSectionNumber < _quote.sections.Count; intSectionNumber++)
            {
                CreateSectionPage(intSectionNumber);
                PopulateSectionPage(intSectionNumber);
            }


            return _document;
        }

        /// <summary>
        /// Defines the styles used to format the MigraDoc document.
        /// </summary>
        void DefineStyles()
        {
            // Get the predefined style Normal.
            var style = _document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Segoe UI";

            style = _document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("156cm", TabAlignment.Right);

            style = _document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called Table based on style Normal.
            style = _document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Segoe UI Semilight";
            style.Font.Size = 9;

            // Create a new style called Title based on style Normal.
            style = _document.Styles.AddStyle("Title", "Normal");
            style.Font.Name = "Segoe UI Semibold";
            style.Font.Size = 9;

            // Create a new style called Reference based on style Normal.
            style = _document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);
        }

        /// <summary>
        /// Creates the summary page template.
        /// </summary>
        void CreateSummaryPage()
        {
            // Each MigraDoc document needs at least one section.
            // Don't confuse the AddSection() with the tables we're defining as "sections" for this project.
            // In this context, "section" essentially means "page".
            var summaryPage = _document.AddSection();

            // Copy Default Page Setup and then change top and bottom margins.
            summaryPage.PageSetup = _document.DefaultPageSetup.Clone();
            summaryPage.PageSetup.TopMargin = "2cm";
            summaryPage.PageSetup.BottomMargin = "2cm";


            // If we choose to add a logo to the printouts, we can do that here.

            /***************************************************
            var image = section.Headers.Primary.AddImage(""); // Replace quotes with image address.
            image.Height = "3.5cm"; // This may need to be adjusted based on image size.
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Right;
            image.WrapFormat.Style = WrapStyle.Through;
            */

            // Create the footer.
            var paragraph = summaryPage.Footers.Primary.AddParagraph();
            paragraph.AddText("Dell Mechanical ● 666 Some Street ● Thunder Bay  ON ● (807) 999-9999");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Create the text frame for the address.
            // We can change this to hold any other information (for example: a client's information)...
            // ...as long as that data is being passed from the main application.
            _quoteName = summaryPage.AddTextFrame();
            _quoteName.Height = "3.0cm";
            _quoteName.Width = "7.0cm";
            _quoteName.Left = ShapePosition.Left;
            _quoteName.RelativeHorizontal = RelativeHorizontal.Margin;
            _quoteName.Top = "1.5cm";
            _quoteName.RelativeVertical = RelativeVertical.Page;

            // Show the sender in the address frame. (This is here in case we change the above address...
            // ...frame to show client's information as opposed to Dell Mechanical's info.)
            paragraph = _quoteName.AddParagraph("Dell Mechanical ● 666 Some Street ● (807) 999-9999");
            paragraph.Format.Font.Size = 7;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 3;

            // Add the print date field.
            paragraph = summaryPage.AddParagraph();
            // We use an empty paragraph to move the first text line below the address field.
            paragraph.Format.LineSpacing = "2.25cm";
            paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
            // And the table title.
            paragraph = summaryPage.AddParagraph();
            paragraph.Format.SpaceBefore = 0;
            paragraph.Style = "Reference";
            paragraph.AddFormattedText("QUOTE SUMMARY", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("Thunder Bay, ");
            paragraph.AddDateField("dd-MMMM-yyyy");

            // Create the table object.
            _table = summaryPage.AddTable();
            _table.Style = "Table";
            _table.Borders.Color = TableBorder;
            _table.Borders.Width = 0.25;
            _table.Borders.Left.Width = 0.5;
            _table.Borders.Right.Width = 0.5;
            _table.Rows.LeftIndent = 0;

            // Before you can add a row, you must define the columns.
            var column = _table.AddColumn("1.2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            // Create the header of the table.
            var row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableGreen;
            row.Cells[0].AddParagraph("Section No.");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("Job Name");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].MergeRight = 2;
            row.Cells[4].AddParagraph("Section Cost");
            row.Cells[4].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[4].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[4].MergeDown = 1;

            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableGreen;
            row.Cells[1].AddParagraph("Labour Hours");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].AddParagraph("Labour Cost");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("Material Cost");
            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;

            _table.SetEdge(0, 0, 5, 2, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);
        }

        /// <summary>
        /// Fills the summary page with data from the quote object.
        /// </summary>
        void PopulateSummaryPage()
        {
            // Fill the quote title in the text frame.
            var paragraph = _quoteName.AddParagraph();
            paragraph.AddText(_quote.title);

            // Create a counter for number of sections.
            int intSectionNumber = 0;

            // Iterate through the test data and transfer data to table.
            foreach (Section section in _quote.sections)
            {
                // Increment the section number.
                intSectionNumber++;

                // Each item fills two rows.
                var row1 = _table.AddRow();
                var row2 = _table.AddRow();
                row1.TopPadding = 1.5;
                row1.Cells[0].Shading.Color = TableGray;
                row1.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                row1.Cells[0].MergeDown = 1;
                row1.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                row1.Cells[1].MergeRight = 2;
                row1.Cells[4].Shading.Color = TableGray;
                row1.Cells[4].MergeDown = 1;

                row1.Cells[0].AddParagraph(intSectionNumber.ToString());
                paragraph = row1.Cells[1].AddParagraph();
                var formattedText = new FormattedText() { Style = "Title" };
                formattedText.AddText(section.title);
                paragraph.Add(formattedText);

                row2.Cells[1].AddParagraph($"{section.totalLabourHours.ToString()} hrs");
                row2.Cells[2].AddParagraph($"{section.totalLabourCost.ToString("c")}");
                row2.Cells[3].AddParagraph($"{section.totalMaterialCost.ToString("c")}");


                row1.Cells[4].AddParagraph($"{section.totalCost.ToString("c")}");
                row1.Cells[4].VerticalAlignment = VerticalAlignment.Bottom;

                _table.SetEdge(0, _table.Rows.Count - 2, 5, 2, Edge.Box, BorderStyle.Single, 0.75);
            }

            // Add an invisible row as a space line to the table.
            var row = _table.AddRow();
            row.Borders.Visible = false;

            // Add the total labour hours row.
            row = _table.AddRow();
            row.Cells[0].AddParagraph("Total Labour Hours");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 3;
            row.Cells[4].AddParagraph($"{_quote.totalLabourHours.ToString()}");
            row.Cells[4].Format.Font.Name = "Segoe UI";

            // Add the additional costs row.
            row = _table.AddRow();
            row.Cells[0].AddParagraph("Extra Costs");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 3;
            row.Cells[4].AddParagraph($"{_quote.extraCosts.ToString("c")}");
            row.Cells[4].Format.Font.Name = "Segoe UI";


            // Add the cost deductions row.
            row = _table.AddRow();
            row.Cells[0].AddParagraph("Cost Deductions");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 3;
            row.Cells[4].AddParagraph($"-({_quote.costDedeductions.ToString("c")})");
            row.Cells[4].Format.Font.Name = "Segoe UI";


            // Add the total cost row.
            row = _table.AddRow();
            row.Cells[0].AddParagraph("Total Cost");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 3;
            row.Cells[4].AddParagraph($"{_quote.totalCost.ToString("c")}");
            row.Cells[4].Format.Font.Name = "Segoe UI";
            row.Cells[4].Format.Font.Bold = true;

            // Set the borders of the specified cell range.
            _table.SetEdge(4, _table.Rows.Count - 4, 1, 4, Edge.Box, BorderStyle.Single, 0.75);

            // Add the comment box.
            paragraph = _document.LastSection.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.Format.SpaceBefore = "1cm";
            paragraph.Format.Borders.Width = 0.75;
            paragraph.Format.Borders.Distance = 3;
            paragraph.Format.Borders.Color = TableBorder;
            paragraph.Format.Shading.Color = TableGray;
            paragraph.AddText("Comments\n\n\n"); // Creates a big box for additional comments.
        }

        /// <summary>
        /// Creates a new Section Page.
        /// </summary>
        void CreateSectionPage(int intSectionNumber)
        {
            var sectionPage = _document.AddSection();

            // Define the page setup. 
            sectionPage.PageSetup = _document.DefaultPageSetup.Clone();
            sectionPage.PageSetup.TopMargin = "1.25cm";
            sectionPage.PageSetup.BottomMargin = "1.25cm";
            sectionPage.PageSetup.LeftMargin = "4cm";
            sectionPage.PageSetup.RightMargin = "4cm";


            // Create the footer.
            var paragraph = sectionPage.Footers.Primary.AddParagraph();
            paragraph.AddText("Dell Mechanical ● 666 Some Street ● Thunder Bay  ON ● (807) 999-9999");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Add the print date field.
            paragraph = sectionPage.AddParagraph();
            // We use an empty paragraph to move the first text line below the address field.
            paragraph.Format.LineSpacing = "2.25cm";
            paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
            // Add the table title.
            paragraph = sectionPage.AddParagraph();
            paragraph.Format.SpaceBefore = 0;
            paragraph.Style = "Reference";
            paragraph.AddFormattedText(_quote.sections[intSectionNumber].title, TextFormat.Underline);
            //**** Need to make it so that the above title is dynamic and changes based on JOB NAME.****//

            // Create the table object.
            _table = sectionPage.AddTable();
            _table.Style = "Table";
            _table.Borders.Color = TableBorder;
            _table.Borders.Width = 0.25;
            _table.Borders.Left.Width = 0.5;
            _table.Borders.Right.Width = 0.5;
            _table.Rows.LeftIndent = 0;

            // Before you can add a row, you must define the columns.
            var column = _table.AddColumn("1.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            // Create the header of the table.
            var row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableGreen;
            row.Cells[0].AddParagraph("Quantity");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("Material Type");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].MergeRight = 3;

            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableGreen;
            row.Cells[1].AddParagraph("Labour Unit Price");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].AddParagraph("Labour Price");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("Material Unit Price");
            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[4].AddParagraph("Material Price");
            row.Cells[4].Format.Alignment = ParagraphAlignment.Left;

            _table.SetEdge(0, 0, 5, 2, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);
        }

        /// <summary>
        /// Fills the section page with dynamic data.
        /// </summary>
        void PopulateSectionPage(int intSectionNumber)
        {          

            // Iterate through each entry in the current section.
            for (int intEntryNumber = 0; intEntryNumber < _quote.sections[intSectionNumber]._materialTypes.Count; intEntryNumber++)
            {
                // Each item fills two rows.
                var row1 = _table.AddRow();
                var row2 = _table.AddRow();
                row1.TopPadding = 1.5;
                row1.Cells[0].Shading.Color = TableGray;
                row1.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                row1.Cells[0].MergeDown = 1;
                row1.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                row1.Cells[1].MergeRight = 3;

                row1.Cells[0].AddParagraph(_quote.sections[intSectionNumber]._quantity[intEntryNumber].ToString());
                var paragraph = row1.Cells[1].AddParagraph();
                var formattedText = new FormattedText() { Style = "Title" };
                formattedText.AddText(_quote.sections[intSectionNumber].materialTypes[intEntryNumber]);
                paragraph.Add(formattedText);

                // Fill all the columns with unit costs and true costs for each entry in the section page.
                // Relies on all section lists having synchronized entries.
                row2.Cells[1].AddParagraph($"{_quote.sections[intSectionNumber].labourUnitCosts[intEntryNumber].ToString("c")}/unit");
                row2.Cells[2].AddParagraph(_quote.sections[intSectionNumber].labourCosts[intEntryNumber].ToString("c"));
                row2.Cells[3].AddParagraph($"{_quote.sections[intSectionNumber].materialUnitCosts[intEntryNumber].ToString("c")}/unit");
                row2.Cells[4].AddParagraph(_quote.sections[intSectionNumber].materialCosts[intEntryNumber].ToString("c"));


                _table.SetEdge(0, _table.Rows.Count - 2, 5, 2, Edge.Box, BorderStyle.Single, 0.75);
            }


            // Add the labour total columns.
            var row = _table.AddRow();
            row.Cells[0].AddParagraph("Labour Total");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 1;
            row.Cells[2].AddParagraph(_quote.sections[intSectionNumber].totalLabourCost.ToString("c"));
            row.Cells[2].Format.Font.Name = "Segoe UI";

            // Add the material total columns.
            row.Cells[3].AddParagraph("Materials Total");
            row.Cells[3].Borders.Visible = false;
            row.Cells[3].Format.Font.Bold = true;
            row.Cells[3].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[4].AddParagraph(_quote.sections[intSectionNumber].totalMaterialCost.ToString("c"));
            row.Cells[4].Format.Font.Name = "Segoe UI";

            // Set the borders of the specified cell range.
            _table.SetEdge(2, _table.Rows.Count - 1, 1, 1, Edge.Box, BorderStyle.Single, 0.75);
            _table.SetEdge(4, _table.Rows.Count - 1, 1, 1, Edge.Box, BorderStyle.Single, 0.75);
        }
    }
}
