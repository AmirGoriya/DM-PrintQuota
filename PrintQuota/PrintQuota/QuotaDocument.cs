using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;

namespace PrintQuota
{
    class QuotaDocument
    {
        /// <summary>
        /// The MigraDoc document that represents the quota.
        /// </summary>
        Document _document;

        /// <summary>
        /// The text frame of the MigraDoc document that contains the address.
        /// </summary>
        TextFrame _addressFrame;

        /// <summary>
        /// The table of the MigraDoc document that contains the quota items.
        /// </summary>
        Table _table;

        // Some pre-defined colors in RGB.
        readonly static Color TableBorder = new Color(41, 111, 81);
        readonly static Color TableGreen = new Color(240, 249, 240);
        readonly static Color TableGray = new Color(242, 242, 242);

        /// <summary>
        /// Creates the quota printout document.
        /// </summary>
        public Document CreateDocument()
        {
            // Create a new MigraDoc document.
            _document = new Document();
            _document.Info.Title = "A sample quota";
            _document.Info.Subject = "Experimenting with MigraDoc.";
            _document.Info.Author = "Amir Goriya";

            DefineStyles();

            CreateSummaryPage();

            PopulateSummaryPage();

            CreateSectionPage();

            PopulateSectionPage();

            CreateSectionPage();

            PopulateSectionPage();

            CreateSectionPage();

            PopulateSectionPage();

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
            var summary = _document.AddSection();

            // Copy Default Page Setup and then change top and bottom margins.
            summary.PageSetup = _document.DefaultPageSetup.Clone();
            summary.PageSetup.TopMargin = "2cm";
            summary.PageSetup.BottomMargin = "2cm";


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
            var paragraph = summary.Footers.Primary.AddParagraph();
            paragraph.AddText("Dell Mechanical ● 666 Some Street ● Thunder Bay  ON ● (807) 999-9999");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Create the text frame for the address.
            // We can change this to hold any other information (for example: a client's information)...
            // ...as long as that data is being passed from the main application.
            _addressFrame = summary.AddTextFrame();
            _addressFrame.Height = "3.0cm";
            _addressFrame.Width = "7.0cm";
            _addressFrame.Left = ShapePosition.Left;
            _addressFrame.RelativeHorizontal = RelativeHorizontal.Margin;
            _addressFrame.Top = "1.5cm";
            _addressFrame.RelativeVertical = RelativeVertical.Page;

            // Show the sender in the address frame. (This is here in case we change the above address...
            // ...frame to show client's information as opposed to Dell Mechanical's info.)
            paragraph = _addressFrame.AddParagraph("Dell Mechanical ● 666 Some Street ● (807) 999-9999");
            paragraph.Format.Font.Size = 7;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 3;

            // Add the print date field.
            paragraph = summary.AddParagraph();
            // We use an empty paragraph to move the first text line below the address field.
            paragraph.Format.LineSpacing = "2.25cm";
            paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
            // And the table title.
            paragraph = summary.AddParagraph();
            paragraph.Format.SpaceBefore = 0;
            paragraph.Style = "Reference";
            paragraph.AddFormattedText("QUOTA SUMMARY", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("Thunder Bay, ");
            paragraph.AddDateField("dd-MMMM-yyyy");

            // Create the table object.
            _table = summary.AddTable();
            _table.Style = "Table";
            _table.Borders.Color = TableBorder;
            _table.Borders.Width = 0.25;
            _table.Borders.Left.Width = 0.5;
            _table.Borders.Right.Width = 0.5;
            _table.Rows.LeftIndent = 0;

            // Before you can add a row, you must define the columns.
            var column = _table.AddColumn("1.2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = _table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            // Create the header of the table.
            var row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableGreen;
            row.Cells[0].AddParagraph("Section");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("Job Name");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].MergeRight = 3;
            row.Cells[5].AddParagraph("Section Cost");
            row.Cells[5].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[5].MergeDown = 1;

            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableGreen;
            row.Cells[1].AddParagraph("Labour Hours");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].AddParagraph("Labour Cost");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("Material");
            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[4].AddParagraph("Extra Cost");
            row.Cells[4].Format.Alignment = ParagraphAlignment.Left;

            _table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);
        }

        /// <summary>
        /// Fills the summary page with dynamic data.
        /// </summary>
        void PopulateSummaryPage()
        {
            // Fill the address in the address text frame.
            // This is where we can receive the client info which can go in the addressframe.
            var paragraph = _addressFrame.AddParagraph();
            paragraph.AddText("Dell Mechanical\n");
            paragraph.AddText("666 Some Street\n");
            paragraph.AddText("Thunder Bay ON  P7B 16A\n");
            paragraph.AddText("(807) 999-999");

            // Initialize total and extra costs and the section number.
            double totalCost = 0;
            double extraCosts = 55;
            int sectionNumber = 0;

            // Create an array of test data.
            var inputData = new[]
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

            // Iterate through the test data and transfer data to table.
            foreach (var input in inputData)
            {
                sectionNumber++;

                // Each item fills two rows.
                var row1 = _table.AddRow();
                var row2 = _table.AddRow();
                row1.TopPadding = 1.5;
                row1.Cells[0].Shading.Color = TableGray;
                row1.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                row1.Cells[0].MergeDown = 1;
                row1.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                row1.Cells[1].MergeRight = 3;
                row1.Cells[5].Shading.Color = TableGray;
                row1.Cells[5].MergeDown = 1;

                row1.Cells[0].AddParagraph(sectionNumber.ToString());
                paragraph = row1.Cells[1].AddParagraph();
                var formattedText = new FormattedText() { Style = "Title" };
                formattedText.AddText(input.JobName);
                paragraph.Add(formattedText);

                row2.Cells[1].AddParagraph($"{input.LabourHours.ToString()} hrs");
                row2.Cells[2].AddParagraph($"{input.LabourCost.ToString("c")}");
                row2.Cells[3].AddParagraph($"{input.MaterialCost.ToString("c")}");
                row2.Cells[4].AddParagraph();

                double sectionCost = (input.LabourHours * input.LabourCost) + input.MaterialCost;
                row1.Cells[5].AddParagraph($"{sectionCost.ToString("c")}");
                row1.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
                totalCost += sectionCost;

                _table.SetEdge(0, _table.Rows.Count - 2, 6, 2, Edge.Box, BorderStyle.Single, 0.75);
            }

            // Add an invisible row as a space line to the table.
            var row = _table.AddRow();
            row.Borders.Visible = false;

            // Add the additional costs row.
            row = _table.AddRow();
            row.Cells[0].AddParagraph("Additional Costs");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            row.Cells[5].AddParagraph($"{extraCosts.ToString("c")}");
            row.Cells[5].Format.Font.Name = "Segoe UI";

            totalCost += extraCosts;

            // Add the total cost row.
            row = _table.AddRow();
            row.Cells[0].AddParagraph("Total Cost");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            row.Cells[5].AddParagraph($"{totalCost.ToString("c")}");
            row.Cells[5].Format.Font.Name = "Segoe UI";
            row.Cells[5].Format.Font.Bold = true;

            // Set the borders of the specified cell range.
            _table.SetEdge(5, _table.Rows.Count - 4, 1, 4, Edge.Box, BorderStyle.Single, 0.75);

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
        void CreateSectionPage()
        {
            var section = _document.AddSection();

            // Define the page setup. 
            section.PageSetup = _document.DefaultPageSetup.Clone();
            section.PageSetup.TopMargin = "1.25cm";
            section.PageSetup.BottomMargin = "1.25cm";
            section.PageSetup.LeftMargin = "4cm";
            section.PageSetup.RightMargin = "4cm";


            // Create the footer.
            var paragraph = section.Footers.Primary.AddParagraph();
            paragraph.AddText("Dell Mechanical ● 666 Some Street ● Thunder Bay  ON ● (807) 999-9999");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Add the print date field.
            paragraph = section.AddParagraph();
            // We use an empty paragraph to move the first text line below the address field.
            paragraph.Format.LineSpacing = "2.25cm";
            paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
            // Add the table title.
            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = 0;
            paragraph.Style = "Reference";
            paragraph.AddFormattedText("JOB NAME", TextFormat.Underline);
            //**** Need to make it so that the above title is dynamic and changes based on JOB NAME.****//

            // Create the table object.
            _table = section.AddTable();
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
            row.Cells[0].AddParagraph("Labour Hours");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("Material");
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
        void PopulateSectionPage()
        {
            // Initialize total and extra costs and the section number.
            double labourTotal = 0;
            double materialTotal = 0;

            // Create an array of test data.
            var inputData = new[]
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

            // Iterate through the test data and transfer data to table.
            foreach (var input in inputData)
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

                row1.Cells[0].AddParagraph(input.LabourHours.ToString());
                var paragraph = row1.Cells[1].AddParagraph();
                var formattedText = new FormattedText() { Style = "Title" };
                formattedText.AddText(input.MaterialDesc);
                paragraph.Add(formattedText);

                row2.Cells[1].AddParagraph($"{input.LabourUnitPrice.ToString("c")}/unit");
                row2.Cells[2].AddParagraph(input.LabourPrice.ToString("c"));
                row2.Cells[3].AddParagraph($"{input.MaterialUnitPrice.ToString("c")}/unit");
                row2.Cells[4].AddParagraph(input.MaterialPrice.ToString("c"));

                labourTotal += input.LabourPrice;
                materialTotal += input.MaterialPrice;


                _table.SetEdge(0, _table.Rows.Count - 2, 5, 2, Edge.Box, BorderStyle.Single, 0.75);
            }


            // Add the labour total columns.
            var row = _table.AddRow();
            row.Cells[0].AddParagraph("Labour Total");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 1;
            row.Cells[2].AddParagraph(labourTotal.ToString("c"));
            row.Cells[2].Format.Font.Name = "Segoe UI";

            // Add the material total columns.
            row.Cells[3].AddParagraph("Materials Total");
            row.Cells[3].Borders.Visible = false;
            row.Cells[3].Format.Font.Bold = true;
            row.Cells[3].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[4].AddParagraph(materialTotal.ToString("c"));
            row.Cells[4].Format.Font.Name = "Segoe UI";

            // Set the borders of the specified cell range.
            _table.SetEdge(2, _table.Rows.Count - 1, 1, 1, Edge.Box, BorderStyle.Single, 0.75);
            _table.SetEdge(4, _table.Rows.Count - 1, 1, 1, Edge.Box, BorderStyle.Single, 0.75);
        }
    }
}
