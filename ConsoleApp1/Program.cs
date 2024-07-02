using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


public class Program
{
    public static void Main()
    {
        string outputPDFPath = "C:\\Users\\gnoor\\source\\repos\\taskk\\taskk\\newDoc.pdf";
        string jsonFilePath = @"C:\Users\gnoor\source\repos\taskk\taskk\data.json";
        string inputWordPath = @"C:\Users\gnoor\source\repos\taskk\taskk\karbarg.docx";
        string outputWordPath = @"C:\Users\gnoor\source\repos\taskk\taskk\newDoc.docx";
        try
        {
            File.Copy(inputWordPath, outputWordPath, true);

            string jsonContent = File.ReadAllText(jsonFilePath);
            dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputWordPath, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;
                Table table = mainPart.Document.Body.Elements<Table>().Last();

                while (table.Elements<TableRow>().Count() > 1)
                {
                    table.RemoveChild(table.Elements<TableRow>().Last());
                }

                int rowIndex = 0;
                foreach (var issue in jsonData.مسائل)
                {
                    TableRow row = CreateTableRow(issue, rowIndex);
                    table.Append(row);
                    rowIndex++;
                }

                mainPart.Document.Save();
            }

            Console.WriteLine("Word document updated successfully. Converting to PDF...");

            Spire.Doc.Document document = new();
            document.LoadFromFile(outputWordPath);
            document.SaveToFile(outputPDFPath, Spire.Doc.FileFormat.PDF);

            Console.WriteLine("PDF created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static TableRow CreateTableRow(dynamic issue, int rowIndex)
    {
        TableRow row = new TableRow();

        AddCellToRow(row, issue.ردیف.ToString());
        AddCellWithNestedTable(row, issue["نشانه‌های مسئله"].ToObject<List<string>>());
        AddCellToRow(row, issue.مسئله.ToString());
        AddCellToRow(row, issue["سطح مسئله"].ToString());

        if (issue["نماگرهای وضعیت مسئله"].Count > 0)
        {
            var indicator = issue["نماگرهای وضعیت مسئله"][0];
            AddCellToRow(row, indicator.عنوان.ToString());
            AddCellToRow(row, indicator["نوع نماگر"].ToString());
            AddCellToRow(row, indicator.فرمول.ToString());
            AddCellToRow(row, indicator["مقدار کنونی"].ToString());
            AddCellToRow(row, indicator["مقدار مطلوب"].ToString());
            AddCellToRow(row, indicator["مرجع اخذ داده"].ToString());
        }
        else
        {
            for (int i = 0; i < 6; i++)
                AddCellToRow(row, "");
        }

        AddCellToRow(row, issue.مأموریت.ToString());

        if (issue["خدمات متناظر"].Count > 0)
        {
            AddCellToRow(row, issue["خدمات متناظر"][0]["کلان خدمت"].ToString());
            AddCellToRow(row, issue["خدمات متناظر"][0].زیرخدمت.ToString());
        }
        else
        {
            AddCellToRow(row, "");
            AddCellToRow(row, "");
        }

        AddCellWithNestedTable(row, issue["دستگاه‌های تابعه همکار"].ToObject<List<string>>());

        if (rowIndex % 2 != 0)
        {
            ApplyBackgroundColorToRow(row, "EAEAEA");
        }

        return row;
    }

    private static void AddCellToRow(TableRow row, string text)
    {
        TableCell cell = new TableCell(
            new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Auto }
            ),
            new Paragraph(
                new ParagraphProperties(new Justification() { Val = JustificationValues.Right }),
                new Run(
                    new RunProperties(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" }),
                    new Text(text)
                )
            )
        );
        row.Append(cell);
    }

    private static void AddCellWithNestedTable(TableRow row, List<string> values)
    {
        TableCell cell = new TableCell();
        Table nestedTable = new Table();

        TableProperties tblProp = new TableProperties(
            new TableBorders(
                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None) },
                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None) },
                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None) },
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None) },
                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None) }
            ),
            new TableWidth() { Width = "100%", Type = TableWidthUnitValues.Pct }
        );
        nestedTable.AppendChild(tblProp);

        foreach (var value in values)
        {
            TableRow nestedRow = new TableRow();
            TableCell nestedCell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Auto }
                ),
                new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Right }),
                    new Run(
                        new RunProperties(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" }),
                        new Text(value)
                    )
                )
            );
            nestedRow.Append(nestedCell);
            nestedTable.Append(nestedRow);
        }

        cell.Append(nestedTable);
        row.Append(cell);
    }

    private static void ApplyBackgroundColorToRow(TableRow row, string colorCode)
    {
        foreach (TableCell cell in row.Elements<TableCell>())
        {
            if (cell.TableCellProperties == null)
            {
                cell.TableCellProperties = new TableCellProperties();
            }

            Shading shading = new Shading()
            {
                Color = "auto",
                Fill = colorCode,
                Val = ShadingPatternValues.Clear
            };

            cell.TableCellProperties.RemoveAllChildren<Shading>();

            cell.TableCellProperties.AppendChild(shading);
        }
    }
}
