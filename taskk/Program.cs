using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

public class Program
{
    public static void Main()
    {
        string outputPDFPath = "C:\\Users\\gnoor\\source\\repos\\taskk\\taskk\\newDoc.pdf";
        string jsonContent = File.ReadAllText("C:\\Users\\gnoor\\source\\repos\\taskk\\taskk\\data.json");
        dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);

        string inputWordPath = "C:\\Users\\gnoor\\source\\repos\\taskk\\taskk\\karbarg.docx";
        string outputWordPath = "C:\\Users\\gnoor\\source\\repos\\taskk\\taskk\\newDoc.docx";

        File.Copy(inputWordPath, outputWordPath, true);

        using (WordprocessingDocument doc = WordprocessingDocument.Open(outputWordPath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Table table = mainPart.Document.Body.Elements<Table>().Last();

            int rowIndex = 0;
            foreach (var issue in jsonData.مسائل)
            {
                TableRow row = new TableRow();

                AddCellToRow(row, issue.ردیف.ToString());
                AddCellToRowWithSubTable(row, issue["نشانه‌های مسئله"].ToObject<List<string>>());
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

                AddCellToRowWithSubTable(row, issue["دستگاه‌های تابعه همکار"].ToObject<List<string>>());

                // Set background color for even rows
                if (rowIndex % 2 != 0)
                {
                    foreach (TableCell cell in row.Elements<TableCell>())
                    {
                        cell.Append(new TableCellProperties(
                            new Shading()
                            {
                                Color = "auto",
                                Fill = "EAEAEA"
                            }
                        ));
                    }
                }

                table.Append(row);
                rowIndex++;
            }

            mainPart.Document.Save();
        }

        // Export to PDF
        Spire.Doc.Document document = new();
        document.LoadFromFile(outputWordPath);
        document.SaveToFile(outputPDFPath, Spire.Doc.FileFormat.PDF);

        Console.WriteLine("Ok");

        void AddCellToRow(TableRow row, string text)
        {
            TableCell cell = new TableCell(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Right }), new Run(new Text(text))));
            ApplyFont(cell);
            row.Append(cell);
        }

        void AddCellToRowWithSubTable(TableRow row, List<string> values)
        {
            TableCell cell = new TableCell();
            Table subTable = new Table();
            subTable.AppendChild(new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }
                )
            ));

            foreach (var value in values)
            {
                TableRow subRow = new TableRow();
                TableCell subCell = new TableCell(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Right }), new Run(new Text(value))));
                ApplyFont(subCell);
                subCell.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Auto }
                ));
                subRow.Append(subCell);
                subTable.Append(subRow);
            }

            cell.Append(subTable);
            row.Append(cell);
        }

        void ApplyFont(TableCell cell)
        {
            foreach (var paragraph in cell.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    if (run.RunProperties == null)
                        run.RunProperties = new RunProperties();
                    run.RunProperties.RunFonts = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
                }
            }
        }
    }
}
