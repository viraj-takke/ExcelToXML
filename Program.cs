using ClosedXML.Excel;
using ExcelToXML.Models;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Xml;
using CellType = NPOI.SS.UserModel.CellType;

namespace ExcelToXML
{
    public class Program
    {
        private static readonly IConfiguration _configuration;

        static Program()
        {
            _configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();
        }

        public static void Main(string[] args)
        {
            try
            {
                var excelPath = _configuration["ExcelPath"];
                var excelName = _configuration["ExcelName"];
                var excelFullPath = Path.Combine(excelPath, excelName);

                var generator = new XmlGenerator(_configuration);
                generator.ProcessExcelFile(excelFullPath);

                Console.WriteLine("XML files generated successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class XmlGenerator
    {
        private readonly IConfiguration _configuration;

        public XmlGenerator(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public void ProcessExcelFile(string excelFilePath)
        {
            try
            {
                if (!File.Exists(excelFilePath))
                {
                    throw new FileNotFoundException($"Excel file not found: {excelFilePath}");
                }

                var orders = new List<OrderData>();
                var orderGroups = new Dictionary<string, OrderData>();

                using (var fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    if (excelFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        workbook = new XSSFWorkbook(fileStream);
                    else if (excelFilePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                        workbook = new HSSFWorkbook(fileStream);
                    else
                        throw new Exception("Unsupported file format. Use .xlsx or .xls.");

                    var sheetName = _configuration["SheetName"];
                    var sheet = workbook.GetSheet(sheetName) ?? throw new Exception($"Worksheet '{sheetName}' not found.");

                    List<int> highlightedRows = GetHighlightedRows(excelFilePath);
                    Console.WriteLine($"Found {highlightedRows.Count} highlighted rows: {string.Join(", ", highlightedRows.OrderBy(x => x))}");

                    int rowNumber = 0;
                    for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        rowNumber = rowIndex + 1;
                        var row = sheet.GetRow(rowIndex);
                        if (row == null) continue;

                        if (highlightedRows.Contains(rowNumber))
                        {
                            try
                            {
                                var id = row.GetCell(0)?.ToString() ?? "";
                                var date = DateTime.Parse(row.GetCell(1)?.ToString() ?? throw new Exception("Date is missing"));
                                var city = row.GetCell(2)?.ToString() ?? "";
                                var category = row.GetCell(3)?.ToString() ?? "";
                                var product = row.GetCell(4)?.ToString() ?? "";
                                var quantity = int.Parse(row.GetCell(5)?.ToString() ?? "0");
                                var contact = row.GetCell(6)?.ToString() ?? "";
                                var unitPrice = decimal.Parse(row.GetCell(7)?.ToString() ?? "0");

                                var evaluator = new XSSFFormulaEvaluator(workbook);

                                var cell = row.GetCell(8);
                                decimal totalPrice = 0;

                                if (cell != null)
                                {
                                    if (cell.CellType == CellType.Formula)
                                    {
                                        var eval = evaluator.Evaluate(cell);
                                        if (eval.CellType == CellType.Numeric)
                                        {
                                            totalPrice = (decimal)eval.NumberValue;
                                        }
                                        else
                                        {
                                            totalPrice = 0;
                                        }
                                    }
                                    else
                                    {
                                        totalPrice = (decimal)cell.NumericCellValue;
                                    }
                                }

                                var contactInfo = ParseContactInfo(contact);

                                string baseId = !string.IsNullOrEmpty(id) ? id : $"ORD-{city?.Replace(" ", "")}";
                                string shipName = !string.IsNullOrEmpty(contactInfo.Name) ? contactInfo.Name : "User";
                                string shippingKey = $"{city}-{contactInfo.Region}".Replace(" ", "").Replace(",", "");
                                string UniqueId = $"{baseId}-{shipName}-{shippingKey}-{date:yyyyMMdd}";

                                if (!orderGroups.ContainsKey(UniqueId))
                                {
                                    orderGroups[UniqueId] = new OrderData
                                    {
                                        UniqueIdentity = UniqueId,
                                        OrderId = id,
                                        OrderDate = date,
                                        OrderPerson = "Food Sales System",
                                        ShipToName = contactInfo.Name,
                                        ShipToAddress = contactInfo.Address,
                                        ShipToCity = city,
                                        ShipToRegion = contactInfo.Region
                                    };
                                }

                                orderGroups[UniqueId].Items.Add(new ItemData
                                {
                                    Title = $"{category} - {product}",
                                    Note = $"Category: {category}",
                                    Quantity = quantity,
                                    Price = unitPrice,
                                    Total = totalPrice
                                });
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Error processing row {rowNumber}: {ex.Message}");
                            }
                        }
                    }
                }

                GenerateXmlFiles(orderGroups.Values.ToList());
            }
            catch (Exception ex)
            {
                throw new Exception($"Something went wrong: {ex.Message}");
            }

        }

        private List<int> GetHighlightedRows(string filePath)
        {
            List<int> highlightedRows = new List<int>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var usedRange = worksheet.RangeUsed();
                int column = usedRange.ColumnCount();

                if (usedRange != null)
                {
                    for (int row = 2; row <= usedRange.RowCount(); row++)
                    {
                        var cell = worksheet.Cell(row, 1);
                        if (cell.Style.Fill.BackgroundColor.Color.Name != "Transparent" &&
                            cell.Style.Fill.BackgroundColor.Color.Name != "White" &&
                            cell.Style.Fill.BackgroundColor.Color.Name != "000000")
                        {
                            highlightedRows.Add(row);
                        }
                    }
                }
            }

            return highlightedRows;
        }

        private static ContactInfo ParseContactInfo(string contact)
        {
            if (contact.Contains(","))
            {
                var parts = contact.Split(',').Select(p => p.Trim()).ToArray();

                if (parts.Length >= 4)
                {
                    string name = parts[0];
                    string streetAddress = parts[1];
                    string city = parts[2];

                    string region = parts[parts.Length - 2];

                    string fullAddress = $"{streetAddress}, {city}";

                    return new ContactInfo
                    {
                        Name = name,
                        Address = fullAddress,
                        City = city,
                        Region = region
                    };
                }
            }

            return new ContactInfo
            {
                Name = "Unknown",
                Address = "Unknown",
                City = "Unknown",
                Region = "Unknown"
            };
        }

        private void GenerateXmlFiles(List<OrderData> orders)
        {
            try
            {
                var outputDir = _configuration["XmlOutputPath"];
                Directory.CreateDirectory(outputDir);

                foreach (var order in orders)
                {
                    using (var writer = XmlWriter.Create(Path.Combine(outputDir, $"{order.UniqueIdentity}.xml")))
                    {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("shiporder");
                        writer.WriteAttributeString("orderid", order.OrderId);
                        writer.WriteAttributeString("orderdate", order.OrderDate.ToString("yyyy-MM-dd"));
                        writer.WriteElementString("orderperson", order.OrderPerson);
                        writer.WriteStartElement("shipto");
                        writer.WriteElementString("name", order.ShipToName);
                        writer.WriteElementString("address", order.ShipToAddress);
                        writer.WriteElementString("city", order.ShipToCity);
                        writer.WriteElementString("region", order.ShipToRegion);
                        writer.WriteEndElement();

                        foreach (var item in order.Items)
                        {
                            writer.WriteStartElement("item");
                            writer.WriteElementString("title", item.Title);

                            if (!string.IsNullOrEmpty(item.Note))
                            {
                                writer.WriteElementString("note", item.Note);
                            }

                            writer.WriteElementString("quantity", item.Quantity.ToString());
                            writer.WriteElementString("price", item.Price.ToString("F2"));
                            writer.WriteElementString("total", item.Total.ToString("F2"));
                            writer.WriteEndElement();
                        }

                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Something went wrong: {ex.Message}");
            }

        }
    }
}