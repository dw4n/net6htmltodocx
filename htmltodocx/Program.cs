using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using Newtonsoft.Json;

namespace htmltodocx
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            string inputFilePath = Path.Combine(currentDirectory, "input.json");
            string outputFilePath = Path.Combine(currentDirectory, "output.docx");

            if (File.Exists(inputFilePath))
            {
                string jsonContent = File.ReadAllText(inputFilePath);
                var items = JsonConvert.DeserializeObject<List<MyJsonObject>>(jsonContent);
                List<string> failedIds = new List<string>();

                // Create a Wordprocessing document.
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    // Add a new main document part.
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure.
                    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    if (items != null)
                    {
                        foreach (var item in items)
                        {
                            try
                            {
                                // Add a paragraph with the JSON ID
                                Paragraph idParagraph = new Paragraph();
                                Run idRun = new Run();
                                Text idText = new Text($"JSON ID: {item.Id}");
                                idRun.Append(idText);
                                idParagraph.Append(idRun);
                                body.Append(idParagraph);

                                string combinedHtml = item.AttributeCmsDescription + item.AttributeLmsDescription;
                                // Convert HTML to OpenXml and add to the document
                                var converter = new HtmlConverter(mainPart);
                                converter.ParseHtml(combinedHtml);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"An error occurred with item ID: {item.Id}. Error: {ex.Message}");
                                failedIds.Add(item.Id);
                            }
                        }
                    }

                    // Save the document after adding all the HTML content.
                    mainPart.Document.Save();
                }

                // Handle the list of failed IDs after processing all items
                if (failedIds.Count > 0)
                {
                    Console.WriteLine("Failed to convert the following IDs:");
                    foreach (var id in failedIds)
                    {
                        Console.WriteLine(id);
                    }
                }
                else
                {
                    Console.WriteLine("All items were converted successfully.");
                }
            }
            else
            {
                Console.WriteLine("The input file 'input.json' was not found.");
            }
        }

        public class MyJsonObject
        {
            public string Id { get; set; }
            public string AttributeCmsDescription { get; set; }
            public string AttributeLmsDescription { get; set; }
        }

        public static void ConvertHtmlToDocx(string htmlContent, string filePath)
        {
            // Create a Wordprocessing document.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a new main document part.
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure.
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Convert HTML to OpenXml and add to the document
                var converter = new HtmlConverter(mainPart);
                converter.ParseHtml(htmlContent);

                // The converter will automatically add the converted content to the body of the document.
                mainPart.Document.Save();
            }
        }
    }
}