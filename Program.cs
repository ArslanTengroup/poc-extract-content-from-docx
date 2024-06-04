using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;


public class WordDocumentProcessor
{
    private readonly HttpClient _httpClient;


    public WordDocumentProcessor(HttpClient httpClient)
    {
        _httpClient = httpClient;
    }


    public async Task<string> ExtractTextFromDocxAsync(string documentUrl)
    {
        try
        {
            using var response = await _httpClient.GetAsync(documentUrl, HttpCompletionOption.ResponseHeadersRead);
            response.EnsureSuccessStatusCode(); //will throw an exception if the status code is not a success code
            using var stream = await response.Content.ReadAsStreamAsync(); //read the content as a stream

            using var wordDocument = WordprocessingDocument.Open(stream, false); //false mean readonly
            return ExtractText(wordDocument);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            return string.Empty;
        }
    }


    private string ExtractText(WordprocessingDocument document)
    {
        StringBuilder sb = new StringBuilder();
        var body = document.MainDocumentPart?.Document?.Body;
        ExtractTextFromElement(body, sb);
        return sb.ToString();
    }


    private void ExtractTextFromElement(OpenXmlElement element, StringBuilder sb)
    {
        if (element == null)
        {
            return;
        }

        foreach (var para in element.Descendants<Paragraph>())
        {
            sb.AppendLine(para.InnerText);
        }
    }
}

class Program
{
    static async Task Main(string[] args)
    {
        using var httpClient = new HttpClient();
        WordDocumentProcessor processor = new WordDocumentProcessor(httpClient);

        string url = ""; //add url for the docx file
        try
        {
            string text = await processor.ExtractTextFromDocxAsync(url);
            Console.WriteLine(text);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to extract text: {ex.Message}");
        }
        Console.ReadLine();
    }
}
