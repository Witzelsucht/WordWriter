using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;

var directory = Directory.GetCurrentDirectory();
var files = Directory.EnumerateFiles(directory, "*.docx");

Console.WriteLine($"Found {files.Count()} docx files.");

List<JObject> jsonObjects = new();
using var reader = new JsonTextReader(new StreamReader("input.json"));
{
    var serializer = new JsonSerializer();
    jsonObjects = serializer.Deserialize<List<JObject>>(reader);
}

Console.WriteLine($"Found {jsonObjects.Count} records in input.json file");

foreach(var file in files)
{
    Console.WriteLine($"Current document: {file}");
    foreach (var jsonObject in jsonObjects)
    {
        var name = (string)jsonObject["filename-prefix"];
        if(name is null)
        {
            Console.WriteLine($"WARNING: one of the records doesn't have 'filename-prefix' property set. Skipping record...");
            continue;
        }
        Console.WriteLine($"Current record {name}");
        var properties = jsonObject.Children();
        using var document = DocX.Load(file);
        foreach (var property in properties)
        {
            if (document.FindUniqueByPattern($@"<{property.Path}>", RegexOptions.IgnoreCase).Count <= 0) continue;
            var replaceOptions = new FunctionReplaceTextOptions()
            {
                FindPattern = $@"<{property.Path.ToLower()}>",
                RegExOptions = RegexOptions.IgnoreCase,
                RegexMatchHandler = (findStr) => (string)jsonObject[property.Path]
            };
            document.ReplaceText(replaceOptions);
            Console.WriteLine($@"<{property.Path}> => {(string)jsonObject[property.Path]}");
        }
        document.SaveAs($"{name}_{file}");
        Console.WriteLine($"Document saved as: {name}_{file}.docx");
    }
}

Console.WriteLine("Replacement finished");

Console.ReadKey();
