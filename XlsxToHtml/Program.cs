namespace XlsxToHtml
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Converting...");
            string sourceDirectory = Directory.GetCurrentDirectory();
            string targetDirectory = Path.Combine(sourceDirectory, "html");
            Directory.CreateDirectory(targetDirectory);
            string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx");
            foreach (var fileItem in xlsxFiles.Select((path, index) => (path, index)))
            {
                Console.WriteLine($"File{fileItem.index + 1}/{xlsxFiles.Length}:{fileItem.path}");
                var html = XlsxToHtml.Convert(fileItem.path);
                var fileName = Path.GetFileNameWithoutExtension(fileItem.path);
                var newFilePath = Path.Combine(targetDirectory, fileName + ".html");
                File.WriteAllText(newFilePath, html);
            }

            Console.WriteLine("Complete!!!");
        }
    }
}