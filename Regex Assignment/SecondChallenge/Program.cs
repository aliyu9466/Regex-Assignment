// Second Challenge

using System;
using System.IO;
using System.Linq;

public class OfficeFileSummary
{
    public static void Main()
    {
        string directoryName = "FileCollection";
        string resultsFileName = "results.txt";

        Directory.CreateDirectory(directoryName);

        int xlsxCount = 0, docxCount = 0, pptxCount = 0;
        long xlsxSize = 0, docxSize = 0, pptxSize = 0;

        DirectoryInfo dirInfo = new DirectoryInfo(directoryName);

        foreach (FileInfo file in dirInfo.GetFiles())
        {
            if (IsOfficeFile(file))
            {
                switch (file.Extension.ToLower())
                {
                    case ".xlsx":
                        xlsxCount++;
                        xlsxSize += file.Length;
                        break;
                    case ".docx":
                        docxCount++;
                        docxSize += file.Length;
                        break;
                    case ".pptx":
                        pptxCount++;
                        pptxSize += file.Length;
                        break;
                }
            }
        }

        using (StreamWriter writer = new StreamWriter(resultsFileName))
        {
            writer.WriteLine("Office File Summary:");
            writer.WriteLine($"Excel files (.xlsx): {xlsxCount}, Total size: {xlsxSize} bytes");
            writer.WriteLine($"Word files (.docx): {docxCount}, Total size: {docxSize} bytes");
            writer.WriteLine($"PowerPoint files (.pptx): {pptxCount}, Total size: {pptxSize} bytes");
        }

        Console.WriteLine($"Results written to {resultsFileName}");
    }

    private static bool IsOfficeFile(FileInfo file)
    {
        string[] officeExtensions = { ".xlsx", ".docx", ".pptx" };
        return officeExtensions.Contains(file.Extension.ToLower());
    }
}
