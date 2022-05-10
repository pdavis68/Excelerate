using System.IO;

namespace Excelerate
{
    internal static class FileSaver
    {
        internal static void Save(ExcelDoc excelDoc, string filename)
        {
            using(var tempFolder = new TempFolder())
            {
                // NOTE: Write worksheets first. In writing the worksheet, other files may be updated (like styles or shared string)
                // so they should be written after the worksheets are written.

                // /xl/worksheets
                var wb = excelDoc.Workbook;
                foreach (var sheet in wb.Worksheets)
                {
                    WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.WORKSHEET_PATH, $"sheet{sheet.SheetId}.xml"), (sheet as Worksheet).GenerateSheetXml());
                }

                // /
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.CONTENTTYPE), excelDoc.ContentTypes.GenerateContentTypesXml());

                // /_rels/.rels
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.RELS), excelDoc.DocRelationships.GenerateRelationshipsXml());

                // /docProps/
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.APPINFO), excelDoc.AppInfo.GenerateAppInfoXml());
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.COREINFO), excelDoc.CoreInfo.GenerateCoreInfoXml());

                // /xl 
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.WORKBOOK), wb.GenerateWorkbookXml());
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.SHAREDSTRINGS), excelDoc.Sharedstrings.GenerateSharedStringsXml());
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.STYLES), excelDoc.StyleManager.GenerateStylesXml());

                // /xl/_rels/workbook.xml.rels
                WriteFile(Path.Combine(tempFolder.FolderPath, ExcelPaths.XLRELS), excelDoc.Workbook.Relationships.GenerateRelationshipsXml());

                // /xl/theme  ---- NOTE - May not do theme

                

                if (!Directory.Exists(Path.GetDirectoryName(filename)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(filename));
                }
                File.WriteAllBytes(filename, tempFolder.CreateZipFile());
            }
        }

        private static void WriteFile(string filepath, string data)
        {
            // Yeah, yeah, yeah... Gonna add some error handling & stuff.
            if (!Directory.Exists(Path.GetDirectoryName(filepath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filepath));
            }
            File.WriteAllText(filepath, data);
        }
    }
}