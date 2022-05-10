using System;
using System.Threading;
using System.Web;

namespace Excelerate
{
    public class Workbook
    {
        private int _sheetId = 0;
        private object _lockOb = new object();
        private WorksheetCollection _worksheets = new WorksheetCollection();
        public Relationships Relationships { get; } = new Relationships();
        public ExcelDoc Document { get; init; }

        public Workbook(ExcelDoc document)
        {
            Document = document;
            AddSheet("Sheet1");
            Relationships.AddRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");
            Relationships.AddRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
        }

        public IWorksheet AddSheet(string name)
        {
            var sheet = new Worksheet(this, name, Interlocked.Increment(ref _sheetId));
            lock(_lockOb)
            {
                sheet.RelId = Relationships.AddRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", $"worksheets/sheet{sheet.SheetId}.xml");
                _worksheets.Add(sheet);
            }
            return sheet;
        }

        public IWorksheet GetSheet(string name)
        {
            lock (_lockOb)
            {
                {
                    return _worksheets[name];
                }
            }
        }

        public void RemoveWorksheet(string name)
        {
            lock(_lockOb)
            {
                _worksheets.Remove(_worksheets[name]);
            }
        }

        public void RenameWorksheet(string oldName, string newName)
        {
            lock(_lockOb)
            {
                var ws = _worksheets[oldName];
                ws.Name = newName; 

            }
        }

        internal WorksheetCollection Worksheets { get { lock(_lockOb) { return _worksheets;} } }

        internal string GenerateWorkbookXml()
        {
            var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + Environment.NewLine;
            xml += $"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";

            xml += $"<fileVersion appName=\"Excelerate\"/>";
            xml += "<bookViews>";
            xml += "<workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"0\" windowHeight=\"0\"/>";
            xml += "</bookViews>";

            xml += "<sheets>";
            foreach(var ws in _worksheets)
            {
                xml += $"<sheet name=\"{HttpUtility.HtmlEncode(ws.Name)}\" sheetId=\"sheet{ws.SheetId}\" r:id=\"{ws.RelId}\" />";
            }
            xml += "</sheets>";

            /// TODO: Need to research 
            xml += "<calcPr calcId=\"0\" refMode=\"R1C1\" iterateCount=\"0\" calcOnSave=\"0\" concurrentCalc=\"0\"/>";
            xml += "</workbook>";
            return xml;
        }        
    }
}