using System;

namespace Excelerate
{
    public class ContentTypes
    {
        private Workbook _workbook;
        public ContentTypes(Workbook workbook)
        {
            _workbook = workbook;
        }

        internal string GenerateContentTypesXml()
        {
            var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + Environment.NewLine;
            xml += "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">";
	        xml += "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>";
	        xml += "<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
            xml += "<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>";
            xml += "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>";
            xml += "<Override PartName=\"/xl/styles.xml\" ContentType =\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
            xml += "<Override PartName=\"/xl/sharedStrings.xml\" ContentType =\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
            xml += "<Override PartName=\"/xl/_rels/workbook.xml.rels\" ContentType =\"application/vnd.openxmlformats-package.relationships+xml\"/>";
            xml += "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>";
	        xml += "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>";
            foreach(var ws in _workbook.Worksheets)
            {
                xml += $"<Override PartName=\"/xl/worksheets/sheet{ws.SheetId}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
            }
            xml += "</Types>";
            return xml;            
        }
    }
}