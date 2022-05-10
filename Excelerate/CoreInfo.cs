using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace Excelerate
{
    public class CoreInfo
    {
        public CoreInfo()
        {
        }

        public string Title { get; set; }
        public string Subject { get; set; }
        public string Creator { get; set; }
        public string Keywords { get; set; }
        public string Description { get; set; }
        public string LastModifiedBy { get; set; }
        public string Revision { get; set; }
        public DateTime Created { get; set; } = DateTime.Now;
        public DateTime Modified { get; set; } = DateTime.Now;
        public string Category { get; set; }
        public string ContentStatus { get; set; }

        internal string GenerateCoreInfoXml()
        {
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + Environment.NewLine;
            xml += "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">";
            xml += AddProp("dc:title", Title);
            xml += AddProp("dc:subject", Subject);
            xml += AddProp("dc:creator", Creator);
            xml += AddProp("cp:keywords", Keywords);
            xml += AddProp("dc:description", Description);
            xml += AddProp("cp:lastModifiedBy", LastModifiedBy);
            xml += AddProp("cp:revision", Revision);
            xml += $"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{Created.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")}</dcterms:created>";
            xml += $"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{Modified.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")}</dcterms:modified>";
            xml += AddProp("cp:category", Category);
            xml += AddProp("cp:contentStatus", ContentStatus);
            xml += "</cp:coreProperties>";
            return xml;
        }

        private string AddProp(string element, string value)
        {
            string encValue = HttpUtility.HtmlEncode(value);
            return string.IsNullOrEmpty(value) ? $"<{element}/>" : $"<{element}>{encValue}</{element}>";
        }
    }
}