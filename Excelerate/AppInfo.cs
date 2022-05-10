using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excelerate
{
    public class AppInfo
    {
        public string Application {get => "Excelerate";}
        public string Manager {get => "";}
        public string Company {get => "";}
        public string HyperlinkBase {get => "";}
        public string AppVersion {get => "0.9.0.0 βeta";}

        internal string GenerateAppInfoXml()
        {
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + Environment.NewLine;
            xml += "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">";
            xml += $"<Application>{Application}</Application>";
            xml += string.IsNullOrEmpty(Manager) ? "<Manager/>" : $"<Manager>{Manager}</Manager>";
            xml += string.IsNullOrEmpty(Company) ? "<Company/>" : $"<Company>{Company}</Company>";
            xml += string.IsNullOrEmpty(HyperlinkBase) ? "<HyperlinkBase/>" : $"<HyperlinkBase>{HyperlinkBase}</HyperlinkBase>";
            xml += $"<AppVersion>{AppVersion}</AppVersion>";
            xml += "</Properties>";
            return xml;
        }
    }
}