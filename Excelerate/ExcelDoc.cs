using System.IO;
using Excelerate.Styles;

namespace Excelerate;
public class ExcelDoc
{
    public SharedStrings Sharedstrings { get; } = new SharedStrings();
    public Workbook Workbook { get; init; }
    public StyleManager StyleManager { get; } = new StyleManager();
    public Relationships DocRelationships { get; } = new Relationships();
    public ContentTypes ContentTypes { get; init; }
    public AppInfo AppInfo { get; } = new AppInfo();
    public CoreInfo CoreInfo { get; } = new CoreInfo();

    public ExcelDoc()
    {
        Workbook = new Workbook(this);
        ContentTypes = new ContentTypes(Workbook);
        DocRelationships.AddRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
        DocRelationships.AddRelationship("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
        DocRelationships.AddRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
    }

    public ExcelDoc(string filename) : this()
    {

    }
    
    public ExcelDoc(Stream stream) : this()
    {

    }

    public void SaveAs(string filename)
    {
        FileSaver.Save(this, filename);
    }

}
