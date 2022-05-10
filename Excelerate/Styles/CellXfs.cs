using System;

namespace Excelerate.Styles
{
    public class CellXfs
    {
        public int XFId { get; set; } = 0;
        public int NumberFmtId { get; set; } = 164;
        public int FontId { get; set; } = 0;
        public int FillId { get; set; } = 0;
        public int BorderId { get; set; } = 0;




        internal string GenerateCellXfsXml()
        {
            string xml = $"<xf numFmtId=\"{NumberFmtId}\" fontId=\"{FontId}\" fillId=\"{FillId}\" borderId=\"{BorderId}\" xfId=\"{XFId}\" applyFont=\"{(((FontId != 0) || (FillId != 0)) ? "true" : "false")}\" applyBorder=\"{((BorderId != 0) ? "true" : "false")}\" applyAlignment=\"false\" applyProtection=\"false\">";
            xml += "<alignment horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"false\" indent=\"0\" shrinkToFit=\"false\"/>";
            xml += "<protection locked=\"true\" hidden=\"false\"/>";
            xml += "</xf>";
            return xml;
        }
    }
}