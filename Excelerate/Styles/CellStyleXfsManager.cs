using System;
using System.Collections.Generic;

namespace Excelerate.Styles
{
    public class CellStyleXfsManager
    {
        private List<CellStyleXfs> _cellStyleXfs = new List<CellStyleXfs>();

        internal string GenerateCellStyleXfsXml()
        {

            string xml = "<cellStyleXfs count=\"20\">";
            xml += "  <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"true\" applyAlignment=\"true\" applyProtection=\"true\">";
            xml += "    <alignment horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"false\" indent=\"0\" shrinkToFit=\"false\" />";
            xml += "    <protection locked=\"true\" hidden=\"false\" />";
            xml += "  </xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"43\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"41\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"44\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"42\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "  <xf numFmtId=\"9\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"></xf>";
            xml += "</cellStyleXfs>";
            return xml;
            /*
            string xml = $"<cellStyleXfs count=\"{_cellStyleXfs.Count}\">";
            foreach(var cellStyleXfs in _cellStyleXfs)
            {
                xml += cellStyleXfs.GenerateCellStyleXfsXml();
            }
            xml += "</cellStyleXfs>";
            return xml;
            */
        }
    }
}