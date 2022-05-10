using System;
using System.Collections.Generic;

namespace Excelerate.Styles
{
    public class StyleManager
    {
        private NumberFormatsManager _numberFormatsManager = new NumberFormatsManager();
        private FontManager _fontManager = new FontManager();
        private FillManager _fillManager = new FillManager();
        private BorderManager _borderManager = new BorderManager();
        private CellStyleXfsManager _cellStyleXfsManager = new CellStyleXfsManager();
        private CellXfsManager _cellXfsManager = new CellXfsManager();
        private CellStylesManager _cellStylesManager = new CellStylesManager();

        public StyleManager()
        {

        }

        public int AddFill(FillBase fill)
        {
            return _fillManager.AddFill(fill);
        }

        public int AddFont(Font font)
        {
            return _fontManager.AddFont(font);
        }

        public int AddCellXfs(CellXfs cellXfs)
        {
            return _cellXfsManager.AddCellXfs(cellXfs);
        }

        internal string GenerateStylesXml()
        {
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + Environment.NewLine;
            xml += "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
            xml += _numberFormatsManager.GenerateNumberFormatXml();  //
            xml += _fontManager.GenerateFontXml(); //
            xml += _fillManager.GenerateFillXml(); //
            xml += _borderManager.GenerateBorderXml(); //
            xml += _cellStyleXfsManager.GenerateCellStyleXfsXml();
            xml += _cellXfsManager.GenerateCellXfsXml();
            xml += _cellStylesManager.GenerateCellStylesXml();
            xml += "</styleSheet>";
            return xml;
        }
        
    }
}