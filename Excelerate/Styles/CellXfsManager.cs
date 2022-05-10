using System;
using System.Collections.Generic;

namespace Excelerate.Styles
{
    public class CellXfsManager
    {
        private List<CellXfs> _cellXfses = new List<CellXfs>();

        public CellXfsManager()
        {
            _cellXfses.Add(new CellXfs());
        }

        public int AddCellXfs(CellXfs cellXfs)
        {
            int loc = _cellXfses.Count;
            _cellXfses.Add(cellXfs);
            return loc;
        }
        internal string GenerateCellXfsXml()
        {
            string xml = $"<cellXfs count=\"{_cellXfses.Count}\">";
            foreach (var cellXfs in _cellXfses)
            {
                xml += cellXfs.GenerateCellXfsXml();
            }
            xml += "</cellXfs>";
            return xml;
        }

    }
}