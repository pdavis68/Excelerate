using System;
using System.Text;
using System.Web;

namespace Excelerate
{
    public class Worksheet : IWorksheet
    {
        private DataGrid _dataGrid = new DataGrid();
        private Workbook _workbook;

        internal Worksheet(Workbook workbook, string name, int sheetId)
        {
            _workbook = workbook;
            Name = name;
            SheetId = sheetId;
        }

        private SharedStrings SharedStrings
        {
            get
            {
                return _workbook.Document.Sharedstrings;
            }
        }

        public object this[string cellReference] 
        { 
            get 
            { 
                var cell = AddressHelper.ToRowCol(cellReference); 
                return _dataGrid?[cell.Y, cell.X]; 
            }
            set
            {
                var cell = AddressHelper.ToRowCol(cellReference);
                if (value is string)
                {
                    int val = SharedStrings.AddString(value as string);
                    _dataGrid[cell.Y, cell.X] = val.ToString();
                }
                else
                {
                    _dataGrid[cell.Y, cell.X] = value;
                }
            }
        }
        
        public object this[int rowIndex, int colIndex] 
        { 
            get => _dataGrid?[rowIndex, colIndex]; 
            set
            {
                if (value is string)
                {
                    int val = SharedStrings.AddString(value as string);
                    _dataGrid[rowIndex, colIndex] = val.ToString();
                }
                else
                {
                    _dataGrid[rowIndex, colIndex] = value;
                }
            }
        }

        public string Name 
        { 
            get; set;
        }

        public int SheetId
        { 
            get; private set;
        }

        public string RelId
        {
            get; internal set;
        }

        internal string GenerateSheetXml()
        {
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + Environment.NewLine;
            xml += "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">";
            xml += "<sheetPr filterMode=\"false\"><pageSetUpPr fitToPage=\"false\"/></sheetPr>";
            xml += "<dimension ref=\"A1\"/>";

            xml += "<sheetViews>";
            xml += "<sheetView showFormulas=\"false\" showGridLines=\"true\" showRowColHeaders=\"true\" showZeros=\"true\" rightToLeft=\"false\" tabSelected=\"true\" showOutlineSymbols=\"true\" defaultGridColor=\"true\" view=\"normal\" topLeftCell=\"A1\" colorId=\"64\" zoomScale=\"100\" zoomScaleNormal=\"100\" zoomScalePageLayoutView=\"100\" workbookViewId=\"0\">";
			xml += "<selection pane=\"topLeft\" activeCell=\"A1\" activeCellId=\"0\" sqref=\"A1\"/>";
            xml += "</sheetView>";
            xml += "</sheetViews>";

            xml += "<sheetFormatPr defaultColWidth=\"11.5703125\" defaultRowHeight=\"12.8\" zeroHeight=\"false\" outlineLevelRow=\"0\" outlineLevelCol=\"0\"/>";

            xml += GenerateSheetData();

            xml += "<printOptions headings=\"false\" gridLines=\"false\" gridLinesSet=\"true\" horizontalCentered=\"false\" verticalCentered=\"false\"/>";
            xml += "<pageMargins left=\"0.7875\" right=\"0.7875\" top=\"1.05277777777778\" bottom=\"1.05277777777778\" header=\"0.7875\" footer=\"0.7875\"/>";
            xml += "<pageSetup paperSize=\"1\" scale=\"100\" fitToWidth=\"1\" fitToHeight=\"1\" pageOrder=\"downThenOver\" orientation=\"portrait\" blackAndWhite=\"false\" draft=\"false\" cellComments=\"none\" firstPageNumber=\"1\" useFirstPageNumber=\"true\" horizontalDpi=\"300\" verticalDpi=\"300\" copies=\"1\"/>";

            xml += "<headerFooter differentFirst=\"false\" differentOddEven=\"false\">";
            xml += "<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>";
            xml += "<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>";
            xml += "</headerFooter>";
            xml += "</worksheet>";
            return xml;
        }

        private string GenerateSheetData()
        {
            string xml = "<sheetData>";

            // 'cause performance y'all
            StringBuilder sb = new StringBuilder();
            for(int rowIdx = 1; rowIdx <= _dataGrid.NumRows; rowIdx++)
            {
                sb.Append($"<row r=\"{rowIdx}\">");
                for(int colIdx = 1; colIdx <= _dataGrid.NumCols; colIdx++)
                {
                    sb.Append(GetCellXml(rowIdx, colIdx));
                }
                sb.Append("</row>");
            }
            xml += sb.ToString();

            xml += "</sheetData>";
            return xml;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <remarks>
        /// cell types: "t={x}"
        ///  n = Number
        ///  s = Shared string
        ///  str = Formula string
        ///  inlineStr = String not in shared string table. Value in @lt;is@gt; tags instead of @lt;v@gt; tags
        ///  e = Cell contains an error
        ///  b - Cell contains a boolean
        /// </remarks>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private string GetCellXml(int row, int col)
        {
            object val = _dataGrid[row, col];
            if (val == null)
            {
                return string.Empty;
            }

            string cellText = string.Empty;
            int style = 0;
            if (val is CellData)
            {
                var cd = (val as CellData);
                style = cd.CellXfsId;
                val = cd.Value;
                if (val is string && !cd.IsSharedString)
                {
                    val = _workbook.Document.Sharedstrings.AddString(val as string).ToString();
                }
            }

            if (val is string)
            {
                cellText = $"<c r=\"{AddressHelper.ToExcelCoordinates(col, row)}\" s=\"{style}\" t=\"s\"><v>{val}</v></c>";
            }
            else if (val is float || val is decimal || val is double || val is byte || val is short ||
                     val is ushort || val is int || val is uint || val is long || val is ulong)
            {
                cellText = $"<c r=\"{AddressHelper.ToExcelCoordinates(col, row)}\" s=\"{style}\" t=\"n\"><v>{HttpUtility.HtmlEncode(val.ToString())}</v></c>";
            }
            else if (val is bool)
            {
                cellText = $"<c r=\"{AddressHelper.ToExcelCoordinates(col, row)}\" s=\"{style}\" t=\"b\"><v>{HttpUtility.HtmlEncode(val.ToString())}</v></c>";
            }
            else
            {
                cellText = $"<c r=\"{AddressHelper.ToExcelCoordinates(col, row)}\" s=\"{style}\" t=\"inlineStr\"><is>{HttpUtility.HtmlEncode(val.ToString())}</is></c>";
            }
            return cellText;

        }

        public void ReleaseMemory()
        {
            
        }
    }
}