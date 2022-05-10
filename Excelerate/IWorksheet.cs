namespace Excelerate
{
    public interface IWorksheet
    {
        object this[int rowIndex, int colIndex]
        {
            get;
            set;
        }
        object this[string cellReference]
        {
            get;
            set;
        }

        string Name { get; set; }
        int SheetId { get; }
        string RelId
        {
            get;
        }


        /// <summary>
        /// Serializes the DataGrid to a temporary file. If it's referenced again, 
        /// the DataGrid will be restored from the file.
        /// </summary
        void ReleaseMemory();
        
    }
}