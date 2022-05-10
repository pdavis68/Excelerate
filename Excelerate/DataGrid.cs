using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Excelerate
{
    public class DataGrid : ISerializable
    {
        private int _maxCol = 1;

        private Dictionary<int, RowItem> _rows = new Dictionary<int, RowItem>();

        public DataGrid()
        {
            
        }

        /// <summary>
        /// ISeralializable constructor.
        /// </summary>
        internal DataGrid(SerializationInfo info, StreamingContext context)
        {
            // Reset the property value using the GetValue method.
            _maxCol = (int) info.GetValue("_maxCol", typeof(int));
            _rows = (Dictionary<int, RowItem>) info.GetValue("_rows", typeof(Dictionary<int, RowItem>));
        }
        internal object this[int rowIndex, int colIndex] 
        { 
            get
            {
                return GetItem(rowIndex, colIndex);
            }
            set
            {
                InsertItem(rowIndex, colIndex, value);
            } 
        }

        public int NumRows { get => _rows.Count; }
        public int NumCols { get => _maxCol; }

        internal object GetItem(int rowIndex, int colIndex)
        {
            if (_rows.ContainsKey(rowIndex))
            {
                var value = _rows[rowIndex].GetItem(colIndex);
                System.Diagnostics.Debug.WriteLine($"GetItem r:{rowIndex} c:{colIndex} v:{value}");
                return value;
            }
            return null;
        }

        public void InsertItem(int rowIndex, int colIndex, object value)
        {
            System.Diagnostics.Debug.WriteLine($"InsertItem r:{rowIndex} c:{colIndex}");
            if (_rows.ContainsKey(rowIndex))
            {
                _maxCol = Math.Max(_maxCol, colIndex);
                _rows[rowIndex].InsertItem(colIndex, value);
            }
            else
            {
                _maxCol = Math.Max(_maxCol, colIndex);
                // Insert any intervening rows if we skip.
                for (int index = _rows.Count + 1; index <= rowIndex; index++)
                {
                    var rowItem = new RowItem();
                    _rows.Add(index, rowItem);

                }
                _rows[rowIndex].InsertItem(colIndex, value);
            }
        }

        public object[] GetRow(int rowIndex)
        {
            if (_rows.ContainsKey(rowIndex))
            {
                return _rows[rowIndex].GetAllItems();
            }
            return new object[_maxCol];
        }


        /// <summary>
        /// ISerializable
        /// </summary>
        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            // Use the AddValue method to specify serialized values.
            info.AddValue("_maxCol", _maxCol, typeof(int));
            info.AddValue("_rows", _rows, typeof(Dictionary<int, RowItem>));
        }

        private class RowItem : ISerializable
        {
            private int _maxCol = 1;
            internal Dictionary<int, object> ColItems { get; private set;} = new Dictionary<int, object>();

            /// <summary>
            /// Still need default constructor
            /// </summary>
            internal RowItem()
            {

            }

            /// <summary>
            /// ISeralializable constructor.
            /// </summary>
            internal RowItem(SerializationInfo info, StreamingContext context)
            {
                // Reset the property value using the GetValue method.
                _maxCol = (int) info.GetValue("_maxCol", typeof(int));
                ColItems = (Dictionary<int, object>) info.GetValue("ColItems", typeof(Dictionary<int, object>));
            }

            internal object GetItem(int colIndex)
            {
                if (ColItems.ContainsKey(colIndex))
                {
                    return ColItems[colIndex];
                }
                return null;
            }

            internal void InsertItem(int colIndex, object value)
            {
                if (ColItems.ContainsKey(colIndex))
                {
                    ColItems[colIndex] = value;
                }
                else
                {
                    _maxCol = Math.Max(_maxCol, colIndex);
                    ColItems.Add(colIndex, value);
                }
            }            

            internal object[] GetAllItems()
            {
                var values = new List<object>();
                for(int index = 1; index <= _maxCol; index++)
                {
                    values.Add(GetItem(index));
                }
                return values.ToArray();
            }

            /// <summary>
            /// ISerializable
            /// </summary>
            public void GetObjectData(SerializationInfo info, StreamingContext context)
            {
                // Use the AddValue method to specify serialized values.
                info.AddValue("_maxCol", _maxCol, typeof(int));
                info.AddValue("ColItems", ColItems, typeof(Dictionary<int, object>));
            }
        }

    }
}