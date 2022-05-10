using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Excelerate
{
    public class WorksheetCollection : ICollection<IWorksheet>
    {
        private List<IWorksheet> _sheets = new List<IWorksheet>();
        public int Count => _sheets.Count;

        public bool IsReadOnly => false;

        public IWorksheet this[string sheetName]
        {
            get
            {
                try
                {
                    return _sheets.FirstOrDefault(p => p.Name == sheetName);
                }
                catch(Exception)
                {
                    throw new Exception($"Worksheet {sheetName} not found in worksheets collection.");
                }
            }
        }

        public IWorksheet this[int index]
        {
            get
            {
                return _sheets[index];
            }
        }

        public void Add(IWorksheet item)
        {
            _sheets.Add(item);
        }

        public void Clear()
        {
            _sheets.Clear();
        }

        public bool Contains(IWorksheet item)
        {
            return _sheets.Contains(item);
        }

        public void CopyTo(IWorksheet[] array, int arrayIndex)
        {
            _sheets.CopyTo(array, arrayIndex);
        }

        public IEnumerator<IWorksheet> GetEnumerator()
        {
            return _sheets.GetEnumerator();
        }

        public bool Remove(IWorksheet item)
        {
            return _sheets.Remove(item);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _sheets.GetEnumerator();
        }
    }
}