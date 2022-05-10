using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace Excelerate
{
    public class SharedStrings
    {
        private object _lockOb = new Object();
        private List<SharedString> _sharedStrings = new List<SharedString>();

        public SharedStrings()
        {

        }

        public int AddString(string newString)
        {
            lock(_lockOb)
            {
                var ss = _sharedStrings.FirstOrDefault(p=>p.Value == newString);
                if (ss != null)
                {
                    ss.References++;
                    return _sharedStrings.IndexOf(ss);
                }
                else
                {
                    var newSS = new SharedString(newString);
                    _sharedStrings.Add(newSS);
                    return _sharedStrings.Count - 1;
                }
            }
        }

        public string GetString(int stringIndex)
        {
            return _sharedStrings[stringIndex].Value;
        }

        public void RemoveString(string newString)
        {
            lock(_lockOb)
            {
                var ss = _sharedStrings.FirstOrDefault(p=>p.Value == newString);
                if (ss != null)
                {
                    if (ss.References == 1)
                    {
                        _sharedStrings.Remove(ss);
                        return;
                    }
                    ss.References--;
                }
            }
        }

        public int GetStringReference(string strToTest)
        {
            lock(_lockOb)
            {
                var ss = _sharedStrings.FirstOrDefault(p=>p.Value == strToTest);
                if (ss == null)
                {
                    return -1;
                }
                return _sharedStrings.IndexOf(ss);
            }
        }


        internal string GenerateSharedStringsXml()
        {
            int totalCount = _sharedStrings.Sum(p=>p.References);
            int uniqueCount = _sharedStrings.Count;
            var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + Environment.NewLine;
            xml += $"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{totalCount}\" uniqueCount=\"{uniqueCount}\">";
            foreach(var ss in _sharedStrings)
            {
                if (string.IsNullOrEmpty(ss.Value))
                {
                    xml += $"<si><t/></si>";
                }
                else
                {
                    xml += $"<si><t>{HttpUtility.HtmlEncode(ss.Value)}</t></si>";
                }
            }
            xml += "</sst>";
            return xml;
        }

        private class SharedString
        {
            internal SharedString(string value) 
            {
                Value = value;
                References = 1;
            }
            public string Value { get; set; }
            public int References { get; set; }


            public override bool Equals(object obj)
            {
                if (obj is SharedString)
                {
                    var ss = obj as SharedString;
                    if (ss.Value == Value)
                    {
                        return true;
                    }
                }
                return false;
            }

            public override int GetHashCode()
            {
                return Value.GetHashCode();
            }
        }
    }
}