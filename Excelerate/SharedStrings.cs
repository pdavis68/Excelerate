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
        private List<SharedStringBase> _sharedStrings = new List<SharedStringBase>();

        public SharedStrings()
        {

        }
        public int AddString(SharedStringBase ssb)
        {
            int loc = _sharedStrings.Count;
            var existing = _sharedStrings.FirstOrDefault(p => p.Equals(ssb));
            if (existing != null)
            {
                existing.References++;
                return _sharedStrings.IndexOf(existing);
            }
            else
            {
                _sharedStrings.Add(ssb);
                return loc;
            }
        }

        public int AddString(string newString)
        {
            lock(_lockOb)
            {
                var ss = _sharedStrings.FirstOrDefault(p=>(p is SharedString) && ((p as SharedString).Value == newString));
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


        public int GetStringReference(string strToTest)
        {
            lock(_lockOb)
            {
                var ss = _sharedStrings.FirstOrDefault(p => (p is SharedString) && ((p as SharedString).Value == strToTest));
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
                if (ss is SharedString && string.IsNullOrEmpty((ss as SharedString).Value))
                {
                    xml += $"<si><t/></si>";
                }
                else
                {
                    xml += ss.GenerateSharedStringXml();
                }
            }
            xml += "</sst>";
            return xml;
        }

        public class SharedString : SharedStringBase
        {
            internal SharedString(string value) 
            {
                Value = value;
                References = 1;
            }
            public string Value { get; set; }

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

            internal override string GenerateSharedStringXml()
            {
                return $"<si><t>{HttpUtility.HtmlEncode(Value)}</t></si>";
            }

            public override int GetHashCode()
            {
                return Value.GetHashCode();
            }
        }
    }
}