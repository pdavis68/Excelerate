using System;
using System.Collections.Generic;
using System.Web;

namespace Excelerate
{
    public class RichTextSharedString : SharedStringBase
    {
        private List<RichTextRun> _runs = new List<RichTextRun>();

        public void AddRun(RichTextRun run)
        {
            _runs.Add(run);
        }

        public void Clear()
        {
            _runs.Clear();
        }

        internal override string GenerateSharedStringXml()
        {
            string runsXml = String.Empty;
            foreach(var run in _runs)
            {
                runsXml += run.GenerateRichTextRunXml();
            }
            return $"<si>{runsXml}</si>";
        }
    }
}
