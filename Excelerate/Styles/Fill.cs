using System.IO;

namespace Excelerate.Styles
{
    public class Fill 
    {
        private FillBase _fill;

        public Fill(FillBase fill)
        {
            _fill = fill;
        }

        public FillBase InternalFill { get { return _fill; } }

        internal string GenerateXml()
        {
            return $"<fill>{_fill.GenerateFillXml()}</fill>";
        }

    }
}