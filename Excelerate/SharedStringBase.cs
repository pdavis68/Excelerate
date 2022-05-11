using System;

namespace Excelerate
{
    public abstract class SharedStringBase
    {
        public int References { get; internal set; } = 1;
        internal abstract string GenerateSharedStringXml();
    }
}
