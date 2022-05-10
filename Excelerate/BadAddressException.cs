using System;
using System.Runtime.Serialization;

namespace Excelerate
{

    [Serializable]
    internal class BadAddressException : Exception
    {
        public BadAddressException() { }
        public BadAddressException(string message) : base(message) { }
        public BadAddressException(string message, Exception inner) : base(message, inner) { }
        protected BadAddressException(
          SerializationInfo info,
          StreamingContext context) : base(info, context) { }
    }
}