using System;

namespace ExcelOrderAddIn.Exceptions
{
    public class ProgrammerErrorException : Exception
    {
        public ProgrammerErrorException()
        {
        }

        public ProgrammerErrorException(string message)
            : base(message)
        {
        }

        public ProgrammerErrorException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
