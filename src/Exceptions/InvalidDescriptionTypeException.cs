using System;

namespace APIExcelValidator.Exceptions
{
    public class InvalidDescriptionTypeException : Exception
    {
        public InvalidDescriptionTypeException(string message) : base(message)
        {
        }

        public InvalidDescriptionTypeException() : base()
        {
        }
    }
}