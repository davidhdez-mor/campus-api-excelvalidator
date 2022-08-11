using System;

namespace APIExcelValidator.Exceptions
{
    public class InvalidNumberTypeException : FormatException
    {
        
        public InvalidNumberTypeException(string message) : base(message)
        {
        }
    }
}