namespace APIExcelValidator.Abstractions
{
    public interface IFieldValidator
    {
        public void ValidateDescription();
        public void ValidateNumbers();
    }
}