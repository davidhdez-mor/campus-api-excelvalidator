using System;
using System.Data;
using APIExcelValidator.Exceptions;

namespace APIExcelValidator.Implementations
{
    public static class ValidateRangeField
    {
        public static bool ValidateDescriptionByRange(this DataTable table, string description, int fromColIndex,
            int toColindex,
            int fromRowIndex, int toRowIndex)
        {
            for (int col = fromColIndex; col < toColindex; col++)
            {
                for (int row = fromRowIndex; row < toRowIndex; row++)
                {
                    string valueFromField = table.Rows[row][col].ToString();
                    if (valueFromField.Length == 0)
                        throw new InvalidDescriptionTypeException($"Campo vacio: Error en la columna {col} fila {row + 2}");
                    if (!valueFromField.Equals(description))
                        throw new InvalidDescriptionTypeException($"Campo no coincide: Error en la columna {col} fila {row + 2}");
                }
            }

            return true;
        }

        public static bool ValidateNumberByRange(this DataTable table, int fromColIndex, int toColIndex, int fromRowIndex,
            int toRowIndex)
        {
            for (int col = fromColIndex; col <= toColIndex; col++)
            {
                for (int row = fromRowIndex; row <= toRowIndex; row++)
                {
                    try
                    {
                        float.Parse(table.Rows[row][col].ToString());
                    }
                    catch (FormatException ex)
                    {
                        throw new InvalidNumberTypeException($"El campo no es un numero: Error en la columna {col} fila {row + 2}");
                    }
                }
            }

            return true;
        }
    }
}