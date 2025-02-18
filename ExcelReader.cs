using Microsoft.Data.SqlClient;
using System.Data;
using System.Data.OleDb;

public class ExcelReader
{
    public DataTable ReadExcel(string filePath)
    {
        var connectionString = GetConnectionString(filePath);
        var dataTable = new DataTable();

        using (var connection = new OleDbConnection(connectionString))
        {
            connection.Open();
            var sheetName = GetSheetName(connection);
            //var command = new OleDbCommand($"SELECT * FROM [{sheetName}]", connection);
            var command = new OleDbCommand($"SELECT * FROM [{sheetName}]", connection);
            var adapter = new OleDbDataAdapter(command);

            adapter.Fill(dataTable);

            //Remove the first 7 rows
            for (int i = 0; i < 7; i++)
            {
                if (dataTable.Rows.Count > 0)
                {
                    dataTable.Rows[0].Delete();
                }
                dataTable.AcceptChanges();
            }

            // Merge the values of the first two rows to create new column names
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string value1 = dataTable.Rows[0][i] != DBNull.Value ? dataTable.Rows[0][i].ToString() : "";
                string value2 = dataTable.Rows[1][i] != DBNull.Value ? dataTable.Rows[1][i].ToString() : i.ToString();
                string newColumnName = $"{value1} {value2}".Trim();
                dataTable.Columns[i].ColumnName = newColumnName;
            }

            // Remove the first two rows as they are now used as column names
            dataTable.Rows[0].Delete();
            dataTable.Rows[1].Delete();
            dataTable.AcceptChanges();
        }

        return dataTable;
    }
    
    private string GetConnectionString(string filePath)
    {
        //return Path.GetExtension(filePath).ToLower() switch
        //{
        //    ".xlsx" => $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'",
        //    ".xls" => $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'",
        //    _ => throw new NotSupportedException("File format not supported")
        //};

        return Path.GetExtension(filePath).ToLower() switch
        {
            ".xlsx" => $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'",
            ".xls" => $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'",
            _ => throw new NotSupportedException("File format not supported")
        };
    }

    private string GetSheetName(OleDbConnection connection)
    {
        var schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        return schemaTable.Rows[0]["TABLE_NAME"].ToString();
    }
}