using ExcelBulkInsert;

class Program
{
    static void Main()
    {
        try
        {
            //Console.WriteLine("Enter Excel file path:");
            //var filePath = Console.ReadLine();
            //var filePath = @"C:\Users\rijoan\Downloads\ok.xlsm";
            var filePath = @"C:\Users\rijoan\Downloads\Black Cat LLC 401(k) Plan_CensusDataImport_09.01.2024.xlsm";

            // Read Excel file
            var excelReader = new ExcelReader();
            var dataTable = excelReader.ReadExcel(filePath);

            // Generate table name (e.g., based on file name)
            var tableName = Path.GetFileNameWithoutExtension(filePath);
            tableName = tableName.Replace(" ", "").Replace(".", "_").Replace("(", "").Replace(")", "");

            // Initialize services
            var connectionString = "Server=SWD-RIJOAN-L;Database=TestDB;User ID=jahangir;Password=Baylor123;Integrated Security=true;TrustServerCertificate=True;";
            var sqlGenerator = new SqlTableGenerator();
            var dbService = new DatabaseService(connectionString);

            // Check if table exists
            if (dbService.TableExists(tableName))
            {
                Console.WriteLine($"Table '{tableName}' already exists. Do you want to overwrite it? (Y/N)");
                var response = Console.ReadLine();
                if (response?.ToUpper() == "Y")
                {
                    try
                    {
                        // Truncate the table
                        dbService.ExecuteCommand($"TRUNCATE TABLE {tableName}");
                        Console.WriteLine($"Table '{tableName}' truncated.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error executing command: {ex.Message}");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Operation canceled.");
                    return;
                }
            }
            else
            {
                // Generate and execute CREATE TABLE script if the table does not exist
                var createTableScript = sqlGenerator.GenerateCreateTableScript(dataTable, tableName);
                dbService.CreateTable(createTableScript);
            }

            // Bulk insert data
            dbService.BulkInsertData(tableName, dataTable);

            Console.WriteLine($"Table '{tableName}' created and data inserted successfully!");

            ////// Check if table exists
            ////if (dbService.TableExists(tableName))
            ////{
            ////    Console.WriteLine($"Table '{tableName}' already exists. Do you want to overwrite it? (Y/N)");
            ////    var response = Console.ReadLine();
            ////    if (response?.ToUpper() != "Y")
            ////    {
            ////        Console.WriteLine("Operation canceled.");
            ////        return;
            ////    }
            ////}

            ////// Generate and execute CREATE TABLE script
            ////var createTableScript = sqlGenerator.GenerateCreateTableScript(dataTable, tableName);
            ////dbService.CreateTable(createTableScript);

            ////// Bulk insert data
            ////dbService.BulkInsertData(tableName, dataTable);

            ////Console.WriteLine($"Table '{tableName}' created and data inserted successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
