using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using IronXL;

class Program
{
    static void Main()
    {
        // Set your Excel file path
        string excelFilePath = "D:\\Person.xlsx";

        // Set your MySQL connection string
        string connectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\Excel\\Excel.mdf;Integrated Security=True";

        // Set your MySQL table name
        string tableName = "Person";

        // Copy the entire worksheet to MySQL table
        CopyWorksheetToMySql(excelFilePath, connectionString, tableName);

        Console.WriteLine("Data successfully copied from Excel to MySQL.");
    }

    static void CopyWorksheetToMySql(string filePath, string connectionString, string tableName)
    {
        WorkBook workbook = WorkBook.Load(filePath);
        WorkSheet sheet = workbook.GetWorkSheet("Sheet1"); // Assuming the data is in the first sheet
        var c = workbook.WorkSheets.Count;
        
        // Create a DataTable to hold the data
        DataTable dataTable = sheet.ToDataTable();

        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Number", typeof(int));

        // Moving data from old columns to new columns
        foreach (DataRow row in dataTable.Rows)
        {
            row["Id"] = row["Column1"];
            row["Name"] = row["Column2"];
            row["Number"] = row["Column3"];
        }

        dataTable.Columns.Remove("Column1");
        dataTable.Columns.Remove("Column2");
        dataTable.Columns.Remove("Column3");

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

           
            InsertDataIntoMySqlTable(connection, tableName, dataTable);
        }
    }

   

    static void InsertDataIntoMySqlTable(SqlConnection connection, string tableName, DataTable dataTable)
    {
        using (SqlTransaction transaction = connection.BeginTransaction())
        {
            using (SqlCommand cmd = new SqlCommand("[sp_ManipulateData]"))
            {
                cmd.Connection = connection;
                cmd.Transaction = transaction;

                //set commandType as stored procedure
                cmd.CommandType = CommandType.StoredProcedure;

                SqlParameter parameter = new SqlParameter("@dtPerson", SqlDbType.Structured)
                {
                    TypeName = "dtPerson", // Replace with your table type name
                    Value = dataTable
                };
                parameter.TypeName = "dbo.dtPerson";
                parameter.SqlDbType = SqlDbType.Structured;
                cmd.Parameters.Add(parameter);

                if (dataTable.Rows.Count > 0)
                { 
                    cmd.ExecuteNonQuery();
                }


                transaction.Commit();
            }
        }
    }
}
