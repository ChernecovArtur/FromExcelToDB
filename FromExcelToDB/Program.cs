using System;
using System.Data;
using System.Data.SqlClient;
using ExcelDataReader;
using System.Text;

public class ExcelToDatabase
{
    private string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=Objects_Identification;Integrated Security=True";

    public void ReadExcelAndSaveToDatabase(string filePath)
    {
        // Открываем файл Excel с помощью библиотеки ExcelDataReader
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding("Windows-1251") // Укажите другую кодировку, например, UTF-8
            }))
            {
                // Читаем данные из Excel-документа
                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true // Указываем, что первая строка содержит заголовки столбцов
                    }
                });

                // Получаем нужную таблицу из DataSet (может потребоваться изменение индекса, в зависимости от структуры вашего Excel-документа)
                DataTable table = result.Tables[0];

                // Подключаемся к базе данных
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Проходимся по каждой строке в таблице и записываем данные в базу данных
                    foreach (DataRow row in table.Rows)
                    {
                        string column1Value = row["A"].ToString(); // Имя столбца в Excel
                        string column2Value = row["D"].ToString(); // Имя столбца в Excel
                        string column3Value = row["AB"].ToString();
                        string column4Value = row["AE"].ToString();
                        string column5Value = row["I"].ToString();

                        // Выполните соответствующую операцию вставки в базу данных
                        string insertQuery = $"INSERT INTO Objects_Identification (unique_id, object_serial_number, object_address, object_subdivision, object_type) VALUES ({column1Value}, '{column2Value}', '{column3Value}', '{column4Value}', '{column5Value}')";
                        SqlCommand command = new SqlCommand(insertQuery, connection);
                        command.ExecuteNonQuery();
                    }

                    connection.Close();
                }
            }
        }
    }

    static void Main()
    {
        string excelFilePath = "C:\\Users\\User\\Desktop\\выгрузка 22.06.xlsm";

        ExcelToDatabase excelToDatabase = new ExcelToDatabase();
        excelToDatabase.ReadExcelAndSaveToDatabase(excelFilePath);

    }


}
