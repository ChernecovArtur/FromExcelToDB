using System;
using System.Data;
using System.Data.SqlClient;
using ExcelDataReader;
using System.Text;
using ClosedXML.Excel;
using Cake.Core.IO;

public class ExcelToDatabase
{
    private string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=Objects_Identification;Integrated Security=True";

    public void ReadExcelAndSaveToDatabase(string FilePath)
    {
        using (var workbook = new XLWorkbook(FilePath))
        {
            var worksheet = workbook.Worksheet(1);

            //подключение к базе данных
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                connection.InfoMessage += Connection_InfoMessage; //добавляем обработчик события InfoMessage для получения сообщений от базы данных

              /*  using (var command = new SqlCommand("SET NAMES UTF8", connection))
                {
                    command.ExecuteNonQuery();
                }*/

                //проходимся по каждой строке в таблице и записываем данные в базу данных
                foreach (var row in worksheet.RowsUsed())
                {
                    if (row.RowNumber() == 1)
                    {
                        continue; //пропустить первую строку с заголовками столбцов
                    }

                    string column1Value = row.Cell("A").Value.ToString(); //id
                    string column2Value = row.Cell("D").Value.ToString(); //name
                    string column3Value = row.Cell("AB").Value.ToString(); //serial number
                    string column4Value = row.Cell("AE").Value.ToString(); //address
                    string column5Value = row.Cell("AF").Value.ToString(); //subdivision
                    string column6Value = row.Cell("I").Value.ToString(); //type_KE

                    //выполнение операции вставки в бд
                    //string insertQuery = $"INSERT INTO registrated_objects (unique_id, object_serial_number, object_address, object_subdivision, object_type) VALUES ({column1Value}, N'{column2Value}', N'{column3Value}', N'{column4Value}', N'{column5Value}')";
                    string insertQuery = $"INSERT INTO registrated_objects (unique_id, object_name, object_type, object_serial_number, object_address, object_subdivision) VALUES ({column1Value}, N'{column2Value}', N'{column6Value}', N'{column3Value}', N'{column4Value}', N'{column5Value}')";

                    SqlCommand command = new SqlCommand(insertQuery, connection);
                    command.ExecuteNonQuery();
                }
                connection.Close();
            }
        }
    }
    private static void Connection_InfoMessage(object sender, SqlInfoMessageEventArgs e)
    {
        //обработка сообщений от базы данных
        Console.WriteLine(e.Message);
    }


    /* public void ReadExcelAndSaveToDatabase(string filePath)
     {
         // Открываем файл Excel с помощью библиотеки ExcelDataReader
         using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
         {
             using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
             {
                 FallbackEncoding = Encoding.GetEncoding("Windows-1251") 
             }))
             {
                 // Читаем данные из Excel-документа
                 DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                 {
                     ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                     {
                         UseHeaderRow = true 
                     }
                 });

                 DataTable table = result.Tables[0];

                 // Подключаемся к базе данных
                 using (SqlConnection connection = new SqlConnection(connectionString))
                 {
                     connection.Open();

                     // Проходимся по каждой строке в таблице и записываем данные в базу данных
                     foreach (DataRow row in table.Rows)
                     {
                         string column1Value = row["A"].ToString();
                         string column2Value = row["D"].ToString();
                         string column3Value = row["AB"].ToString();
                         string column4Value = row["AE"].ToString();
                         string column5Value = row["I"].ToString();

                         string insertQuery = $"INSERT INTO Objects_Identification (unique_id, object_serial_number, object_address, object_subdivision, object_type) VALUES ({column1Value}, '{column2Value}', '{column3Value}', '{column4Value}', '{column5Value}')";
                         SqlCommand command = new SqlCommand(insertQuery, connection);
                         command.ExecuteNonQuery();
                     }

                     connection.Close();
                 }
             }
         }
     }*/

    static void Main()
    {
        string excelFilePath = "C:\\Users\\User\\Desktop\\выгрузка 22.06.xlsm";

        ExcelToDatabase excelToDatabase = new ExcelToDatabase();
        excelToDatabase.ReadExcelAndSaveToDatabase(excelFilePath);

    }


}
