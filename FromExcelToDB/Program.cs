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

    //метод считывает данные из файла Excel и сохраняет их в базе данных.
    //параметры:
    //FilePath: путь к файлу Excel, из которого нужно считать данные.
    //для работы функции необходимы библиотеки ClosedXML и System.Data.SqlClient.

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

                var firstRow = worksheet.Row(1); //запись значений первой строки таблицы

                //gодготовка словаря для хранения соответствия заголовков столбцов и их индексов
                var columnHeaders = new Dictionary<string, int>();

                //заполнение словаря соответствия заголовков и нумерации столбцов
                foreach (var cell in firstRow.Cells())
                {
                    string columnHeader = cell.Value.ToString();
                    int columnIndex = cell.Address.ColumnNumber;
                    columnHeaders[columnHeader] = columnIndex;
                }

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    //получение значений для каждого столбца на основе заголовков
                    string column1Header = "Unique_Id";
                    string column2Header = "Название";
                    string column3Header = "Серийный_номер";
                    string column4Header = "Размещение_адрес";
                    string column5Header = "Размещение_помещение";
                    string column6Header = "Тип_КЕ";

                    string column1Value = row.Cell(columnHeaders[column1Header]).Value.ToString(); // id
                    string column2Value = row.Cell(columnHeaders[column2Header]).Value.ToString(); // name
                    string column3Value = row.Cell(columnHeaders[column3Header]).Value.ToString(); // serial number
                    string column4Value = row.Cell(columnHeaders[column4Header]).Value.ToString(); // address
                    string column5Value = row.Cell(columnHeaders[column5Header]).Value.ToString(); // subdivision
                    string column6Value = row.Cell(columnHeaders[column6Header]).Value.ToString(); // type_KE


                    //проверка наличия уникального ключа в БД
                    string selectQuery = $"SELECT COUNT(*) FROM registrated_objects WHERE unique_id = {column1Value}";
                    SqlCommand selectCommand = new SqlCommand(selectQuery, connection);
                    int existingCount = (int)selectCommand.ExecuteScalar();


                    //при повторе, значение с ключом обновляем, наче добавляем новое значение в таблицу
                    if (existingCount > 0)
                    {
                        // Если запись уже существует, обновляем значения остальных полей
                        string updateQuery = $"UPDATE registrated_objects SET object_name = N'{column2Value}', object_type = N'{column6Value}', object_serial_number = N'{column3Value}', object_address = N'{column4Value}', object_subdivision = N'{column5Value}' WHERE unique_id = {column1Value}";
                        SqlCommand updateCommand = new SqlCommand(updateQuery, connection);
                        updateCommand.ExecuteNonQuery();
                    }
                    else
                    {
                        // Если запись не существует, выполняем операцию вставки
                        string insertQuery = $"INSERT INTO registrated_objects (unique_id, object_name, object_type, object_serial_number, object_address, object_subdivision) VALUES ({column1Value}, N'{column2Value}', N'{column6Value}', N'{column3Value}', N'{column4Value}', N'{column5Value}')";
                        SqlCommand insertCommand = new SqlCommand(insertQuery, connection);
                        insertCommand.ExecuteNonQuery();
                    }
                }


               /* // выполнение операции вставки в БД
                string insertQuery = $"INSERT INTO registrated_objects (unique_id, object_name, object_type, object_serial_number, object_address, object_subdivision) VALUES ({column1Value}, N'{column2Value}', N'{column6Value}', N'{column3Value}', N'{column4Value}', N'{column5Value}')";

                    SqlCommand command = new SqlCommand(insertQuery, connection);
                    command.ExecuteNonQuery();
                }*/

                //Console.WriteLine(firstRow.Cell("Unique_Id").Value.ToString());

                /*foreach (var row in worksheet.RowsUsed().Skip(1))
                {

                    
                    // Получение значений для каждого столбца на основе заголовков
                    string column1Value = row.Cell("Unique_Id").Value.ToString(); // id
                    string column2Value = row.Cell("Название").Value.ToString(); // name
                    string column3Value = row.Cell("Серийный_номер").Value.ToString(); // serial number
                    string column4Value = row.Cell("Размещение_адрес").Value.ToString(); // address
                    string column5Value = row.Cell("Размещение_помещение").Value.ToString(); // subdivision
                    string column6Value = row.Cell("Тип_КЕ").Value.ToString(); // type_KE

                    string insertQuery = $"INSERT INTO registrated_objects (unique_id, object_name, object_type, object_serial_number, object_address, object_subdivision) VALUES ({column1Value}, N'{column2Value}', N'{column6Value}', N'{column3Value}', N'{column4Value}', N'{column5Value}')";

                    SqlCommand command = new SqlCommand(insertQuery, connection);
                    command.ExecuteNonQuery();
                }*/

                    //проходимся по каждой строке в таблице и записываем данные в базу данных
                    /*foreach (var row in worksheet.RowsUsed())
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
                    }*/




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
