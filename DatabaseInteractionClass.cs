using System;
using System.Collections.Generic;
using System.Text;
using System;
using System.Data.SqlClient;
using System.Linq;
using OfficeOpenXml; 
using System.Security.Cryptography.X509Certificates;

namespace ProgettoExcelClaudia
{
    public class DatabaseInteraction
    {
        public async Task DatabaseConnection()
        {
            string connectionString = "Server=insertservername;Database=insertdbname;User ID=insertuserdidhere;Password=insertpassword;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    string sqlQuery = "SELECT * FROM AEREO, AEROPORTO, VOLO";
                    
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string colonna1 = reader.GetString(0);
                                int colonna2 = reader.GetInt32(1);
                               
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Errore di connessione al database: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                }
            }
        }




        public async Task ExcelFileCreation ()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage excelPackage = new ExcelPackage();
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("info_voli_germania_italia");
                DatabaseInteraction databaseInteraction = new DatabaseInteraction();

                Console.WriteLine("Inserisci il percorso dove vuoi salvarlo: ");
                string filePath = Console.ReadLine();


                FileInfo excelFile = new FileInfo(filePath);



                excelPackage.SaveAs(excelFile);

                Console.WriteLine("Bravo! File creato.");

                // Scrivi l'intestazione delle colonne
                worksheet.Cells["A1"].Value = "IDVOLO";
                worksheet.Cells["B1"].Value = "ORAARR";
                worksheet.Cells["C1"].Value = "CITTAPART";
                worksheet.Cells["D1"].Value = "CITTAARR";
                worksheet.Cells["E1"].Value = "TIPOAEREO";

               
                // worksheet.Cells["F1"].Value = "NUMPASSEGGERI";

                // Esecuzione della query per ottenere i dati dei voli
                string sqlQuery = "SELECT IDVOLO, ORAARR, CITTAPART, CITTAARR, TIPOAEREO FROM VOLO WHERE CITTAARR = 'DEU' AND CITTAPART = 'ITA'";

                // Connessione al database e popolamento dei dati nel file Excel
                string connectionString = "Server=insertserverhere;Database=insertdbhere;User ID=insertuseridhere;Password=insertpasswordhere;";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            int row = 2; // Inizia dalla riga 2, lasciando la riga 1 per l'intestazione

                            // Leggi i dati dal SqlDataReader e scrivi nelle celle corrispondenti nel file Excel
                            while (reader.Read())
                            {
                                worksheet.Cells[$"A{row}"].Value = reader["IDVOLO"].ToString();
                                worksheet.Cells[$"B{row}"].Value = reader["ORAARR"].ToString();
                                worksheet.Cells[$"C{row}"].Value = reader["CITTAPART"].ToString();
                                worksheet.Cells[$"D{row}"].Value = reader["CITTAARR"].ToString();
                                worksheet.Cells[$"E{row}"].Value = reader["TIPOAEREO"].ToString();
                                //worksheet.Cells[$"F1{row}"].Value = reader["NUMPASSEGGERI"].ToString();

                                row++;
                            }
                        }
                    }
                }

              
              
            }
            catch (Exception ex)
            {
                Console.WriteLine("Errore durante la creazione del file Excel: " + ex.Message);
            }
        }

    }
}
