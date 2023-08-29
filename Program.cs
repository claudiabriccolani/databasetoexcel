namespace ProgettoExcelClaudia
{
    internal class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Vuoi generare il tuo file Excel?");
            DatabaseInteraction databaseInteraction = new DatabaseInteraction();
            databaseInteraction.DatabaseConnection();

            databaseInteraction.ExcelFileCreation();

            
        }



        






    }
}