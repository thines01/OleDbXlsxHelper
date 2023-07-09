using System;
using System.Data.OleDb;

namespace OleDbXlsxHelper
{
   using DB_OLEDB_XLSX;

   public partial class COleDbXlsxHelper
   {
      public static string GetSheet1Name(string strInFile)
      {
         string strRetVal = "";

         try
         {
            CDB_OLEDB_XLSX xlsx = new CDB_OLEDB_XLSX(strInFile);

            using (OleDbConnection conn = new OleDbConnection(xlsx.csb.ToString()))
            {
               conn.Open();
               strRetVal = CDB_OLEDB_XLSX.GetSheet1Name(conn);
               conn.Close();
            }
         }
         catch (Exception exc)
         {
            System.Diagnostics.Debug.WriteLine("Exception: " + exc.Message);
         }

         return strRetVal;
      }
   }
}