using System.Data.OleDb;

namespace OleDbXlsxHelper
{
   using DB_OLEDB_XLSX;
   using DbHelper;

   public partial class COleDbXlsxHelper : CDbHelper
   {
      private static string GetOutConnection(CDB_OLEDB_XLSX xlsx) => xlsx.csb.ToString()
            .Replace("IMEX=1", "")
            .Replace("READONLY=TRUE", "READONLY=FALSE");

      private static OleDbConnection GetConn(string strConnect) => new OleDbConnection(strConnect);
   }
}