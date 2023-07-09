using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace OleDbXlsxHelper
{
   public partial class COleDbXlsxHelper
   {
      public static bool Load<T>(List<T> master, string strConnection, string strSelectSQL, Action<List<T>, IDataReader> DoAdd, ref string strError)
      {
         bool blnRetVal = true;

         try
         {
            using (OleDbConnection conn = new OleDbConnection(strConnection))
            {
               conn.Open();

               using (OleDbDataReader rdr = (new OleDbCommand(strSelectSQL, conn)).ExecuteReader())
               {
                  while (rdr.Read())
                  {
                     DoAdd(master, rdr);
                  }
                  //
                  rdr.Close();
               }
               //
               conn.Close();
            }
         }
         catch (Exception exc)
         {
            blnRetVal = false;
            strError = exc.Message;
         }

         return blnRetVal;
      }
   }
}