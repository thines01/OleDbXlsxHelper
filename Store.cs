using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;

namespace OleDbXlsxHelper
{
   using DbHelper;
   using DB_OLEDB_XLSX;
   using IDbMaster;
   
   public partial class COleDbXlsxHelper : CDbHelper
   {
      //private static Func<OleDbCommand, Object, List<string>, OleDbCommand> ParamFiller = (cmd, obj, lst_strParams) =>
        // COleDbXlsxHelper.FillParams(cmd, COleDbXlsxHelper.GetObjToMap(obj, lst_strParams, "\t"));
      private static string _strOutFile { get; set; }
      private static List<string> _lst_strParams { get; set; }

      public bool Store(IDbMaster master, ref string strError)
      {
         bool blnRetVal = true;

         CDB_OLEDB_XLSX xlsx = new CDB_OLEDB_XLSX(_strOutFile);

         string strOutConnection = GetOutConnection(xlsx);

         // OleDbCommand ParamFiller(OleDbCommand cmd, object obj, List<string> lst_strParams) =>
           // COleDbXlsxHelper.FillParams(cmd, COleDbXlsxHelper.GetObjToMap(obj, lst_strParams, "\t"));

         OleDbParameter AddParam(string strParam) => new OleDbParameter(strParam, OleDbType.VarChar);

         using (OleDbConnection conn = new OleDbConnection(strOutConnection))
         {
            using (OleDbCommand cmdDelete = new OleDbCommand(_strDeleteSQL, conn))
            {
               using (OleDbCommand cmdInsert = new OleDbCommand(_strInsertSQL, conn))
               {
                  _lst_strParams.ForEach(s => DoAdd(cmdInsert, s, AddParam));

                  //blnRetVal = COleDbXlsxHelper.Store(master, cmdInsert, cmdDelete, _lst_strParams, conn, ParamFiller, ref strError);
               }
            }
         }

         return blnRetVal;
         //return COleDbXlsxHelper.Store(master, _strInsertSQL, _strDeleteSQL, _lst_strColumns, strOutConnection, ref strError);
      }
      /*
      public static bool Store<T>(List<T> master, string strInsertSQL, string strDeleteSQL, List<string> lst_strParams, string strConnection, ref string strError)
      {
         bool blnRetVal = true;

         string strSQL = strInsertSQL;
         
         try
         {
            using (OleDbConnection conn = new OleDbConnection(strConnection))
            {
               conn.Open();
               OleDbTransaction trans = conn.BeginTransaction();
               /////////////////////////////////////////////////////////////////
               // Delete is not available lin OLEDB
               // (new OleDbCommand(strDeleteSQL, conn, trans) { CommandTimeout = 300 }).ExecuteNonQuery();
               OleDbCommand cmd = (new OleDbCommand(strSQL, conn, trans) { CommandTimeout = 300 });
               lst_strParams.ForEach(s => DoAdd(cmd, s, AddParam));
               master.ForEach
               (
                  obj =>
                     COleDbXlsxHelper.FillParams(cmd, COleDbXlsxHelper.GetObjToMap(obj, lst_strParams, "\t"))
                     .ExecuteNonQuery()
               );
               trans.Commit();
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
      */
   }
}