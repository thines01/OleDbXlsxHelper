using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace OleDbXlsxHelper
{
   using DbHelper;
   public partial class COleDbXlsxHelper : CDbHelper
   {
      public static OleDbParameter AddParam(string strParam) => new OleDbParameter(strParam, OleDbType.VarChar);

      public static Action<OleDbCommand, IEnumerable<string>> InitParams = (cmd, lstParams) =>
      {
         Action<string> DoAdd = strParam => cmd.Parameters.Add(new OleDbParameter(strParam, OleDbType.VarChar));
         lstParams.ToList().ForEach(DoAdd);
      };
   }
}