using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace OleDbXlsxHelper
{
   using DbHelper;
   public partial class COleDbXlsxHelper : CDbHelper
   {
      public static OleDbCommand ParamFiller(OleDbCommand cmd, object obj, List<string> lst_strParams) =>
            COleDbXlsxHelper.FillParams(cmd, COleDbXlsxHelper.GetObjToMap(obj, lst_strParams, "\t"));
   }
}