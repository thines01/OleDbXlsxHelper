using System.Collections.Generic;

namespace OleDbXlsxHelper
{
   using DbHelper;

   public partial class COleDbXlsxHelper : CDbHelper
   {
      public COleDbXlsxHelper(string strPrefixChar, string strTableName, List<string> lst_strColumns)
         : base(strPrefixChar, strTableName, lst_strColumns)
      {
      }
   }
}