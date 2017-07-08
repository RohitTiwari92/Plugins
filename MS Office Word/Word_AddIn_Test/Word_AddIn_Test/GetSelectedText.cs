using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_AddIn_Test
{
    public class GetSelectedText
    {
       private  Microsoft.Office.Interop.Word.Selection wordSelection;
        public GetSelectedText( Microsoft.Office.Interop.Word.Selection _wordSelection)
        {
            wordSelection=_wordSelection;
        }

        public string getSelectedText()
        {
             string selectText = string.Empty;
             if (wordSelection != null && wordSelection.Range != null)
             {
                 selectText = wordSelection.Text;
             }
             return selectText;
        }
 
    }
}
