using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using word=Microsoft.Office.Interop.Word;

namespace Word_AddIn_Test
{
   public class tables
    {
       private word.Shapes shapes;
       public tables(word.Shapes _shapes)
       {
           shapes = _shapes;
       }

       public int GetAllShapesCount()
       {
           return shapes.Count;
       }

       public int GetAllTableCount()
       {
           foreach (word.Shape item in shapes)
           {

           }

           return 0;
       }

    }
}
