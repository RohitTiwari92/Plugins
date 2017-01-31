using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;

namespace PluginPowerPoint
{
    class AddTable
    {
        public void Add(Slides slides)
        {
            
            GetCount(slides);
           
        }

        void GetCount(Slides slides)
        {
            int i=1;
            foreach (Slide slide in slides)
            {
                
                Microsoft.Office.Interop.PowerPoint.Shape pptShape = slide.Shapes.AddTable(i, i);
                i++;
                if(i==7)
                {
                    i = 1;
                }
               
            }

        }
    }
}
