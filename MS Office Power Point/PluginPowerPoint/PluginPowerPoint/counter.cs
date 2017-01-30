using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPT = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PluginPowerPoint
{
    class counter
    {
        
        public int detectSmartArt(Slides slides)
        {
            int count = 0;
            GetCount(slides, ref count);
            return count;
        }

         void GetCount(Slides slides,ref int count)
        {
            foreach (Slide slide in slides)
            {
                extractSlideInfo( slide,ref count);
            }
            
        }

        private void extractSlideInfo(Slide slide,ref int count)
        {
            foreach (PPT.Shape shape in slide.Shapes)
                extractInfoFromShape(shape,ref count);
           
        }

        private void extractInfoFromShape(PPT.Shape shape,ref int count)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                GetDataFromGroupItem(shape,ref count);
            }
            else if (shape.Type == MsoShapeType.msoSmartArt)
            {
                GetDataFromSmartArt(shape,ref count);
            }
            else if (shape.Type == MsoShapeType.msoPlaceholder && shape.PlaceholderFormat.ContainedType == MsoShapeType.msoSmartArt)
            {
                GetDataFromPlaceHolder(shape,ref count);
            }
            
        }


        private void GetDataFromPlaceHolder(PPT.Shape shape,ref int count)
        {
            try
            {
                GetDataFromSmartArt(shape, ref count);
            }
            catch
            {
               
                    try
                    {
                        GetDataFromGroupItem(shape, ref count);
                       
                    }
                    catch
                    {

                    }
                
            }
        }

        private void GetDataFromSmartArt(PPT.Shape shape,ref int count)
        {
            
                SmartArtNodes nodes = shape.SmartArt.AllNodes;

                foreach (SmartArtNode node in nodes)
                {
                    count++;
                }
           
        }

        private void GetDataFromGroupItem(PPT.Shape shape,ref int count)
        {
            foreach (PPT.Shape myShape in shape.GroupItems)
                extractInfoFromShape(myShape,ref count);
        }
    }
    
}
