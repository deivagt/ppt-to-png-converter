using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using System.IO;

namespace ppt_converter
{
    class converter
    {
        public void toPng(string inputFile)
        {
            Presentation pptPresentation = null;
            try
            {
               

                Application Application = new Application();
                pptPresentation = Application.Presentations
                .Open(inputFile, MsoTriState.msoFalse, MsoTriState.msoFalse
                , MsoTriState.msoFalse);
                
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                return;
            }
            String outputDirectory = Directory.GetCurrentDirectory() + "\\output";
            createDirectory(outputDirectory);
            int counter = 1;
            foreach (_Slide slide in pptPresentation.Slides)
            {
                string slideName = outputDirectory + "\\slide" + counter + ".png";
                slide.Export(slideName, "png", 1920, 1080);
                counter++;
            }
           


        }

        void createDirectory(String path)
        {
            
            try
            {
               
                if (Directory.Exists(path))
                {
                    Console.WriteLine("That path exists already.");
                    return;
                }

                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));

                
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            
        }
    }
}
