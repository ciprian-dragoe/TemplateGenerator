using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TemplateGenerator
{
    public class TemplateManager
    {
        static void Main(string[] args)
        {
            DocxTemplateGenerator templateGenerator = new DocxTemplateGenerator(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\Contract Dolce Sport Mansat -7.docx");
            DBFreader dbfReader = new DBFreader(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\SATNETTE.DBF");

            while (dbfReader.currentIndex < dbfReader.totalNumberElements)
            {
                DataRow readData = dbfReader.getRowAndAdvanceIndex(templateGenerator.columnNamesFromDBF);
                templateGenerator.replaceKeywordsInTemplate(readData);
                templateGenerator.saveNewDocXfile(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\GeneratedTemplates\");
            }
            Console.WriteLine("Finished.");
            Console.ReadKey();
        }
    }
}
