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
            for (int i = 0; i < 1000000; i++)
            {

            }
            Console.WriteLine("ceva");
            Console.ReadKey();

            DocxTemplateGenerator templateGenerator = new DocxTemplateGenerator(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\Contract Dolce Sport Mansat -7.docx");
            DBFreader dbfReader = new DBFreader(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\SATNETTE.DBF");
            int count = 0;
            foreach (var dataRow in dbfReader.readRows(templateGenerator.columnNamesFromDBF))
            {
                templateGenerator.replaceKeywordsInTemplate(dataRow);
                Console.WriteLine(count++);
                //templateGenerator.saveNewDocXfile(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\GeneratedTemplates\");
            }
            Console.WriteLine("Finished.");
            Console.ReadKey();
        }
    }
}
