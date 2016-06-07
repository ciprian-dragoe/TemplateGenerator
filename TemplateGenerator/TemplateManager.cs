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
            string pathToWorkingFiles = @"d:\CLOUD-SYNC\Cipi\QSYNC\6_PROJECTS\TemplateGenerator\TemplateGeneratorTest\bin\Debug\TestFiles\";

            DocxTemplateGenerator generator = new DocxTemplateGenerator(Path.GetFullPath(pathToWorkingFiles + "Contract Dolce Sport Mansat -5.docx"));
            DBFreader reader = new DBFreader(pathToWorkingFiles + "SATNETTE.DBF");
            string expected1 = "Acesta este un test cu I.I. MARIN IONEL si STEZII FN.Se pare ca merge sa schimbe bine numele si prenumele.";
            string expected2 = "Acesta este un test cu BANCA COOP PROGRESUL Sib si P-TA A. VLAICU.Se pare ca merge sa schimbe bine numele si prenumele.";

            DataRow newValues = reader.readRows(generator.columnNamesFromDBF).First();
            generator.replaceKeywordsInTemplate(newValues);
            generator.saveNewDocXfile(pathToWorkingFiles + "GeneratedTemplates\\");

            newValues = reader.readRows(generator.columnNamesFromDBF).First();
            generator.replaceKeywordsInTemplate(newValues);
            generator.saveNewDocXfile(pathToWorkingFiles + "GeneratedTemplates\\");

            
            //DocxTemplateGenerator templateGenerator = new DocxTemplateGenerator(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\Contract Dolce Sport Mansat -7.docx");
            //DBFreader dbfReader = new DBFreader(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\SATNETTE.DBF");
            //int count = 0;
            //foreach (var dataRow in dbfReader.readRows(templateGenerator.columnNamesFromDBF))
            //{
            //    templateGenerator.replaceKeywordsInTemplate(dataRow);
            //    Console.WriteLine(count++);
            //    //templateGenerator.saveNewDocXfile(@"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\GeneratedTemplates\");
            //}
            //Console.WriteLine("Finished.");
            //Console.ReadKey();
        }
    }
}
