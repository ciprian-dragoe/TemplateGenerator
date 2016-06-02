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
            //var watch = System.Diagnostics.Stopwatch.StartNew();
            //ReadDBF.displayDBFcontents(@"c:\Users\Cipi\Desktop\SATNETTE.DBF");
            //watch.Stop();
            //var elapsedMs = watch.ElapsedMilliseconds;
            //Console.WriteLine("Creand baza de date: {0}", elapsedMs/1000);

            //WordReplacer.replaceParagraph(@"c:\Users\Cipi\Desktop\Test.docx", ":)");
            //WordReplacer.displayFoundText(@"c:\Users\Cipi\Desktop\Test.docx", "[");
            //DocxTemplateReader.replaceTextInParagraph(@"c:\Users\Cipi\Desktop\Contract Dolce Sport Mansat -1.docx", "TEST");

            //DocxTemplateGenerator templateGenerator = new DocxTemplateGenerator(@"c:\Users\Cipi\Desktop\Test.docx");

            string pathToWorkingFiles = @"d:\CLOUD-SYNC\Cipi\Dropbox\Projects\TemplateGenerator\WorkingFiles\";

            DocxTemplateReader readDocx = new DocxTemplateReader(pathToWorkingFiles + "Contract Dolce Sport Mansat -3.docx");
            List<string> expectedList = new List<string>();
            expectedList.Add("Nume");
            expectedList.Add("mail");
            expectedList.Add("prenume");
            expectedList.Add("nume");

            DataTable actualTable = readDocx.GetKeywords();
            List<string> actualList = new List<string>();

            foreach (DataColumn column in actualTable.Columns)
            {
                actualList.Add(column.ColumnName);
            }


            //while (dbfReader.currentIndex < dbfReader.totalNumberElements)
            //{
            //    DataTable readData = dbfReader.getRowAndAdvanceIndex(templateGenerator.columnNamesFromDBF);
            //    templateGenerator.replaceKeywordsInTemplate(readData);
            //    templateGenerator.saveNewDocXfile(@"c:\Users\Cipi\Desktop\TEMP\");
            //}

            //Console.ReadKey();
        }
    }
}
