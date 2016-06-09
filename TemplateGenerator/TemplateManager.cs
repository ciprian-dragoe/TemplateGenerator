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
        public static IEnumerable<string> generateDocxTemplates(string pathToDBF, string pathToDocx, string pathToSaveDir)
        {
            if (!File.Exists(pathToDBF))
            {
                throw new Exception("Nu ati dat o locatie valida pentru fisierul dbf.");
            }

            if (!File.Exists(pathToDocx))
            {
                throw new Exception("Nu ati dat o locatie valida pentru fisierul docx.");
            }

            if (!Directory.Exists(pathToSaveDir))
            {
                throw new Exception("Locatia pentru salvarea fisierelor generate este invalida.");
            }

            DocxTemplateGenerator templateGenerator = new DocxTemplateGenerator(pathToDocx);
            DBFreader dbfReader = new DBFreader(pathToDBF);
            foreach (var dataRow in dbfReader.readRows(templateGenerator.columnNamesFromDBF))
            {
                templateGenerator.replaceKeywordsInTemplate(dataRow);
                templateGenerator.saveNewDocXfile(pathToSaveDir);
                yield return templateGenerator.newGeneratedTemplateName;
            }
        }
    }
}
