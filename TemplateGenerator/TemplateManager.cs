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
        public static int Main(string[] args)
        {
            if (!File.Exists(args[0]))
            {
                throw new Exception("Nu ati dat o locatie valida pentru fisierul dbf.");
            }

            if (!File.Exists(args[1]))
            {
                throw new Exception("Nu ati dat o locatie valida pentru fisierul docx.");
            }

            if (!Directory.Exists(args[2]))
            {
                throw new Exception("Locatia pentru salvarea fisierelor generate este invalida.");
            }

            DocxTemplateGenerator templateGenerator = new DocxTemplateGenerator(args[1]);
            DBFreader dbfReader = new DBFreader(args[0]);
            foreach (var dataRow in dbfReader.readRows(templateGenerator.columnNamesFromDBF))
            {
                templateGenerator.replaceKeywordsInTemplate(dataRow);
                templateGenerator.saveNewDocXfile(args[2]);
            }
            return 0;
        }
    }
}
