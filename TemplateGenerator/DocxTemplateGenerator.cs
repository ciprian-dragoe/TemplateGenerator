using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.Data;

namespace TemplateGenerator
{
    public class DocxTemplateGenerator
    {
        private DocxTemplateReader templateSource;
        public DataTable columnNamesFromDBF { get; private set; }
        private DocX generatedTemplate;
        private string newGeneratedTemplateName;

        public DocxTemplateGenerator(string pathToTemplateDocX)
        {
            templateSource = new DocxTemplateReader(pathToTemplateDocX);
            columnNamesFromDBF = templateSource.GetKeywords();
            try
            {
                columnNamesFromDBF.Columns.Add("NUME"); // "NUME" is always required because it generates the newly created document's name
            }
            catch (System.Data.DuplicateNameException)
            {
                // if the column name already exist then do not add it
            }
        }

        ~DocxTemplateGenerator()
        {
            if (generatedTemplate != null)
            {
                ((IDisposable)generatedTemplate).Dispose();
            }
        }

        public void replaceKeywordsInTemplate(DataRow readDataFromDBF)
        {
            newGeneratedTemplateName = (string)readDataFromDBF["NUME"] + ".docx";
            generatedTemplate = templateSource.wordTemplateSource.Copy();
            foreach (Paragraph paragraph in generatedTemplate.Paragraphs)
            {
                foreach (DataColumn column in readDataFromDBF.Table.Columns)
                {
                    try
                    {
                        string textToReplace = templateSource.patternStartKeyword + column.ColumnName + templateSource.patternEndKeyword;
                        paragraph.ReplaceText(textToReplace, (string)readDataFromDBF[column]);
                    }
                    catch (System.InvalidCastException)
                    {
                        // create Log & on console window
                        throw new Exception(string.Format("Column \"{0}\" does not exist in DBF file.", column.ColumnName));
                    }
                    
                }
                        
            }
        }

        public void saveNewDocXfile(string path)
        {
            try
            {
                generatedTemplate.SaveAs(path + newGeneratedTemplateName);
            }
            catch (System.IO.IOException)
            {
                Console.WriteLine("Fisierul \"{0}\" does deja deschis si nu poate sa fie modificat.", path + newGeneratedTemplateName);
                Console.ReadKey();
                Environment.Exit(3);
            }
            ((IDisposable)generatedTemplate).Dispose();
        }
    }
}
