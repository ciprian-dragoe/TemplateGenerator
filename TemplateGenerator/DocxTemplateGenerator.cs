using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.Data;
using System.IO;

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
                        throw new Exception(string.Format("Cuvantul de inlocuit \"{0}\" din fisierul word nu are un corespondent ca si coloana in fisierul DBF.", column.ColumnName));
                    }
                    
                }
                        
            }
        }

        public void saveNewDocXfile(string path)
        {
            try
            { 
                generatedTemplate.SaveAs(path + "\\" + newGeneratedTemplateName);
                
            }
            catch (System.IO.IOException)
            {
                throw new Exception(string.Format("Fisierul \"{0}\" este deja deschis si nu poate sa fie modificat.", path + newGeneratedTemplateName));
            }
            ((IDisposable)generatedTemplate).Dispose();
        }
    }
}
