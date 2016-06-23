using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using System.Drawing;
using System.Data;

namespace TemplateGenerator_Logic
{
    public class DocxTemplateReader
    {
        public string patternStartKeyword { get; private set; }
        public string patternEndKeyword { get; private set; }
        public DocX wordTemplateSource { get; private set; }

        public DocxTemplateReader(string pathToFile)
        {
            patternStartKeyword = "{[";
            patternEndKeyword = "]}";
            try
            {
                wordTemplateSource = DocX.Load(pathToFile);
            }
            catch (System.IO.IOException)
            {
                throw new Exception(string.Format("Fisierul \"{0}\" este deschis in alta aplicatie.", pathToFile));
            }
            
        }

        ~DocxTemplateReader()
        {
            if (wordTemplateSource != null)
            {
                ((IDisposable)wordTemplateSource).Dispose();
            }
        }

        // returns a dataTable with the column names matching the keywords from the docx file
        public DataTable GetKeywords()
        {
            DataTable returnValue = new DataTable();

            foreach (Paragraph paragraph in wordTemplateSource.Paragraphs)
            {
                List<int> foundDelimitorStartingPosition = paragraph.FindAll(patternStartKeyword);
                List<int> foundDelimitorEndingPosition = paragraph.FindAll(patternEndKeyword);
                for (int i = foundDelimitorStartingPosition.Count - 1; i >= 0; i--)   // count backwards because if text is replaced from the beginning it can change the position of the "gasitPozitiiDelimitatorInceput"
                {
                    var start = foundDelimitorStartingPosition[i];
                    var stop = foundDelimitorEndingPosition[i];
                    string textToSearch = paragraph.Text.Substring(start + patternStartKeyword.Length, stop - start - patternEndKeyword.Length);
                    try
                    {
                        returnValue.Columns.Add(textToSearch);
                    }
                    catch (System.Data.DuplicateNameException)
                    {
                        // if the column name already exist then do not add it
                    }
                }   // foundDelimitorStartingPosition
            }   // paragraph
            return returnValue;
        }
    }
}
