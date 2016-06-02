using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using System.Drawing;
using System.Data;

namespace TemplateGenerator
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
                Console.WriteLine("Fisierul \"{0}\" este deschis in alta aplicatie.", pathToFile);
                Console.ReadKey();
                Environment.Exit(1);
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

        /*
       public static void createDocX(string nameOfFile)
       {
           using (DocX wordDoc = DocX.Create(nameOfFile))
           {
               Paragraph firstParagraph = wordDoc.InsertParagraph();
               firstParagraph.Append("Hello World").Font(new FontFamily("Arial Black"));
               wordDoc.Save();
           }
       }

       public static void replaceParagraph(string pathToFile, string textThatReplacesParagraph)
       {
           using (DocX wordDoc = DocX.Load(pathToFile))
           {
               foreach(Bookmark bookmark in wordDoc.Bookmarks)
               {
                   Console.WriteLine("Bookmark gasit: {0}", bookmark.Name);
                   wordDoc.Bookmarks[bookmark.Name].SetText(textThatReplacesParagraph);
               }

               wordDoc.Save();
           }
       }

       public static void replaceText(string pathToFile, string textToSearch, string textToReplace)
       {
           using (DocX wordDoc = DocX.Load(pathToFile))
           {
               wordDoc.ReplaceText(textToSearch, textToReplace);

               wordDoc.Save();
           }
       }

       public static void displayFoundText(string pathToFile, string textToSearch)
       {
           using (DocX wordDoc = DocX.Load(pathToFile))
           {
               List<int> gasitPozitii = wordDoc.FindAll(textToSearch);
               foreach (int pozitie in gasitPozitii)
               {
                   Console.WriteLine("Am gasit la linia: {0}", pozitie);
               }
           }
       }

       public static void replaceTextInParagraph(string pathToFile, string textToReplaceWith)
       {
           using (DocX wordDoc = DocX.Load(pathToFile))
           {
               foreach (Paragraph paragraph in wordDoc.Paragraphs)
               {
                   Console.WriteLine("_____________________________________");
                   Console.WriteLine(paragraph.Text);
                   List<int> gasitPozitiiDelimitatorInceput = paragraph.FindAll("{[");
                   List<int> gasitPozitiiDelimitatorSfarsit = paragraph.FindAll("]}");
                   for (int i = gasitPozitiiDelimitatorInceput.Count-1; i >=0 ; i--)   // count backwards because if text is replaced from the beginning it can change the position of the "gasitPozitiiDelimitatorInceput"
                   {
                       var start = gasitPozitiiDelimitatorInceput[i];
                       var stop = gasitPozitiiDelimitatorSfarsit[i];
                       string textToSearch = paragraph.Text.Substring(start, stop - start + 2);
                       Console.WriteLine("Pozitie inceput: {0}", start);
                       Console.WriteLine("Pozitie sfarsit: {0}", stop);
                       Console.WriteLine("Caut: {0}", textToSearch);
                       paragraph.ReplaceText(textToSearch, textToReplaceWith);
                   }
                   Console.WriteLine("*********************************DUPA*********************************");
                   Console.WriteLine(paragraph.Text);
                   //wordDoc.Save();
               }   // paragraph
               wordDoc.Save();
           }   // wordDoc
       }   // replaceTextInParagraph
       */
    }
}
