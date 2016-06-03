using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateGenerator;
using System.Data;
using Novacode;

namespace TemplateGeneratorTest
{
    [TestClass]
    public class DocxTemplateGeneratorTest
    {
        private string pathToWorkingFiles = @"d:\CLOUD-SYNC\Cipi\Qsync\6_PROJECTS\TemplateGenerator\WorkingFiles\";

        [TestMethod]
        public void ReplaceNumeStradaCorrect()
        {
            DocxTemplateGenerator generator = new DocxTemplateGenerator(pathToWorkingFiles + "Contract Dolce Sport Mansat -5.docx");
            DBFreader reader = new DBFreader(pathToWorkingFiles +"SATNETTE.DBF");
            string expected = "Acesta este un test cu I.I. MARIN IONEL si STEZII FN.Se pare ca merge sa schimbe bine numele si prenumele.";

            DataRow newValues = reader.getRowAndAdvanceIndex(generator.columnNamesFromDBF);
            generator.replaceKeywordsInTemplate(newValues);
            generator.saveNewDocXfile(pathToWorkingFiles + "GeneratedTemplates\\");

            using (DocX generatedTemplate = DocX.Load(pathToWorkingFiles + @"\GeneratedTemplates\I.I. MARIN IONEL.docx"))
            {
                string actual = "";
                foreach (Paragraph paragraf in generatedTemplate.Paragraphs)
                {
                    actual = actual + paragraf.Text;
                }
                Assert.AreEqual(expected, actual);
            }   // using
        }   // ReplaceNumeStradaCorrect

        [TestMethod]
        public void ReplaceNumeEmptyTELMANCorrect()
        {
            DocxTemplateGenerator generator = new DocxTemplateGenerator(pathToWorkingFiles + "Contract Dolce Sport Mansat -6.docx");
            DBFreader reader = new DBFreader(pathToWorkingFiles + "SATNETTE.DBF");
            string expected = "Acesta este un test cu I.I. MARIN IONEL si .Se pare ca merge sa schimbe bine numele si prenumele.";

            DataRow newValues = reader.getRowAndAdvanceIndex(generator.columnNamesFromDBF);
            generator.replaceKeywordsInTemplate(newValues);
            generator.saveNewDocXfile(pathToWorkingFiles + "GeneratedTemplates\\");

            using (DocX generatedTemplate = DocX.Load(pathToWorkingFiles + @"\GeneratedTemplates\I.I. MARIN IONEL.docx"))
            {
                string actual = "";
                foreach (Paragraph paragraf in generatedTemplate.Paragraphs)
                {
                    actual = actual + paragraf.Text;
                }
                Assert.AreEqual(expected, actual);
            }   // using
        }   // ReplaceNumeEmptyTELMANCorrect

        [TestMethod]
        public void Generate2Templates()
        {
            DocxTemplateGenerator generator = new DocxTemplateGenerator(pathToWorkingFiles + "Contract Dolce Sport Mansat -5.docx");
            DBFreader reader = new DBFreader(pathToWorkingFiles + "SATNETTE.DBF");
            string expected1 = "Acesta este un test cu I.I. MARIN IONEL si STEZII FN.Se pare ca merge sa schimbe bine numele si prenumele.";
            string expected2 = "Acesta este un test cu BANCA COOP PROGRESUL Sib si P-TA A. VLAICU.Se pare ca merge sa schimbe bine numele si prenumele.";

            DataRow newValues = reader.getRowAndAdvanceIndex(generator.columnNamesFromDBF);
            generator.replaceKeywordsInTemplate(newValues);
            generator.saveNewDocXfile(pathToWorkingFiles + "GeneratedTemplates\\");

            newValues = reader.getRowAndAdvanceIndex(generator.columnNamesFromDBF);
            generator.replaceKeywordsInTemplate(newValues);
            generator.saveNewDocXfile(pathToWorkingFiles + "GeneratedTemplates\\");

            using (DocX generatedTemplate = DocX.Load(pathToWorkingFiles + @"\GeneratedTemplates\I.I. MARIN IONEL.docx"))
            {
                string actual = "";
                foreach (Paragraph paragraf in generatedTemplate.Paragraphs)
                {
                    actual = actual + paragraf.Text;
                }
                Assert.AreEqual(expected1, actual);
            }   // using

            using (DocX generatedTemplate = DocX.Load(pathToWorkingFiles + @"\GeneratedTemplates\BANCA COOP PROGRESUL Sib.docx"))
            {
                string actual = "";
                foreach (Paragraph paragraf in generatedTemplate.Paragraphs)
                {
                    actual = actual + paragraf.Text;
                }
                Assert.AreEqual(expected2, actual);
            }   // using
        }   // Generate2Templates
    }
}
