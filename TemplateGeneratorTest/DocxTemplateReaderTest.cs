using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateGenerator;
using System.Data;
using System.Collections.Generic;

namespace TemplateGeneratorTest
{
    [TestClass]
    public class DocxTemplateReaderTest
    {
        private string pathToWorkingFiles = @"d:\CLOUD-SYNC\Cipi\Dropbox\Projects\TemplateGenerator\WorkingFiles\";

        [TestMethod]
        public void KeywordsReadCorrectly()
        {
            DocxTemplateReader readDocx = new DocxTemplateReader(pathToWorkingFiles + "Contract Dolce Sport Mansat -1.docx");
            List<string> expectedList = new List<string>();
            expectedList.Add("mail");
            expectedList.Add("prenume");
            expectedList.Add("nume");

            DataTable actualTable = readDocx.GetKeywords();
            List<string> actualList = new List<string>();

            foreach (DataColumn column in actualTable.Columns)
            {
                actualList.Add(column.ColumnName);
            }

            Assert.AreEqual(expectedList[0], actualList[0]);
            Assert.AreEqual(expectedList[1], actualList[1]);
            Assert.AreEqual(expectedList[2], actualList[2]);
        }

        [TestMethod]
        public void MultipleIdenticalKeywordsReadCorrectly()
        {
            DocxTemplateReader readDocx = new DocxTemplateReader(pathToWorkingFiles + "Contract Dolce Sport Mansat -2.docx");
            List<string> expectedList = new List<string>();
            expectedList.Add("nume");
            expectedList.Add("mail");
            expectedList.Add("prenume");

            DataTable actualTable = readDocx.GetKeywords();
            List<string> actualList = new List<string>();

            foreach (DataColumn column in actualTable.Columns)
            {
                actualList.Add(column.ColumnName);
            }

            Assert.AreEqual(expectedList[0], actualList[0]);
            Assert.AreEqual(expectedList[1], actualList[1]);
            Assert.AreEqual(expectedList[2], actualList[2]);
        }

        [TestMethod]
        public void CaseSensitiveKeywordsReadCorrectly()
        {
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

            Assert.AreEqual(expectedList[0], actualList[0]);
            Assert.AreEqual(expectedList[1], actualList[1]);
            Assert.AreEqual(expectedList[2], actualList[2]);
        }
    }
}
