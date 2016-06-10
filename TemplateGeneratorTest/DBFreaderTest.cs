using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateGenerator;
using System.Data;
using System.Linq;

namespace TemplateGeneratorTest
{
    [TestClass]
    public class DBFreaderTest
    {
        private string defaultPathDBF = @"TestFiles\SATNETTE.DBF";

        [TestMethod]
        public void FirstRowReadCorrectly()
        {
            // prec
            DBFreader readFile = new DBFreader(defaultPathDBF);

            DataTable requiredColumns = new DataTable();
            requiredColumns.Columns.Add("NUME");

            DataTable expectedTable = requiredColumns.Clone();
            DataRow expectedRow = expectedTable.NewRow();
            expectedRow["NUME"] = "I.I. MARIN IONEL";
            expectedTable.Rows.Add(expectedRow);

            // act
            DataTable actual = expectedTable.Clone();

            actual.Rows.Add(readFile.readRows(requiredColumns).First().ItemArray);

            // expt
            Assert.AreEqual(expectedTable.Rows[0]["NUME"], actual.Rows[0]["NUME"]);
        }

        [TestMethod]
        public void FirstRowMultipleColumnsReadCorrectly()
        {
            // prec
            DBFreader readFile = new DBFreader(defaultPathDBF);

            DataTable requiredColumns = new DataTable();
            requiredColumns.Columns.Add("STRADA");
            requiredColumns.Columns.Add("NUME");

            DataTable expectedTable = requiredColumns.Clone();

            DataRow expectedRow = expectedTable.NewRow();
            expectedRow["NUME"] = "I.I. MARIN IONEL";
            expectedRow["STRADA"] = "STEZII FN";
            expectedTable.Rows.Add(expectedRow);

            // act
            DataTable actual = expectedTable.Clone();
            actual.Rows.Add(readFile.readRows(requiredColumns).First().ItemArray);

            // expt
            Assert.AreEqual(expectedTable.Rows[0]["NUME"], actual.Rows[0]["NUME"]);
            Assert.AreEqual(expectedTable.Rows[0]["STRADA"], actual.Rows[0]["STRADA"]);
        }

        [TestMethod]
        public void TwoRowsMultipleColumnsReadCorrectly()
        {
            // prec
            DBFreader readFile = new DBFreader(defaultPathDBF);

            DataTable requiredColumns = new DataTable();
            requiredColumns.Columns.Add("STRADA");
            requiredColumns.Columns.Add("NUME");

            DataTable expectedTable = requiredColumns.Clone();
            DataRow firstExpectedRow = expectedTable.NewRow();
            firstExpectedRow["NUME"] = "I.I. MARIN IONEL";
            firstExpectedRow["STRADA"] = "STEZII FN";
            expectedTable.Rows.Add(firstExpectedRow);

            DataRow secondExpectedRow = expectedTable.NewRow();
            secondExpectedRow["NUME"] = "BANCA COOP PROGRESUL Sib";
            secondExpectedRow["STRADA"] = "P-TA A. VLAICU";
            expectedTable.Rows.Add(secondExpectedRow);

            // act
            DataTable actual = expectedTable.Clone();
            actual.Rows.Add(readFile.readRows(requiredColumns).ElementAt(0).ItemArray);
            actual.Rows.Add(readFile.readRows(requiredColumns).ElementAt(1).ItemArray);

            // expt
            Assert.AreEqual(expectedTable.Rows[0]["NUME"], actual.Rows[0]["NUME"]);
            Assert.AreEqual(expectedTable.Rows[0]["STRADA"], actual.Rows[0]["STRADA"]);
            Assert.AreEqual(expectedTable.Rows[1]["NUME"], actual.Rows[1]["NUME"]);
            Assert.AreEqual(expectedTable.Rows[1]["STRADA"], actual.Rows[1]["STRADA"]);
        }
    }
}
