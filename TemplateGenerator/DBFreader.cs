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
    public class DBFreader
    {
        // This is the file header for a DBF. We do this special layout with everything
        // packed so we can read straight from disk into the structure to populate it
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        private struct DBFHeader
        {
            public byte version;
            public byte updateYear;
            public byte updateMonth;
            public byte updateDay;
            public Int32 numRecords;
            public Int16 headerLen;
            public Int16 recordLen;
            public Int16 reserved1;
            public byte incompleteTrans;
            public byte encryptionFlag;
            public Int32 reserved2;
            public Int64 reserved3;
            public byte MDX;
            public byte language;
            public Int16 reserved4;
        }
        private DBFHeader header;
        // This is the field descriptor structure. 
        // There will be one of these for each column in the table.
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        private struct FieldDescriptor
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 11)]
            public string fieldName;
            public char fieldType;
            public Int32 address;
            public byte fieldLen;
            public byte count;
            public Int16 reserved1;
            public byte workArea;
            public Int16 reserved2;
            public byte flag;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 7)]
            public byte[] reserved3;
            public byte indexFlag;
        }
        private ArrayList fields = new ArrayList(); // contains the columnTypes from the DBF
        public int totalNumberElements { get; private set; }
        private string pathToDBF;

        // open the DBF and init the FieldDescriptor & DBFHeader & totalNumberElements, place the file index at the first row 
        public DBFreader(string pathToDBF)
        {
            this.pathToDBF = pathToDBF;
            try
            {
                using (var dbfFile = new BinaryReader(File.OpenRead(pathToDBF)))
                {
                    // Read the header into a buffer    
                    byte[] buffer = dbfFile.ReadBytes(Marshal.SizeOf(typeof(DBFHeader)));

                    // Marshall the header into a DBFHeader structure
                    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                    header = (DBFHeader)Marshal.PtrToStructure(
                                        handle.AddrOfPinnedObject(), typeof(DBFHeader));
                    handle.Free();
                    totalNumberElements = header.numRecords;
                }   // using
            }
            catch (System.IO.IOException)
            {
                throw new Exception(string.Format("Nu se poate deschide fisierul \"{0}\"", pathToDBF));
            }

        }

        private void ReadHeader(BinaryReader dbfFile)
        {
            // Read the header into a buffer    
            byte[] buffer = dbfFile.ReadBytes(Marshal.SizeOf(typeof(DBFHeader)));

            // Marshall the header into a DBFHeader structure
            GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            header = (DBFHeader)Marshal.PtrToStructure(
                                handle.AddrOfPinnedObject(), typeof(DBFHeader));
            handle.Free();

            // Read in all the field descriptors. 
            // Per the spec, 13 (0D) marks the end of the field descriptors
            fields = new ArrayList();
            while ((13 != dbfFile.PeekChar()))
            {
                buffer = dbfFile.ReadBytes(Marshal.SizeOf(typeof(FieldDescriptor)));
                handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                fields.Add((FieldDescriptor)Marshal.PtrToStructure(
                            handle.AddrOfPinnedObject(), typeof(FieldDescriptor)));
                handle.Free();
            }

            dbfFile.ReadBytes(2);   // move 1 byte index in file so that current position is the first row & column data
            //currentIndex = 1;
            totalNumberElements = header.numRecords;
        }

        public IEnumerable<DataRow> readRows(DataTable relevantColumnsDataTable)
        {
            using (var dbfFile = new BinaryReader(File.OpenRead(pathToDBF)))
            {
                ReadHeader(dbfFile);
                for (int i = 0; i < totalNumberElements; i++)
                {
                    yield return getRowAndAdvanceIndex(relevantColumnsDataTable, dbfFile);
                }
            }
        }

        private DataRow getRowAndAdvanceIndex(DataTable relevantColumnsDataTable, BinaryReader dbfFile)
        {
            // First we'll read the entire record into a buffer and then read each 
            // field from the buffer. This helps account for any extra space at the 
            // end of each record and probably performs better.
            byte[] buffer = dbfFile.ReadBytes(header.recordLen);

            BinaryReader recReader = new BinaryReader(new MemoryStream(buffer));
            DataRow returnValue = relevantColumnsDataTable.Clone().NewRow();

            // mark the relevant columns from witch data will be extracted from the DBF
            for (int i = 0; i < fields.Count; i++)
            {
                FieldDescriptor match = (FieldDescriptor)fields[i];
                match.indexFlag = 0;
                fields[i] = match;
                foreach (DataColumn column in relevantColumnsDataTable.Columns)
                {
                    if (column.ColumnName == ((FieldDescriptor)fields[i]).fieldName)
                    {
                        match = (FieldDescriptor)fields[i];
                        match.indexFlag = 1;
                        fields[i] = match;
                        break;
                    }
                }
            }

            foreach (FieldDescriptor field in fields)
            {

                if (field.indexFlag == 1)
                {
                    switch (field.fieldType)
                    {
                        case 'D': // Date (YYYYMMDD)
                            var year = Encoding.ASCII.GetString(recReader.ReadBytes(4));
                            var month = Encoding.ASCII.GetString(recReader.ReadBytes(2));
                            var day = Encoding.ASCII.GetString(recReader.ReadBytes(2));
                            returnValue[field.fieldName] = System.DBNull.Value;
                            try
                            {
                                if ((Int32.Parse(year) > 1900))
                                {
                                    returnValue[field.fieldName] = new DateTime(Int32.Parse(year),
                                                               Int32.Parse(month), Int32.Parse(day));
                                }
                            }
                            catch
                            { }

                            break;
                        case 'C':
                            string readCharacters = Encoding.ASCII.GetString(recReader.ReadBytes(field.fieldLen));
                            returnValue[field.fieldName] = readCharacters.TrimEnd();
                            break;
                        case 'N':
                            byte[] readData = recReader.ReadBytes(field.fieldLen);
                            Array.Resize(ref readData, 8);  // resize the array dimension so that it has at least 4bytes required by ToInt32 (smaller will fail)
                            if (BitConverter.IsLittleEndian)
                                Array.Reverse(readData);
                            int readNumber = BitConverter.ToInt32(readData, 0);
                            returnValue[field.fieldName] = readNumber;
                            break;
                    }   // switch
                }   // if indexFlag == 1
                else
                {
                    recReader.ReadBytes(field.fieldLen);
                }
            }   // foreach

            recReader.Close();
            return returnValue;
        }   // getRowAndAdvanceIndex
    }
}
