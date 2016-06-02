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

        public int currentIndex { get; private set; }

        public int totalNumberElements { get; private set; }

        private BinaryReader dbfFile;

        // open the DBF and init the FieldDescriptor & DBFHeader & totalNumberElements, place the file index at the first row 
        public DBFreader(string pathToDBF)
        {
            // Read the header into a buffer    
            try
            {
                dbfFile = new BinaryReader(File.OpenRead(pathToDBF));
            }
            catch (System.IO.IOException)
            {
                Console.WriteLine("Fisierul \"{0}\" este deschis in alta aplicatie.", pathToDBF);
                Console.ReadKey();
                Environment.Exit(1);
            }
            
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
            currentIndex = 1;
            totalNumberElements = header.numRecords;
        }

        ~DBFreader ()
        {
            if (dbfFile!=null)
            {
                dbfFile.Close();
            }
        }


        public DataRow getRowAndAdvanceIndex(DataTable relevantColumnsDataTable)
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


        // old implementation
        //public List<String> getRowData(List<string> dataFromThisColumns)
        //{
        //    // Read the header into a buffer
        //    BinaryReader br = new BinaryReader(File.OpenRead("ceva"));
        //    byte[] buffer = br.ReadBytes(Marshal.SizeOf(typeof(DBFHeader)));

        //    // Marshall the header into a DBFHeader structure
        //    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        //    DBFHeader header = (DBFHeader)Marshal.PtrToStructure(
        //                        handle.AddrOfPinnedObject(), typeof(DBFHeader));
        //    handle.Free();

        //    // Read in all the field descriptors. 
        //    // Per the spec, 13 (0D) marks the end of the field descriptors
        //    ArrayList fields = new ArrayList();
        //    while ((13 != br.PeekChar()))
        //    {
        //        buffer = br.ReadBytes(Marshal.SizeOf(typeof(FieldDescriptor)));
        //        handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        //        fields.Add((FieldDescriptor)Marshal.PtrToStructure(
        //                    handle.AddrOfPinnedObject(), typeof(FieldDescriptor)));
        //        handle.Free();
        //    }

        //    DataTable dt = new DataTable();
        //    DataColumn col = null;
        //    foreach (FieldDescriptor field in fields)
        //    {
        //        switch (field.fieldType)
        //        {
        //            case 'N':
        //                col = new DataColumn(field.fieldName, typeof(Int32));
        //                break;
        //            case 'C':
        //                col = new DataColumn(field.fieldName, typeof(string));
        //                break;
        //            case 'D':
        //                col = new DataColumn(field.fieldName, typeof(DateTime));
        //                break;
        //            case 'L':
        //                col = new DataColumn(field.fieldName, typeof(bool));
        //                break;
        //        }
        //        dt.Columns.Add(col);
        //    }

        //    // Read in all the records
        //    for (int counter = 1; counter <= header.numRecords; counter++)
        //    {
        //        // First we'll read the entire record into a buffer and then read each 
        //        // field from the buffer. This helps account for any extra space at the 
        //        // end of each record and probably performs better.
        //        buffer = br.ReadBytes(header.recordLen);
        //        BinaryReader recReader = new BinaryReader(new MemoryStream(buffer));

        //        // Loop through each field in a record
        //        DataRow row = dt.NewRow();
        //        foreach (FieldDescriptor field in fields)
        //        {
        //            switch (field.fieldType)
        //            {
        //                case 'D': // Date (YYYYMMDD)
        //                    var year = Encoding.ASCII.GetString(recReader.ReadBytes(4));
        //                    var month = Encoding.ASCII.GetString(recReader.ReadBytes(2));
        //                    var day = Encoding.ASCII.GetString(recReader.ReadBytes(2));
        //                    row[field.fieldName] = System.DBNull.Value;
        //                    try
        //                    {
        //                        if ((Int32.Parse(year) > 1900))
        //                        {
        //                            row[field.fieldName] = new DateTime(Int32.Parse(year),
        //                                                       Int32.Parse(month), Int32.Parse(day));
        //                        }
        //                    }
        //                    catch
        //                    { }

        //                    break;
        //                case 'C':
        //                    string readCharacters = Encoding.ASCII.GetString(recReader.ReadBytes(field.fieldLen));
        //                    row[field.fieldName] = readCharacters;
        //                    break;
        //                case 'N':
        //                    byte[] readData = recReader.ReadBytes(field.fieldLen);
        //                    Array.Resize(ref readData, 8);  // resize the array dimension so that it has at least 4bytes required by ToInt32 (smaller will fail)
        //                    if (BitConverter.IsLittleEndian)
        //                        Array.Reverse(readData);
        //                    int readNumber = BitConverter.ToInt32(readData, 0);
        //                    row[field.fieldName] = readNumber;
        //                    break;
        //            }   // switch
        //            //Console.WriteLine("Rand {0}|   Coloana {1}|   VALOARE|{2}|", counter, field.fieldName, row[field.fieldName]);
        //        }   // FieldDescriptor
        //        recReader.Close();
        //        dt.Rows.Add(row);
        //    }   // counter
        //    return new List<String>();
        //}   // getRowData
    }
}
