using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _Excel = Microsoft.Office.Interop.Excel;

namespace IDcardsExcelAutomationGui
{
    internal class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        
        public Excel()
        {
            // constructor for creating a blank excel wb
        }

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet); // creating a workbook
            this.ws = wb.Worksheets[1]; // creating ws
        }
        public void CreateNewSheet()
        {
            Worksheet tempsheet = wb.Worksheets.Add(After: ws);
        }
        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return "";
        }
        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }
        public void SelectWorksheet(int SheetNumber)
        {
            try // incase that there is no sheet by sheet number
            {
                this.ws = wb.Worksheets[SheetNumber];   // this is simply assigning this current instance to current ws.
            }
            catch(Exception)
            {
                CreateNewSheet();
            }
            
        }
        public void DeleteWorksheet(int SheetNumber)  
        {
            wb.Worksheets[SheetNumber].Delete();
        }
        public void Close()
        {
            wb.Close(true);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }

        public int getLastRow()
        {            
            Range r = ws.UsedRange;
            int countRecords = r.Rows.Count;
            return countRecords;
        }

        public string[] getSN(int lastRow)
        {
            int startRow = 10;  // DroppedUnivFile starts from line 10
            int SN_clmn = 4;    // DroppedUnivFile column is at 4
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration
            string[] sn = new string[length];

            for (int i = 0; i < length; i++)
            {
                sn[i] = ws.Cells[startRow + i, SN_clmn].Value2;
            }

            return sn;
        }

        public string[] getEN(int lastRow)
        {
            int startRow = 10;  // DroppedUnivFile starts from line 10
            int SN_clmn = 3;    // DroppedUnivFile column is at 4
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration
            string[] en = new string[length];

            for (int i = 0; i < length; i++)
            {
                en[i] = ws.Cells[startRow + i, SN_clmn].Value2;
            }

            return en;
        }

        public string[] getSex(int lastRow)
        {
            int startRow = 10;  // DroppedUnivFile starts from line 10
            int SN_clmn = 9;    // DroppedUnivFile column is at 4
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration
            string[] sex = new string[length];

            for (int i = 0; i < length; i++)
            {
                sex[i] = ws.Cells[startRow + i, SN_clmn].Value2;
            }

            return sex;
        }
        // getting card-number
        public string[] getCN(int lastRow)
        {
            int startRow = 2;  // Dropped out of System starts from row 2
            int CN_clmn = 12;    // Card number column
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration
            string[] cn = new string[length];

            for (int i = 0; i < length; i++)
            {
                cn[i] = ws.Cells[startRow + i, CN_clmn].Value2;
            }
            return cn;
        }
        //getting Degree(차수)
        public string[] getDegree(int lastRow)
        {
            int startRow = 2;  // Dropped out of System starts from row 2
            int CN_clmn = 14;    // Card number column
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration
            string[] cn = new string[length];

            for (int i = 0; i < length; i++)
            {
                cn[i] = ws.Cells[startRow + i, CN_clmn].Value2;
            }
            return cn;
        }
        //getting CSN
        public string[] getCSN(int lastRow)
        {
            int startRow = 2;  // Dropped out of System starts from row 2
            int CN_clmn = 13;    // Card number column
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration
            string[] cn = new string[length];

            for (int i = 0; i < length; i++)
            {
                cn[i] = ws.Cells[startRow + i, CN_clmn].Value2;
            }
            return cn;
        }

        public string[] getRFID(int lastRow, string[] csn)
        {
            int startRow = 2;  // Dropped out of System starts from row 2            
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration
            string[] rfid = new string[length];            
            for (int i = 0; i < length; i++)
            {
                rfid[i] = csn[i].Substring(6,2) + csn[i].Substring(4,2) + csn[i].Substring(2, 2) + csn[i].Substring(0, 2);
                
            }
            return rfid;
        }

        public void goSaveToLibrary(string path)
        {
            wb.SaveAs(path);
        }

        public void writeCN_CSN_RFID(string[] cn, int lastRow, string[] csn, string[] rfid)
        {
            int startRow = 10;  // DroppedUnivFile starts from line 10            
            int length = lastRow - startRow + 1;  // starting point should not be taken into cosideration           
            int formRow = 9;
            int cnColumn = 10; int csnColumn = 11; int rfidColumn = 12;
            // format
            ws.Cells[formRow, cnColumn].Value2 = "카드번호";
            ws.Cells[formRow, csnColumn].Value2 = "CSN";
            ws.Cells[formRow, rfidColumn].Value2 = "RFID";

            for (int i = 0; i < length; i++)
            {
                ws.Cells[startRow + i, cnColumn].Value2 = cn[i];
                ws.Cells[startRow + i, csnColumn].Value2 = csn[i];
                ws.Cells[startRow + i, rfidColumn].Value2 = rfid[i];
            }
        }
    }
}
