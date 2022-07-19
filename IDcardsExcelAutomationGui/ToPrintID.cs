using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _Excel = Microsoft.Office.Interop.Excel;

namespace IDcardsExcelAutomationGui
{
    internal class ToPrintID
    {
        _Application toprintid = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        
        public void CreateTheForm(string path, string[] studentNumber, string[] englishName, string selectedOrg)
        {
            this.wb = toprintid.Workbooks.Add(XlWBATemplate.xlWBATWorksheet); // creating a workbook
            ws = wb.Worksheets[1]; // creating a ws

            // building a boilerplate
            ws.Cells[1, 1] = "번호"; ws.Cells[1, 2] = "카드번호"; ws.Cells[1, 3] = "한글명"; ws.Cells[1, 4] = "영문명"; ws.Cells[1, 5] = "소속명"; ws.Cells[1, 6] = "학(사)번"; ws.Cells[1, 7] = "직위명"; ws.Cells[1, 8] = "재직상태"; ws.Cells[1, 9] = "신분명"; ws.Cells[1, 10] = "이메일"; ws.Cells[1, 11] = "비고";

            // writing Student numbers
            int startingRow = 2; int studentColumn = 1; int studentColumn2 = 6;
            int ugColumn = 7; int jobstatusColumn = 8; int orgColumn = 5;
            int startingColumn_en = 3; int startingColumn2_en = 4;            

            for (int i = 0; i < studentNumber.Length; i++)
            {
                ws.Cells[startingRow + i, studentColumn] = studentNumber[i];
                ws.Cells[startingRow + i, studentColumn2] = studentNumber[i];
                ws.Cells[startingRow + i, ugColumn] = "UG";    // writing job-status 
                ws.Cells[startingRow + i, jobstatusColumn] = "재직";    // writing job-status 
                ws.Cells[startingRow + i, orgColumn] = selectedOrg;    // writing organization
                ws.Cells[startingRow + i, startingColumn_en] = englishName[i]; // writing English Names
                ws.Cells[startingRow + i, startingColumn2_en] = englishName[i]; // writing English Names
            }

            wb.SaveAs(path);            
        }
        public void CreateStaffFormSave(string path, string[] studentNumber, string[] englishName, string selectedOrg, string[] sexs)
        {
            // fit organizations into the form
            switch(selectedOrg)
            {
                case "GMUK":
                    selectedOrg = "조지메이슨대학교";
                    break;
                case "SUNY":
                    selectedOrg = "한국뉴욕주립대학교";
                    break;
                case "UTAH":
                    selectedOrg = "유타대학교";
                    break;
                case "GHENT":
                    selectedOrg = "겐트대학교";
                    break;
            }

            this.wb = toprintid.Workbooks.Add(XlWBATemplate.xlWBATWorksheet); // creating a workbook
            ws = wb.Worksheets[1]; // creating a ws

            // building a boilerplate
            ws.Cells[1, 1] = "이름"; ws.Cells[1, 2] = "사원번호"; ws.Cells[1, 3] = "조직"; ws.Cells[1, 4] = "직급"; ws.Cells[1, 5] = "사원구분"; ws.Cells[1, 6] = "연락처"; ws.Cells[1, 7] = "근무상태"; ws.Cells[1, 8] = "휴대전화"; ws.Cells[1, 9] = "학교명"; ws.Cells[1, 10] = "학교신분명"; ws.Cells[1, 11] = "기타"; ws.Cells[1, 12] = "성별"; ws.Cells[1, 13] = "입학년도"; ws.Cells[1, 14] = "이메일"; ws.Cells[1, 15] = "사진(파일명)";

            // writing Student numbers
            int startingRow = 2; int nameColumn = 1; int SNColumn = 2;
            int orgColumn = 3; int rankColumn = 4; int workTypeColumn = 5;
            int workStatusColumn = 7; int phoneColumn = 8;
            int schoolNameColumn = 9; int schoolName2Column = 10; int etcColumn = 11; int sex = 12;
            int admitYearColumn = 13; 
            string eum = "이음_";
            
            // extracting year, month, day
            string todayToForm = getYearMonthDay() ;
            
            for (int i = 0; i < studentNumber.Length; i++)
            {
                string sexFinal = sexCheck(sexs[i]);
                ws.Cells[startingRow + i, nameColumn] = eum + englishName[i];
                ws.Cells[startingRow + i, SNColumn] = studentNumber[i];
                ws.Cells[startingRow + i, orgColumn] = selectedOrg;    
                ws.Cells[startingRow + i, rankColumn] = "학부생";    
                ws.Cells[startingRow + i, workTypeColumn] = "정규직"; 
                ws.Cells[startingRow + i, workStatusColumn] = "재직";
                ws.Cells[startingRow + i, schoolNameColumn] = selectedOrg;
                ws.Cells[startingRow + i, schoolName2Column] = selectedOrg;
                ws.Cells[startingRow + i, etcColumn] = todayToForm + sexFinal;
                ws.Cells[startingRow + i, sex] = sexFinal;
                ws.Cells[startingRow + i, admitYearColumn] = todayToForm;
            }            

            wb.SaveAs(path);            
        }

        public string sexCheck(string s)
        {
            string m = "남"; string f = "여";

            string sexCheck = "what the hell";
            if (s == "M") sexCheck = m;
            if (s == "남") sexCheck = m;
            if (s == "F") sexCheck = f;
            if (s == "여") sexCheck = f;

            return sexCheck;
        }
        public string getYearMonthDay()
        {           
            // Date
            var tdy = DateTime.Today;
            string td = tdy.ToShortDateString();

            // extracting year month day
            string[] todayToForms = td.Split('/');
            string todayToForm = todayToForms[2] + todayToForms[0] + todayToForms[1];

            return todayToForm;

        }

        public void CreateCardFormSave(string path, string[] sn, string[] englishName, string selectedOrg, string[] cn)
        {
            // fit organizations into the form
            switch (selectedOrg)
            {
                case "GMUK":
                    selectedOrg = "조지메이슨대학교";
                    break;
                case "SUNY":
                    selectedOrg = "한국뉴욕주립대학교";
                    break;
                case "UTAH":
                    selectedOrg = "유타대학교";
                    break;
                case "GHENT":
                    selectedOrg = "겐트대학교";
                    break;
            }

            this.wb = toprintid.Workbooks.Add(XlWBATemplate.xlWBATWorksheet); // creating a workbook
            ws = wb.Worksheets[1]; // creating a ws

            // building a boilerplate
            ws.Cells[1, 1] = "카드 번호"; ws.Cells[1, 2] = "상태"; ws.Cells[1, 3] = "카드구분"; ws.Cells[1, 4] = "카드종류"; ws.Cells[1, 5] = "재발급횟수"; ws.Cells[1, 6] = "사원 번호"; ws.Cells[1, 7] = "카드 유효기간"; ws.Cells[1, 8] = "사용자 정의 1"; ws.Cells[1, 9] = "사용자 정의 2"; ws.Cells[1, 10] = "사용자 정의 3"; ws.Cells[1, 11] = "사용자 정의 4"; ws.Cells[1, 12] = "사용자 정의 5"; ws.Cells[1, 13] = "마스터 권한"; ws.Cells[1, 14] = "출입 모드"; ws.Cells[1, 15] = "사원이름"; ws.Cells[1, 16] = "조직";

            // writing Student numbers
            int startingRow = 2; int cnColumn = 1; int statusColumn = 2;
            int typeColumn = 3; int basicColumn = 4; int reissueColumn = 5;
            int snColumn = 6; int validColumn = 7; int previledgeColumn = 13; 
            int inAndOutColumn = 14; int snameColumn = 15; int orgColumn = 16;
            
            string eum = "이음_";

            // extracting year, month, day
            string todayToForm = getYearMonthDay();
           
            for (int i = 0; i < sn.Length; i++)
            {
                //string sexFinal = sexCheck(sexs[i]);
                ws.Cells[startingRow + i, cnColumn] = cn[i];          // 카드 번호
                ws.Cells[startingRow + i, statusColumn] = "정상";     //  상태 
                ws.Cells[startingRow + i, typeColumn] = "일반";       // 카드 구분
                ws.Cells[startingRow + i, basicColumn] = "기본 카드";  // 카드 종류
                ws.Cells[startingRow + i, reissueColumn] = "0";      // 재발급 횟수
                ws.Cells[startingRow + i, snColumn] = sn[i];        // 사원번호  
                ws.Cells[startingRow + i, validColumn] = "2099-12-31";  // 카드유효기간
                ws.Cells[startingRow + i, previledgeColumn] = "일반";  // 마스터권한
                ws.Cells[startingRow + i, snameColumn] = eum + englishName[i]; // 이름               
                ws.Cells[startingRow + i, orgColumn] = selectedOrg;
            }

            wb.SaveAs(path);            
        }
        public void Close()
        {
            wb.Close(true);
            toprintid.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(toprintid);
        }
    }
}
