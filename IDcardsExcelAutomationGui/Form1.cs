using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IDcardsExcelAutomationGui
{
    public partial class Form_Univ_In_File : Form
    {
        string[] droppedFiles;  // PathAndFiles -> Dropped from Univ
        string droppedPath;
        string[] sn;            // student numbers
        string[] en;            // english names
        string selectedOrg =""; // organization combobox value       
        int ActualStudentNumber; // student counting
        string[] sex;
        string[] cn;
        string[] csn;
        string[] deg;
        string[] rfid;
        string unvi_sending_file_path_name;
        int UnivlastRow;
        //Excel toLibrary;

        string saveLocationToLibrary; //library file

        public Form_Univ_In_File()
        {
            InitializeComponent();
        }

        private void lsb_Univ_Dropped_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }
        // Universities files dropped--> 
        public void lsb_Univ_Dropped_DragDrop(object sender, DragEventArgs e)
        {
            // display a dropped file 
            droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
            lsb_Univ_Dropped.Items.Add(droppedFiles[0]);

            // path only
            string currentDirectory = Path.GetDirectoryName(droppedFiles[0]);
            droppedPath = Path.GetFullPath(currentDirectory);            

            // get the file name
            string filename = getFileName(droppedFiles[0]);

            // univ-sending-file-path & name
            unvi_sending_file_path_name = droppedFiles[0];            

            // diplay status
            ID_print_converting_status.Text = "파일이 정상적으로 인식됨.";
            ID_print_converting_status.BackColor = Color.ForestGreen;
            lsb_Univ_Dropped.BackColor = Color.ForestGreen;
            btn_org_write.BackColor = Color.Yellow;
            combo_org.BackColor = Color.Yellow;
            btn_org_write.Text = "조직 선택 ✔️클릭";
        }

        // organization selection from combobox 
        private void btn_org_write_Click(object sender, EventArgs e)
        {
            selectedOrg = getComboSelected();

            // displaying status
            ID_print_converting_status.Text = "✔️조직이 선택됨.";
            lsb_Univ_Dropped.BackColor = Color.White; // remove signal
            
            btn_ToPrintID.BackColor = Color.Yellow;
            btn_ToPrintID.Text = "프린트 파일 생성 ✔️클릭";

            combo_org.BackColor = Color.White;
            btn_org_write.BackColor = Color.White;    // remove signal

        }

        // selected combobox
        public string getComboSelected()
        {
            if (this.combo_org.SelectedItem != null)           // System.NullreferenceException handling
            {
                return this.combo_org.SelectedItem.ToString();
            } 
            
            else   // if null, then forcifully put a defaul value
            {
                this.combo_org.SelectedItem = "GMUK";
                return this.combo_org.SelectedItem.ToString();
            }            
        }
        private void Form_Univ_In_File_Load(object sender, EventArgs e)
        {
            this.combo_org.SelectedText = "GMUK";
        }

        public void btn_ToPrintID_Click(object sender, EventArgs e)
        {
            // displaying status
            btn_ToPrintID.BackColor = Color.White;
            ID_print_converting_status.Text = "파일...생성중...\n\n...기다려...쫌만...더...";
            ID_print_converting_status.BackColor = Color.GreenYellow;

            // ✔️open University-sending file            
            Excel droppedFile = new Excel(droppedFiles[0], 1);
            UnivlastRow = droppedFile.getLastRow();

            sn = droppedFile.getSN(UnivlastRow);            // Read contents: student Numbers
            en = droppedFile.getEN(UnivlastRow);
            sex = droppedFile.getSex(UnivlastRow);            
            
            // create and save a new file for printing ID card file            
            ToPrintIDForm(UnivlastRow);

            // open sending-to-library file
            //copyToLibrary();    

            lsb_Univ_Dropped.Items.Clear();  // remove from the listbox

            // ✔️close workbook : University sending-file            
            droppedFile.Close();

            // status            
            ID_print_converting_status.Text = "✔️ID 카드출력용 파일 생성완료!!!";
            ID_print_converting_status.BackColor = Color.Goldenrod;

            lsb_System_Dropped.BackColor = Color.Yellow;
            MessageBox.Show("두번째 노란 박스에 파일을\n 끌어다 놓으세요!!!");
        }
        public void ToPrintIDForm(int UnivlastRow)
        {
            // boilerplateRow & Actual Number of Students
            int boilerplateRow = 9;
            ActualStudentNumber = UnivlastRow - boilerplateRow;
            
            // creating-file for ID printing
            ToPrintID tpi = new ToPrintID();

            // saving id-print-file
            string idprint = "▣ID카드 프린트용 파일_";
            string saveLocation = droppedPath + @"\" + idprint + ActualStudentNumber.ToString();
            
            tpi.CreateTheForm(saveLocation, sn, en, selectedOrg);  // sn is string[] of students

            // close workbook: id-print-file
            tpi.Close();
        }

        public string getFileName(string path)
        {
            return Path.GetFileNameWithoutExtension(path);
        }

        // ID Print System Files dropped-->
        private void lsb_System_Dropped_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }
        private void lsb_System_Dropped_DragDrop(object sender, DragEventArgs e)
        {
            btn_ToPrintID.BackColor = Color.White; // remove signal

            // display a dropped file 
            droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
            lsb_System_Dropped.Items.Add(droppedFiles[0]);

            // path only
            string currentdirectory = Path.GetDirectoryName(droppedFiles[0]);
            droppedPath = Path.GetFullPath(currentdirectory);

            // get the file name
            string filename = getFileName(droppedFiles[0]);

            // displaying
            ID_print_converting_status.Text = "✔️파일이 정상적으로 인식됨.";
            ID_print_converting_status.BackColor = Color.ForestGreen;

            btn_ToStaff_Info.BackColor = Color.Yellow;  // to signal
            btn_ToStaff_Info.Text = "사원정보 파일생성 ✔️클릭";
        }

        

        private void btn_ToStaff_Info_Click(object sender, EventArgs e)
        {
            // displaying
            ID_print_converting_status.Text = "사원정보 파일을 생성중...\n\n...인내심 요청중...";
            ID_print_converting_status.BackColor = Color.GreenYellow;

            // open system-creating-file           
            Excel droppedFile = new Excel(droppedFiles[0], 1);
            int lastRow = droppedFile.getLastRow();

            cn = droppedFile.getCN(lastRow);
            csn = droppedFile.getCSN(lastRow);
            rfid = droppedFile.getRFID(lastRow, csn);
            deg = droppedFile.getDegree(lastRow);

            // create and save a new file for printing ID card file            
            CreateStaffForm(lastRow);

            // ID card printing form : status            
            ID_print_converting_status.Text = "사원정보 파일 생성 완료!!!";
            ID_print_converting_status.BackColor = Color.Goldenrod; 
            btn_ToCard_Info.BackColor= Color.Yellow; // to signal 
            btn_ToCard_Info.Text = "카드정보 생성 ✔️클릭";
            btn_ToStaff_Info.BackColor = Color.White; // remove signal

            droppedFile.Close();
        }
        public void CreateStaffForm(int lastRow)
        {
            // boilerplateRow & Actual Number of Students
            int boilerplateRow = 1;
            ActualStudentNumber = lastRow - boilerplateRow;

            // creating(open) staff-info-form
            ToPrintID tpi = new ToPrintID();

            // path of staff-info-form
            string idprint = "▣▣사원정보 입력용 파일_";
            string saveLocation = droppedPath + @"\" + idprint + ActualStudentNumber.ToString();

            // copy from system-file to staff-info-file
            tpi.CreateStaffFormSave(saveLocation, sn, en, selectedOrg, sex);  // sn is string[] of students' name 
            // close staff-info-file
            tpi.Close();
            lsb_Univ_Dropped.Items.Clear();  // after converting file is removed and related arrarys are cleared.
            
        }

        private void btn_ToCard_Info_Click(object sender, EventArgs e)
        {
            // status
            ID_print_converting_status.Text = "카드정보 생성중!!! \n\n5초만 기다려줭..";
            ID_print_converting_status.BackColor = Color.GreenYellow;

            // open system-dropped-file           
            Excel droppedFile = new Excel(droppedFiles[0], 1);
            int lastRow = droppedFile.getLastRow();

            // open univ-sending-file
            copyToLibrary(unvi_sending_file_path_name, lastRow,cn, csn, rfid);           
                        
            // create and save a new file for printing ID card file            
            CreateCardForm(lastRow);

            // status-displaying           
            ID_print_converting_status.Text = "✔️카드정보 생성완료!!!";
            ID_print_converting_status.BackColor = Color.Goldenrod;

            // status-displaying
            btn_ToStaff_Info.BackColor = Color.White;
            btn_ToCard_Info.BackColor = Color.White;

            // close workbook : system-dropped-file
            droppedFile.Close();
            
            // remove signal
            lsb_System_Dropped.BackColor = Color.White;

            MessageBox.Show("축하합니다. 모든 작업이 끝났어요!");

            // release arrays
            Array.Clear(droppedFiles, 0, droppedFiles.Length);
        }
        public void CreateCardForm(int lastRow)
        {
            // boilerplateRow & Actual Number of Students
            int boilerplateRow = 1;
            ActualStudentNumber = lastRow - boilerplateRow;

            // open card-info-file
            ToPrintID tpi = new ToPrintID();

            // path card-info-file
            string idprint = "▣▣▣카드정보 입력용 파일_";
            string ext = ".xlsx";
            string saveLocation = droppedPath + @"\" + idprint + ActualStudentNumber + ext;

            tpi.CreateCardFormSave(saveLocation, sn, en, selectedOrg, cn);  // sn is string[] of students 

            // after converting file is removed and related arrarys are cleared.
            tpi.Close();
            lsb_System_Dropped.Items.Clear();              
        }
        // create Tolibrary file
        public void copyToLibrary(string unvi_sending_file_path_name, int lastRow, string[] cn, string[] csn, string[] rfid)
        {
            // get the dropped from university
            Excel toLibrary = new Excel(unvi_sending_file_path_name, 1);  // open a file with Path and Sheet number 1

            // saving-location
            string idprint = "✔️도서관발송용파일" + "_" + selectedOrg + "_";
            string ext = ".xlsx";
            saveLocationToLibrary = droppedPath + @"\" + idprint + ActualStudentNumber + ext;

            // write card number
            toLibrary.writeCN_CSN_RFID(cn, UnivlastRow, csn, rfid);
            toLibrary.SaveAs(saveLocationToLibrary);

            // close workbook : file to library
            toLibrary.Close();

        }
    }
}
