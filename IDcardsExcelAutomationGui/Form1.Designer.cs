namespace IDcardsExcelAutomationGui
{
    partial class Form_Univ_In_File
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.btn_ToPrintID = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_ToStaff_Info = new System.Windows.Forms.Button();
            this.btn_ToCard_Info = new System.Windows.Forms.Button();
            this.lsb_Univ_Dropped = new System.Windows.Forms.ListBox();
            this.ID_print_converting_status = new System.Windows.Forms.Label();
            this.btn_org_write = new System.Windows.Forms.Button();
            this.lsb_System_Dropped = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.combo_org = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Gulim", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.ForeColor = System.Drawing.Color.SaddleBrown;
            this.label1.Location = new System.Drawing.Point(36, 162);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "받은파일(대학→)";
            // 
            // btn_ToPrintID
            // 
            this.btn_ToPrintID.Font = new System.Drawing.Font("Gulim", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_ToPrintID.Location = new System.Drawing.Point(293, 209);
            this.btn_ToPrintID.Name = "btn_ToPrintID";
            this.btn_ToPrintID.Size = new System.Drawing.Size(283, 36);
            this.btn_ToPrintID.TabIndex = 2;
            this.btn_ToPrintID.Text = "ID카드 프린트 파일생성";
            this.btn_ToPrintID.UseVisualStyleBackColor = true;
            this.btn_ToPrintID.Click += new System.EventHandler(this.btn_ToPrintID_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Gulim", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label2.ForeColor = System.Drawing.Color.SaddleBrown;
            this.label2.Location = new System.Drawing.Point(36, 503);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(269, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "프린트 ID카드 완료파일(시스템→)";
            // 
            // btn_ToStaff_Info
            // 
            this.btn_ToStaff_Info.Font = new System.Drawing.Font("Gulim", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_ToStaff_Info.Location = new System.Drawing.Point(323, 612);
            this.btn_ToStaff_Info.Name = "btn_ToStaff_Info";
            this.btn_ToStaff_Info.Size = new System.Drawing.Size(283, 36);
            this.btn_ToStaff_Info.TabIndex = 2;
            this.btn_ToStaff_Info.Text = "사원정보 파일생성";
            this.btn_ToStaff_Info.UseVisualStyleBackColor = true;
            this.btn_ToStaff_Info.Click += new System.EventHandler(this.btn_ToStaff_Info_Click);
            // 
            // btn_ToCard_Info
            // 
            this.btn_ToCard_Info.Font = new System.Drawing.Font("Gulim", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_ToCard_Info.Location = new System.Drawing.Point(323, 671);
            this.btn_ToCard_Info.Name = "btn_ToCard_Info";
            this.btn_ToCard_Info.Size = new System.Drawing.Size(283, 36);
            this.btn_ToCard_Info.TabIndex = 2;
            this.btn_ToCard_Info.Text = "카드정보 파일생성";
            this.btn_ToCard_Info.UseVisualStyleBackColor = true;
            this.btn_ToCard_Info.Click += new System.EventHandler(this.btn_ToCard_Info_Click);
            // 
            // lsb_Univ_Dropped
            // 
            this.lsb_Univ_Dropped.AllowDrop = true;
            this.lsb_Univ_Dropped.Font = new System.Drawing.Font("Gulim", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.lsb_Univ_Dropped.FormattingEnabled = true;
            this.lsb_Univ_Dropped.ItemHeight = 16;
            this.lsb_Univ_Dropped.Location = new System.Drawing.Point(39, 181);
            this.lsb_Univ_Dropped.Name = "lsb_Univ_Dropped";
            this.lsb_Univ_Dropped.Size = new System.Drawing.Size(567, 84);
            this.lsb_Univ_Dropped.TabIndex = 4;
            this.lsb_Univ_Dropped.DragDrop += new System.Windows.Forms.DragEventHandler(this.lsb_Univ_Dropped_DragDrop);
            this.lsb_Univ_Dropped.DragEnter += new System.Windows.Forms.DragEventHandler(this.lsb_Univ_Dropped_DragEnter);
            // 
            // ID_print_converting_status
            // 
            this.ID_print_converting_status.AutoSize = true;
            this.ID_print_converting_status.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ID_print_converting_status.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ID_print_converting_status.Font = new System.Drawing.Font("Gulim", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.ID_print_converting_status.Location = new System.Drawing.Point(33, 18);
            this.ID_print_converting_status.Name = "ID_print_converting_status";
            this.ID_print_converting_status.Size = new System.Drawing.Size(505, 24);
            this.ID_print_converting_status.TabIndex = 5;
            this.ID_print_converting_status.Text = "아래 첫번째 박스에 파일을 끌어다 놓으세요";
            // 
            // btn_org_write
            // 
            this.btn_org_write.Font = new System.Drawing.Font("Gulim", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_org_write.Location = new System.Drawing.Point(382, 157);
            this.btn_org_write.Name = "btn_org_write";
            this.btn_org_write.Size = new System.Drawing.Size(194, 36);
            this.btn_org_write.TabIndex = 7;
            this.btn_org_write.Text = "조직입력 버튼";
            this.btn_org_write.UseVisualStyleBackColor = true;
            this.btn_org_write.Click += new System.EventHandler(this.btn_org_write_Click);
            // 
            // lsb_System_Dropped
            // 
            this.lsb_System_Dropped.AllowDrop = true;
            this.lsb_System_Dropped.Font = new System.Drawing.Font("Gulim", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.lsb_System_Dropped.FormattingEnabled = true;
            this.lsb_System_Dropped.ItemHeight = 16;
            this.lsb_System_Dropped.Location = new System.Drawing.Point(39, 522);
            this.lsb_System_Dropped.Name = "lsb_System_Dropped";
            this.lsb_System_Dropped.Size = new System.Drawing.Size(567, 84);
            this.lsb_System_Dropped.TabIndex = 8;
            this.lsb_System_Dropped.DragDrop += new System.Windows.Forms.DragEventHandler(this.lsb_System_Dropped_DragDrop);
            this.lsb_System_Dropped.DragEnter += new System.Windows.Forms.DragEventHandler(this.lsb_System_Dropped_DragEnter);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.LightGray;
            this.groupBox1.Controls.Add(this.combo_org);
            this.groupBox1.Controls.Add(this.btn_org_write);
            this.groupBox1.Controls.Add(this.btn_ToPrintID);
            this.groupBox1.Location = new System.Drawing.Point(30, 124);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(587, 281);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            // 
            // combo_org
            // 
            this.combo_org.DropDownHeight = 120;
            this.combo_org.Font = new System.Drawing.Font("Gulim", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.combo_org.FormattingEnabled = true;
            this.combo_org.IntegralHeight = false;
            this.combo_org.ItemHeight = 21;
            this.combo_org.Items.AddRange(new object[] {
            "GMUK",
            "UTAH",
            "SUNY",
            "GHENT"});
            this.combo_org.Location = new System.Drawing.Point(223, 157);
            this.combo_org.Name = "combo_org";
            this.combo_org.Size = new System.Drawing.Size(153, 29);
            this.combo_org.TabIndex = 6;
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.LightGray;
            this.groupBox2.Location = new System.Drawing.Point(30, 461);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(587, 269);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            // 
            // Form_Univ_In_File
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(638, 773);
            this.Controls.Add(this.lsb_System_Dropped);
            this.Controls.Add(this.ID_print_converting_status);
            this.Controls.Add(this.lsb_Univ_Dropped);
            this.Controls.Add(this.btn_ToCard_Info);
            this.Controls.Add(this.btn_ToStaff_Info);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Font = new System.Drawing.Font("Gulim", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Name = "Form_Univ_In_File";
            this.Text = "ID카드 발급 파일생성기(제작: Tommy)";
            this.Load += new System.EventHandler(this.Form_Univ_In_File_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_ToPrintID;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_ToStaff_Info;
        private System.Windows.Forms.Button btn_ToCard_Info;
        private System.Windows.Forms.ListBox lsb_Univ_Dropped;
        private System.Windows.Forms.Label ID_print_converting_status;
        private System.Windows.Forms.Button btn_org_write;
        private System.Windows.Forms.ListBox lsb_System_Dropped;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox combo_org;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}

