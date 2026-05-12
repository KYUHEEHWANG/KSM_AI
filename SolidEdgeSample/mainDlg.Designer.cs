
namespace KSM_SolidEdge
{
    partial class mainDlg
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnExecute = new System.Windows.Forms.Button();
            this.grpSource = new System.Windows.Forms.GroupBox();
            this.btnSourceBrowse = new System.Windows.Forms.Button();
            this.txtSourcePath = new System.Windows.Forms.TextBox();
            this.lblSource = new System.Windows.Forms.Label();
            this.btnOutputBrowse = new System.Windows.Forms.Button();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            this.lblOutput = new System.Windows.Forms.Label();
            this.grpOutput = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cboImageType = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colResult = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnReset = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.grpSource.SuspendLayout();
            this.grpOutput.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExecute
            // 
            this.btnExecute.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnExecute.Location = new System.Drawing.Point(650, 590);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(120, 30);
            this.btnExecute.TabIndex = 0;
            this.btnExecute.Text = "실행";
            this.btnExecute.UseVisualStyleBackColor = false;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // grpSource
            // 
            this.grpSource.Controls.Add(this.btnSourceBrowse);
            this.grpSource.Controls.Add(this.txtSourcePath);
            this.grpSource.Controls.Add(this.lblSource);
            this.grpSource.Location = new System.Drawing.Point(10, 10);
            this.grpSource.Name = "grpSource";
            this.grpSource.Size = new System.Drawing.Size(760, 80);
            this.grpSource.TabIndex = 1;
            this.grpSource.TabStop = false;
            this.grpSource.Text = "3D 파일 위치";
            // 
            // btnSourceBrowse
            // 
            this.btnSourceBrowse.Location = new System.Drawing.Point(630, 26);
            this.btnSourceBrowse.Name = "btnSourceBrowse";
            this.btnSourceBrowse.Size = new System.Drawing.Size(100, 26);
            this.btnSourceBrowse.TabIndex = 2;
            this.btnSourceBrowse.Text = "찾아보기...";
            this.btnSourceBrowse.UseVisualStyleBackColor = true;
            this.btnSourceBrowse.Click += new System.EventHandler(this.btnSourceBrowse_Click);
            // 
            // txtSourcePath
            // 
            this.txtSourcePath.Location = new System.Drawing.Point(100, 28);
            this.txtSourcePath.Name = "txtSourcePath";
            this.txtSourcePath.Size = new System.Drawing.Size(520, 21);
            this.txtSourcePath.TabIndex = 1;
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(15, 30);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(57, 12);
            this.lblSource.TabIndex = 0;
            this.lblSource.Text = "소스 폴더";
            // 
            // btnOutputBrowse
            // 
            this.btnOutputBrowse.Location = new System.Drawing.Point(630, 26);
            this.btnOutputBrowse.Name = "btnOutputBrowse";
            this.btnOutputBrowse.Size = new System.Drawing.Size(100, 26);
            this.btnOutputBrowse.TabIndex = 2;
            this.btnOutputBrowse.Text = "찾아보기...";
            this.btnOutputBrowse.UseVisualStyleBackColor = true;
            this.btnOutputBrowse.Click += new System.EventHandler(this.btnOutputBrowse_Click);
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Location = new System.Drawing.Point(100, 28);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.Size = new System.Drawing.Size(520, 21);
            this.txtOutputPath.TabIndex = 1;
            // 
            // lblOutput
            // 
            this.lblOutput.AutoSize = true;
            this.lblOutput.Location = new System.Drawing.Point(15, 30);
            this.lblOutput.Name = "lblOutput";
            this.lblOutput.Size = new System.Drawing.Size(57, 12);
            this.lblOutput.TabIndex = 0;
            this.lblOutput.Text = "출력 폴더";
            // 
            // grpOutput
            // 
            this.grpOutput.Controls.Add(this.btnOutputBrowse);
            this.grpOutput.Controls.Add(this.txtOutputPath);
            this.grpOutput.Controls.Add(this.lblOutput);
            this.grpOutput.Location = new System.Drawing.Point(12, 96);
            this.grpOutput.Name = "grpOutput";
            this.grpOutput.Size = new System.Drawing.Size(760, 80);
            this.grpOutput.TabIndex = 2;
            this.grpOutput.TabStop = false;
            this.grpOutput.Text = "변환 파일 저장 위치";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cboImageType);
            this.groupBox1.Location = new System.Drawing.Point(12, 182);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(760, 55);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "이미지 유형";
            // 
            // cboImageType
            // 
            this.cboImageType.FormattingEnabled = true;
            this.cboImageType.Location = new System.Drawing.Point(100, 20);
            this.cboImageType.Name = "cboImageType";
            this.cboImageType.Size = new System.Drawing.Size(121, 20);
            this.cboImageType.TabIndex = 4;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dataGridView1);
            this.groupBox2.Location = new System.Drawing.Point(12, 243);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(758, 337);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "변환 결과 목록";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colStatus,
            this.colFileName,
            this.colResult});
            this.dataGridView1.Location = new System.Drawing.Point(17, 21);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(735, 304);
            this.dataGridView1.TabIndex = 0;
            // 
            // colNo
            // 
            this.colNo.HeaderText = "번호";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.Width = 80;
            // 
            // colStatus
            // 
            this.colStatus.HeaderText = "상태";
            this.colStatus.Name = "colStatus";
            this.colStatus.ReadOnly = true;
            this.colStatus.Width = 80;
            // 
            // colFileName
            // 
            this.colFileName.HeaderText = "파일명";
            this.colFileName.Name = "colFileName";
            this.colFileName.ReadOnly = true;
            this.colFileName.Width = 430;
            // 
            // colResult
            // 
            this.colResult.HeaderText = "결과";
            this.colResult.Name = "colResult";
            this.colResult.ReadOnly = true;
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(398, 590);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(120, 30);
            this.btnReset.TabIndex = 5;
            this.btnReset.Text = "초기화";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(27, 593);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(365, 23);
            this.progressBar1.TabIndex = 1;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(167, 619);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(57, 12);
            this.lblStatus.TabIndex = 6;
            this.lblStatus.Text = "대기 중...";
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnCancel.ForeColor = System.Drawing.SystemColors.Window;
            this.btnCancel.Location = new System.Drawing.Point(524, 590);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(120, 30);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "취소";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // mainDlg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 652);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.grpOutput);
            this.Controls.Add(this.grpSource);
            this.Controls.Add(this.btnExecute);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "mainDlg";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "파일 변환기";
            this.grpSource.ResumeLayout(false);
            this.grpSource.PerformLayout();
            this.grpOutput.ResumeLayout(false);
            this.grpOutput.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.GroupBox grpSource;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.Button btnSourceBrowse;
        private System.Windows.Forms.TextBox txtSourcePath;
        private System.Windows.Forms.Button btnOutputBrowse;
        private System.Windows.Forms.TextBox txtOutputPath;
        private System.Windows.Forms.Label lblOutput;
        private System.Windows.Forms.GroupBox grpOutput;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cboImageType;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colResult;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnCancel;
    }
}

