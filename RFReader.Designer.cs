namespace OA_PS_RFID
{
    partial class RFReader
    {
        /******************************************************************************/
        /// <summary>
        /// Required designer variable.
        private Pgrdrv2.CReaderCtl rdr = new Pgrdrv2.CReaderCtl();
        private bool readeravailable;

         private System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();            
        /******************************************************************************/



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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.Recption_daily_attend = new System.Windows.Forms.Label();
            this.RFCmbPorttype = new System.Windows.Forms.ComboBox();
            this.RFTxtPort = new System.Windows.Forms.TextBox();
            this.RFTxtReaderid = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.RFBtnSetdata = new System.Windows.Forms.Button();
            this.RFBtnReaderresponse = new System.Windows.Forms.Button();
            this.RFBtnConnect = new System.Windows.Forms.Button();
            this.RfListbox1 = new System.Windows.Forms.ListBox();
            this.RfBtnClear1 = new System.Windows.Forms.Button();
            this.RFBtnDisconnect = new System.Windows.Forms.Button();
            this.LblShow = new System.Windows.Forms.Label();
            this.RFReaderGroupBox = new System.Windows.Forms.GroupBox();
            this.RFReaderBtnRecep_window = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.RFReaderGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(682, 24);
            this.menuStrip1.TabIndex = 43;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem,
            this.exitToolStripMenuItem1});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(96, 22);
            this.exitToolStripMenuItem.Text = "Main";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem1
            // 
            this.exitToolStripMenuItem1.Name = "exitToolStripMenuItem1";
            this.exitToolStripMenuItem1.Size = new System.Drawing.Size(96, 22);
            this.exitToolStripMenuItem1.Text = "Exit";
            this.exitToolStripMenuItem1.Click += new System.EventHandler(this.exitToolStripMenuItem1_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // Recption_daily_attend
            // 
            this.Recption_daily_attend.BackColor = System.Drawing.Color.WhiteSmoke;
            this.Recption_daily_attend.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Recption_daily_attend.Location = new System.Drawing.Point(-4, 24);
            this.Recption_daily_attend.Name = "Recption_daily_attend";
            this.Recption_daily_attend.Size = new System.Drawing.Size(686, 24);
            this.Recption_daily_attend.TabIndex = 62;
            this.Recption_daily_attend.Text = "                                      INITIALIZING RF READER";
            // 
            // RFCmbPorttype
            // 
            this.RFCmbPorttype.BackColor = System.Drawing.SystemColors.Window;
            this.RFCmbPorttype.Cursor = System.Windows.Forms.Cursors.Default;
            this.RFCmbPorttype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.RFCmbPorttype.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RFCmbPorttype.ForeColor = System.Drawing.SystemColors.WindowText;
            this.RFCmbPorttype.Items.AddRange(new object[] {
            "TCP/IP",
            "Serial Port",
            "Parallel"});
            this.RFCmbPorttype.Location = new System.Drawing.Point(79, 66);
            this.RFCmbPorttype.Name = "RFCmbPorttype";
            this.RFCmbPorttype.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.RFCmbPorttype.Size = new System.Drawing.Size(102, 22);
            this.RFCmbPorttype.TabIndex = 1;
            this.RFCmbPorttype.Leave += new System.EventHandler(this.RFCmbPorttype_Leave);
            // 
            // RFTxtPort
            // 
            this.RFTxtPort.AcceptsReturn = true;
            this.RFTxtPort.BackColor = System.Drawing.SystemColors.Window;
            this.RFTxtPort.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.RFTxtPort.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RFTxtPort.ForeColor = System.Drawing.SystemColors.WindowText;
            this.RFTxtPort.Location = new System.Drawing.Point(79, 105);
            this.RFTxtPort.MaxLength = 4;
            this.RFTxtPort.Name = "RFTxtPort";
            this.RFTxtPort.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.RFTxtPort.Size = new System.Drawing.Size(53, 20);
            this.RFTxtPort.TabIndex = 3;
            this.RFTxtPort.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.RFTxtPort_KeyPress);
            // 
            // RFTxtReaderid
            // 
            this.RFTxtReaderid.BackColor = System.Drawing.SystemColors.Window;
            this.RFTxtReaderid.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.RFTxtReaderid.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RFTxtReaderid.ForeColor = System.Drawing.SystemColors.WindowText;
            this.RFTxtReaderid.Location = new System.Drawing.Point(79, 151);
            this.RFTxtReaderid.MaxLength = 2;
            this.RFTxtReaderid.Name = "RFTxtReaderid";
            this.RFTxtReaderid.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.RFTxtReaderid.Size = new System.Drawing.Size(41, 20);
            this.RFTxtReaderid.TabIndex = 5;
            this.RFTxtReaderid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.RFTxtReaderid_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Port Type";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 109);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Port Number";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 155);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Reader ID";
            // 
            // RFBtnSetdata
            // 
            this.RFBtnSetdata.Enabled = false;
            this.RFBtnSetdata.Location = new System.Drawing.Point(238, 151);
            this.RFBtnSetdata.Name = "RFBtnSetdata";
            this.RFBtnSetdata.Size = new System.Drawing.Size(65, 23);
            this.RFBtnSetdata.TabIndex = 7;
            this.RFBtnSetdata.Text = "Set Date";
            this.RFBtnSetdata.UseVisualStyleBackColor = true;
            this.RFBtnSetdata.Click += new System.EventHandler(this.RFBtnSetdata_Click);
            // 
            // RFBtnReaderresponse
            // 
            this.RFBtnReaderresponse.Enabled = false;
            this.RFBtnReaderresponse.Location = new System.Drawing.Point(157, 192);
            this.RFBtnReaderresponse.Name = "RFBtnReaderresponse";
            this.RFBtnReaderresponse.Size = new System.Drawing.Size(102, 26);
            this.RFBtnReaderresponse.TabIndex = 8;
            this.RFBtnReaderresponse.Text = "Reader Response";
            this.RFBtnReaderresponse.UseVisualStyleBackColor = true;
            this.RFBtnReaderresponse.Click += new System.EventHandler(this.RFBtnReaderresponse_Click);
            // 
            // RFBtnConnect
            // 
            this.RFBtnConnect.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.RFBtnConnect.Location = new System.Drawing.Point(157, 150);
            this.RFBtnConnect.Name = "RFBtnConnect";
            this.RFBtnConnect.Size = new System.Drawing.Size(65, 23);
            this.RFBtnConnect.TabIndex = 6;
            this.RFBtnConnect.Text = "Connect";
            this.RFBtnConnect.UseVisualStyleBackColor = true;
            this.RFBtnConnect.Click += new System.EventHandler(this.RFBtnConnect_Click);
            // 
            // RfListbox1
            // 
            this.RfListbox1.FormattingEnabled = true;
            this.RfListbox1.Items.AddRange(new object[] {
            "---------------OUT PUT---------------"});
            this.RfListbox1.Location = new System.Drawing.Point(366, 66);
            this.RfListbox1.Name = "RfListbox1";
            this.RfListbox1.Size = new System.Drawing.Size(146, 108);
            this.RfListbox1.TabIndex = 11;
            // 
            // RfBtnClear1
            // 
            this.RfBtnClear1.Location = new System.Drawing.Point(366, 180);
            this.RfBtnClear1.Name = "RfBtnClear1";
            this.RfBtnClear1.Size = new System.Drawing.Size(146, 26);
            this.RfBtnClear1.TabIndex = 10;
            this.RfBtnClear1.Text = "Clear";
            this.RfBtnClear1.UseVisualStyleBackColor = true;
            this.RfBtnClear1.Click += new System.EventHandler(this.RfBtnClear_Click);
            // 
            // RFBtnDisconnect
            // 
            this.RFBtnDisconnect.Enabled = false;
            this.RFBtnDisconnect.Location = new System.Drawing.Point(441, 242);
            this.RFBtnDisconnect.Name = "RFBtnDisconnect";
            this.RFBtnDisconnect.Size = new System.Drawing.Size(71, 23);
            this.RFBtnDisconnect.TabIndex = 40;
            this.RFBtnDisconnect.Text = "Disconnect";
            this.RFBtnDisconnect.UseVisualStyleBackColor = true;
            this.RFBtnDisconnect.Click += new System.EventHandler(this.RFBtnDisconnect_Click);
            // 
            // LblShow
            // 
            this.LblShow.AutoSize = true;
            this.LblShow.Location = new System.Drawing.Point(363, 50);
            this.LblShow.Name = "LblShow";
            this.LblShow.Size = new System.Drawing.Size(43, 13);
            this.LblShow.TabIndex = 12;
            this.LblShow.Text = "Out Put";
            // 
            // RFReaderGroupBox
            // 
            this.RFReaderGroupBox.Controls.Add(this.RFReaderBtnRecep_window);
            this.RFReaderGroupBox.Controls.Add(this.LblShow);
            this.RFReaderGroupBox.Controls.Add(this.RFBtnDisconnect);
            this.RFReaderGroupBox.Controls.Add(this.RfBtnClear1);
            this.RFReaderGroupBox.Controls.Add(this.RfListbox1);
            this.RFReaderGroupBox.Controls.Add(this.RFBtnConnect);
            this.RFReaderGroupBox.Controls.Add(this.RFBtnReaderresponse);
            this.RFReaderGroupBox.Controls.Add(this.RFBtnSetdata);
            this.RFReaderGroupBox.Controls.Add(this.label1);
            this.RFReaderGroupBox.Controls.Add(this.label2);
            this.RFReaderGroupBox.Controls.Add(this.label3);
            this.RFReaderGroupBox.Controls.Add(this.RFTxtReaderid);
            this.RFReaderGroupBox.Controls.Add(this.RFTxtPort);
            this.RFReaderGroupBox.Controls.Add(this.RFCmbPorttype);
            this.RFReaderGroupBox.Location = new System.Drawing.Point(8, 69);
            this.RFReaderGroupBox.Name = "RFReaderGroupBox";
            this.RFReaderGroupBox.Size = new System.Drawing.Size(662, 356);
            this.RFReaderGroupBox.TabIndex = 0;
            this.RFReaderGroupBox.TabStop = false;
            this.RFReaderGroupBox.Text = "Check Reader";
            this.RFReaderGroupBox.Enter += new System.EventHandler(this.RFReaderGroupBox_Enter);
            // 
            // RFReaderBtnRecep_window
            // 
            this.RFReaderBtnRecep_window.Enabled = false;
            this.RFReaderBtnRecep_window.Location = new System.Drawing.Point(157, 224);
            this.RFReaderBtnRecep_window.Name = "RFReaderBtnRecep_window";
            this.RFReaderBtnRecep_window.Size = new System.Drawing.Size(120, 31);
            this.RFReaderBtnRecep_window.TabIndex = 59;
            this.RFReaderBtnRecep_window.Text = "Receptionst Window";
            this.RFReaderBtnRecep_window.UseVisualStyleBackColor = true;
            this.RFReaderBtnRecep_window.Click += new System.EventHandler(this.RFReaderBtnRecep_window_Click);
            // 
            // RFReader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(682, 458);
            this.Controls.Add(this.Recption_daily_attend);
            this.Controls.Add(this.RFReaderGroupBox);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RFReader";
            this.Text = "RFReader";
            this.Activated += new System.EventHandler(this.RFReader_Activated);
            this.Load += new System.EventHandler(this.RFReader_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.RFReaderGroupBox.ResumeLayout(false);
            this.RFReaderGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.Label Recption_daily_attend;
        public System.Windows.Forms.ComboBox RFCmbPorttype;
        private System.Windows.Forms.TextBox RFTxtPort;
        public System.Windows.Forms.TextBox RFTxtReaderid;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button RFBtnSetdata;
        private System.Windows.Forms.Button RFBtnReaderresponse;
        private System.Windows.Forms.Button RFBtnConnect;
        private System.Windows.Forms.ListBox RfListbox1;
        private System.Windows.Forms.Button RfBtnClear1;
        private System.Windows.Forms.Button RFBtnDisconnect;
        private System.Windows.Forms.Label LblShow;
        private System.Windows.Forms.GroupBox RFReaderGroupBox;
        private System.Windows.Forms.Button RFReaderBtnRecep_window;
    }
}