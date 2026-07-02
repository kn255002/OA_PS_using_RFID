using System.Windows.Forms;
namespace OA_PS_RFID
{
    partial class RFTag_Loader
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        /// 
        
        Deo_Module.Update_Database db_obj = new Deo_Module.Update_Database();
        string gv_tagid;

        private void Tagadd_fun(string tagid)
        {
            string sqlstmt = "Insert into Tag_ids Values ("+tagid+",'"+Deo_Module.Safeinfo.Empname+"')";
            if (db_obj.Dml_Updatefun(sqlstmt) >= 1)
            {
                //MessageBox.Show("Record Inserted Successfully.", "Record Inserted", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
               
            TxtRF_tag_id.Text = "";
        }
       
















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
            System.Windows.Forms.Label LblRF_tag_id;
            this.BtnLoad = new System.Windows.Forms.Button();
            this.BtnExit = new System.Windows.Forms.Button();
            this.LblEntered_id = new System.Windows.Forms.Label();
            this.TxtRF_tag_id = new System.Windows.Forms.TextBox();
            this.Gridvu_TagIds = new System.Windows.Forms.DataGridView();
            this.BtnEdit = new System.Windows.Forms.DataGridViewButtonColumn();
            this.BtnDelete = new System.Windows.Forms.DataGridViewButtonColumn();
            this.Tabl_Control = new System.Windows.Forms.TabControl();
            this.Tab_Tag_addtion = new System.Windows.Forms.TabPage();
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.Tab_Tag_allocation = new System.Windows.Forms.TabPage();
            this.BtnRftag_Loader_main = new System.Windows.Forms.Button();
            this.Menue_RFtag_allocation = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tagAllocationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mainToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.LblLoginame = new System.Windows.Forms.Label();
            this.LblLoginstatus = new System.Windows.Forms.Label();
            this.LblNullname = new System.Windows.Forms.Label();
            this.LblNullogin = new System.Windows.Forms.Label();
            LblRF_tag_id = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.Gridvu_TagIds)).BeginInit();
            this.Tabl_Control.SuspendLayout();
            this.Tab_Tag_addtion.SuspendLayout();
            this.GroupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.Menue_RFtag_allocation.SuspendLayout();
            this.SuspendLayout();
            // 
            // LblRF_tag_id
            // 
            LblRF_tag_id.AutoSize = true;
            LblRF_tag_id.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            LblRF_tag_id.Location = new System.Drawing.Point(21, 40);
            LblRF_tag_id.Name = "LblRF_tag_id";
            LblRF_tag_id.Size = new System.Drawing.Size(55, 13);
            LblRF_tag_id.TabIndex = 19;
            LblRF_tag_id.Text = "RF Tag Id";
            // 
            // BtnLoad
            // 
            this.BtnLoad.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnLoad.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.BtnLoad.Location = new System.Drawing.Point(251, 35);
            this.BtnLoad.Name = "BtnLoad";
            this.BtnLoad.Size = new System.Drawing.Size(56, 23);
            this.BtnLoad.TabIndex = 18;
            this.BtnLoad.Text = "Load";
            this.BtnLoad.UseVisualStyleBackColor = true;
            this.BtnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // BtnExit
            // 
            this.BtnExit.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.BtnExit.Location = new System.Drawing.Point(884, 696);
            this.BtnExit.Name = "BtnExit";
            this.BtnExit.Size = new System.Drawing.Size(75, 23);
            this.BtnExit.TabIndex = 17;
            this.BtnExit.Text = "Exit";
            this.BtnExit.UseVisualStyleBackColor = true;
            this.BtnExit.Click += new System.EventHandler(this.BtnExit_Click);
            // 
            // LblEntered_id
            // 
            this.LblEntered_id.AutoSize = true;
            this.LblEntered_id.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.LblEntered_id.Location = new System.Drawing.Point(7, 22);
            this.LblEntered_id.Name = "LblEntered_id";
            this.LblEntered_id.Size = new System.Drawing.Size(104, 13);
            this.LblEntered_id.TabIndex = 15;
            this.LblEntered_id.Text = "Loaded RF Tag IDs.";
            // 
            // TxtRF_tag_id
            // 
            this.TxtRF_tag_id.Location = new System.Drawing.Point(82, 36);
            this.TxtRF_tag_id.Name = "TxtRF_tag_id";
            this.TxtRF_tag_id.Size = new System.Drawing.Size(163, 20);
            this.TxtRF_tag_id.TabIndex = 14;
            this.TxtRF_tag_id.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtRF_tag_id_KeyPress);
            // 
            // Gridvu_TagIds
            // 
            this.Gridvu_TagIds.AllowUserToAddRows = false;
            this.Gridvu_TagIds.AllowUserToDeleteRows = false;
            this.Gridvu_TagIds.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Gridvu_TagIds.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.BtnEdit,
            this.BtnDelete});
            this.Gridvu_TagIds.Location = new System.Drawing.Point(24, 38);
            this.Gridvu_TagIds.Name = "Gridvu_TagIds";
            this.Gridvu_TagIds.ReadOnly = true;
            this.Gridvu_TagIds.Size = new System.Drawing.Size(518, 357);
            this.Gridvu_TagIds.TabIndex = 21;
            this.Gridvu_TagIds.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.Gridvu_TagIds_CellClick);
            // 
            // BtnEdit
            // 
            this.BtnEdit.HeaderText = "Edit";
            this.BtnEdit.Name = "BtnEdit";
            this.BtnEdit.ReadOnly = true;
            this.BtnEdit.Text = "Edit";
            this.BtnEdit.UseColumnTextForButtonValue = true;
            // 
            // BtnDelete
            // 
            this.BtnDelete.HeaderText = "Delete";
            this.BtnDelete.Name = "BtnDelete";
            this.BtnDelete.ReadOnly = true;
            this.BtnDelete.Text = "Delete";
            this.BtnDelete.UseColumnTextForButtonValue = true;
            // 
            // Tabl_Control
            // 
            this.Tabl_Control.Controls.Add(this.Tab_Tag_addtion);
            this.Tabl_Control.Controls.Add(this.Tab_Tag_allocation);
            this.Tabl_Control.Location = new System.Drawing.Point(21, 65);
            this.Tabl_Control.Name = "Tabl_Control";
            this.Tabl_Control.SelectedIndex = 0;
            this.Tabl_Control.Size = new System.Drawing.Size(940, 623);
            this.Tabl_Control.TabIndex = 22;
            this.Tabl_Control.Selected += new System.Windows.Forms.TabControlEventHandler(this.Tabl_Control_Selected);
            // 
            // Tab_Tag_addtion
            // 
            this.Tab_Tag_addtion.BackColor = System.Drawing.Color.AliceBlue;
            this.Tab_Tag_addtion.Controls.Add(this.GroupBox2);
            this.Tab_Tag_addtion.Controls.Add(this.groupBox1);
            this.Tab_Tag_addtion.Location = new System.Drawing.Point(4, 22);
            this.Tab_Tag_addtion.Name = "Tab_Tag_addtion";
            this.Tab_Tag_addtion.Padding = new System.Windows.Forms.Padding(3);
            this.Tab_Tag_addtion.Size = new System.Drawing.Size(932, 597);
            this.Tab_Tag_addtion.TabIndex = 0;
            this.Tab_Tag_addtion.Text = "Tags Addition";
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.Gridvu_TagIds);
            this.GroupBox2.Controls.Add(this.LblEntered_id);
            this.GroupBox2.Location = new System.Drawing.Point(72, 140);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(561, 428);
            this.GroupBox2.TabIndex = 23;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Entered Tags";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnCancel);
            this.groupBox1.Controls.Add(this.BtnLoad);
            this.groupBox1.Controls.Add(this.TxtRF_tag_id);
            this.groupBox1.Controls.Add(LblRF_tag_id);
            this.groupBox1.Location = new System.Drawing.Point(72, 34);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(561, 100);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add Tags";
            // 
            // BtnCancel
            // 
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.BtnCancel.Location = new System.Drawing.Point(313, 35);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(56, 23);
            this.BtnCancel.TabIndex = 20;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // Tab_Tag_allocation
            // 
            this.Tab_Tag_allocation.Location = new System.Drawing.Point(4, 22);
            this.Tab_Tag_allocation.Name = "Tab_Tag_allocation";
            this.Tab_Tag_allocation.Padding = new System.Windows.Forms.Padding(3);
            this.Tab_Tag_allocation.Size = new System.Drawing.Size(824, 484);
            this.Tab_Tag_allocation.TabIndex = 1;
            this.Tab_Tag_allocation.Text = "Tag Allocation";
            this.Tab_Tag_allocation.UseVisualStyleBackColor = true;
            // 
            // BtnRftag_Loader_main
            // 
            this.BtnRftag_Loader_main.Location = new System.Drawing.Point(803, 696);
            this.BtnRftag_Loader_main.Name = "BtnRftag_Loader_main";
            this.BtnRftag_Loader_main.Size = new System.Drawing.Size(75, 23);
            this.BtnRftag_Loader_main.TabIndex = 54;
            this.BtnRftag_Loader_main.Text = "Main";
            this.BtnRftag_Loader_main.UseVisualStyleBackColor = true;
            this.BtnRftag_Loader_main.Click += new System.EventHandler(this.BtnRftag_Loader_main_Click);
            // 
            // Menue_RFtag_allocation
            // 
            this.Menue_RFtag_allocation.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.Menue_RFtag_allocation.Location = new System.Drawing.Point(0, 0);
            this.Menue_RFtag_allocation.Name = "Menue_RFtag_allocation";
            this.Menue_RFtag_allocation.Size = new System.Drawing.Size(1009, 24);
            this.Menue_RFtag_allocation.TabIndex = 55;
            this.Menue_RFtag_allocation.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tagAllocationToolStripMenuItem,
            this.mainToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "File";
            this.fileToolStripMenuItem.MouseEnter += new System.EventHandler(this.fileToolStripMenuItem_MouseEnter);
            // 
            // tagAllocationToolStripMenuItem
            // 
            this.tagAllocationToolStripMenuItem.Name = "tagAllocationToolStripMenuItem";
            this.tagAllocationToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.tagAllocationToolStripMenuItem.Text = "Tag Allocation";
            this.tagAllocationToolStripMenuItem.Click += new System.EventHandler(this.tagAllocationToolStripMenuItem_Click);
            // 
            // mainToolStripMenuItem
            // 
            this.mainToolStripMenuItem.Name = "mainToolStripMenuItem";
            this.mainToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.mainToolStripMenuItem.Text = "Main";
            this.mainToolStripMenuItem.Click += new System.EventHandler(this.mainToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem1});
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.aboutToolStripMenuItem.Text = "Help";
            // 
            // aboutToolStripMenuItem1
            // 
            this.aboutToolStripMenuItem1.Name = "aboutToolStripMenuItem1";
            this.aboutToolStripMenuItem1.Size = new System.Drawing.Size(114, 22);
            this.aboutToolStripMenuItem1.Text = "About";
            this.aboutToolStripMenuItem1.Click += new System.EventHandler(this.aboutToolStripMenuItem1_Click);
            // 
            // LblLoginame
            // 
            this.LblLoginame.AllowDrop = true;
            this.LblLoginame.AutoSize = true;
            this.LblLoginame.BackColor = System.Drawing.Color.WhiteSmoke;
            this.LblLoginame.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblLoginame.Location = new System.Drawing.Point(18, 701);
            this.LblLoginame.Name = "LblLoginame";
            this.LblLoginame.Size = new System.Drawing.Size(67, 13);
            this.LblLoginame.TabIndex = 26;
            this.LblLoginame.Text = "Login Name:";
            // 
            // LblLoginstatus
            // 
            this.LblLoginstatus.AllowDrop = true;
            this.LblLoginstatus.AutoSize = true;
            this.LblLoginstatus.BackColor = System.Drawing.Color.WhiteSmoke;
            this.LblLoginstatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblLoginstatus.Location = new System.Drawing.Point(176, 701);
            this.LblLoginstatus.Name = "LblLoginstatus";
            this.LblLoginstatus.Size = new System.Drawing.Size(69, 13);
            this.LblLoginstatus.TabIndex = 25;
            this.LblLoginstatus.Text = "Login Status:";
            // 
            // LblNullname
            // 
            this.LblNullname.AutoSize = true;
            this.LblNullname.BackColor = System.Drawing.Color.WhiteSmoke;
            this.LblNullname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblNullname.ForeColor = System.Drawing.SystemColors.ControlText;
            this.LblNullname.Location = new System.Drawing.Point(102, 701);
            this.LblNullname.Name = "LblNullname";
            this.LblNullname.Size = new System.Drawing.Size(56, 13);
            this.LblNullname.TabIndex = 27;
            this.LblNullname.Text = "Null Name";
            // 
            // LblNullogin
            // 
            this.LblNullogin.AutoSize = true;
            this.LblNullogin.BackColor = System.Drawing.Color.WhiteSmoke;
            this.LblNullogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblNullogin.ForeColor = System.Drawing.SystemColors.ControlText;
            this.LblNullogin.Location = new System.Drawing.Point(261, 701);
            this.LblNullogin.Name = "LblNullogin";
            this.LblNullogin.Size = new System.Drawing.Size(54, 13);
            this.LblNullogin.TabIndex = 24;
            this.LblNullogin.Text = "Null Login";
            // 
            // RFTag_Loader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Desktop;
            this.ClientSize = new System.Drawing.Size(1009, 746);
            this.Controls.Add(this.LblNullogin);
            this.Controls.Add(this.BtnRftag_Loader_main);
            this.Controls.Add(this.LblNullname);
            this.Controls.Add(this.LblLoginstatus);
            this.Controls.Add(this.Tabl_Control);
            this.Controls.Add(this.LblLoginame);
            this.Controls.Add(this.BtnExit);
            this.Controls.Add(this.Menue_RFtag_allocation);
            this.MainMenuStrip = this.Menue_RFtag_allocation;
            this.MaximizeBox = false;
            this.Name = "RFTag_Loader";
            this.Text = "HR RFTag Allocation";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.RFTag_Loader_Activated);
            this.Load += new System.EventHandler(this.RFTag_Loader_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Gridvu_TagIds)).EndInit();
            this.Tabl_Control.ResumeLayout(false);
            this.Tab_Tag_addtion.ResumeLayout(false);
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.Menue_RFtag_allocation.ResumeLayout(false);
            this.Menue_RFtag_allocation.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnLoad;
        private System.Windows.Forms.Button BtnExit;
        private System.Windows.Forms.Label LblEntered_id;
        private System.Windows.Forms.TextBox TxtRF_tag_id;
        private System.Windows.Forms.DataGridView Gridvu_TagIds;
        private System.Windows.Forms.TabControl Tabl_Control;
        private System.Windows.Forms.TabPage Tab_Tag_addtion;
        private System.Windows.Forms.TabPage Tab_Tag_allocation;
        private System.Windows.Forms.Button BtnRftag_Loader_main;
        private System.Windows.Forms.GroupBox GroupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.MenuStrip Menue_RFtag_allocation;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tagAllocationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mainToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem1;
        private DataGridViewButtonColumn BtnEdit;
        private DataGridViewButtonColumn BtnDelete;
        private Label LblLoginame;
        private Label LblLoginstatus;
        private Label LblNullname;
        private Label LblNullogin;
        private Button BtnCancel;
    }
}