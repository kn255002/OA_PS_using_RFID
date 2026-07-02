namespace Deo_Module
{
    partial class Changepwd
    {
        // Required designer variable.
        //
        //
        //Data members only.

        private Update_Database gl_dbaseclsobj = new Update_Database();
        /*****************************************************************************************/
        /**********************************Now my Functions.**************************************/


        /*****************************************************************************************/
        /*****************************************************************************************/
        //
        //
        //
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
            this.BtnChangepwd = new System.Windows.Forms.Button();
            this.TxtOldpwd = new System.Windows.Forms.TextBox();
            this.TxtNewpwd = new System.Windows.Forms.TextBox();
            this.LblOldpwd = new System.Windows.Forms.Label();
            this.LblNewpwd = new System.Windows.Forms.Label();
            this.LblRetypenewpwd = new System.Windows.Forms.Label();
            this.TxtRetypenewpwd = new System.Windows.Forms.TextBox();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.DeleteFormsRichTxtDelforms = new System.Windows.Forms.RichTextBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnChangepwd
            // 
            this.BtnChangepwd.Location = new System.Drawing.Point(224, 155);
            this.BtnChangepwd.Name = "BtnChangepwd";
            this.BtnChangepwd.Size = new System.Drawing.Size(102, 27);
            this.BtnChangepwd.TabIndex = 6;
            this.BtnChangepwd.Text = "Change Password";
            this.BtnChangepwd.UseVisualStyleBackColor = true;
            this.BtnChangepwd.Click += new System.EventHandler(this.BtnChangepwd_Click);
            // 
            // TxtOldpwd
            // 
            this.TxtOldpwd.Location = new System.Drawing.Point(181, 42);
            this.TxtOldpwd.Name = "TxtOldpwd";
            this.TxtOldpwd.Size = new System.Drawing.Size(234, 20);
            this.TxtOldpwd.TabIndex = 3;
            this.TxtOldpwd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtOldpwd_KeyPress);
            // 
            // TxtNewpwd
            // 
            this.TxtNewpwd.Location = new System.Drawing.Point(181, 66);
            this.TxtNewpwd.Name = "TxtNewpwd";
            this.TxtNewpwd.Size = new System.Drawing.Size(234, 20);
            this.TxtNewpwd.TabIndex = 4;
            this.TxtNewpwd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtNewpwd_KeyPress);
            // 
            // LblOldpwd
            // 
            this.LblOldpwd.AutoSize = true;
            this.LblOldpwd.BackColor = System.Drawing.Color.White;
            this.LblOldpwd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblOldpwd.Location = new System.Drawing.Point(28, 46);
            this.LblOldpwd.Name = "LblOldpwd";
            this.LblOldpwd.Size = new System.Drawing.Size(84, 13);
            this.LblOldpwd.TabIndex = 0;
            this.LblOldpwd.Text = "Old Password";
            // 
            // LblNewpwd
            // 
            this.LblNewpwd.AutoSize = true;
            this.LblNewpwd.BackColor = System.Drawing.Color.White;
            this.LblNewpwd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblNewpwd.Location = new System.Drawing.Point(28, 70);
            this.LblNewpwd.Name = "LblNewpwd";
            this.LblNewpwd.Size = new System.Drawing.Size(90, 13);
            this.LblNewpwd.TabIndex = 1;
            this.LblNewpwd.Text = "New Password";
            // 
            // LblRetypenewpwd
            // 
            this.LblRetypenewpwd.AutoSize = true;
            this.LblRetypenewpwd.BackColor = System.Drawing.Color.White;
            this.LblRetypenewpwd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblRetypenewpwd.Location = new System.Drawing.Point(28, 98);
            this.LblRetypenewpwd.Name = "LblRetypenewpwd";
            this.LblRetypenewpwd.Size = new System.Drawing.Size(142, 13);
            this.LblRetypenewpwd.TabIndex = 2;
            this.LblRetypenewpwd.Text = "Re-Type New Password";
            // 
            // TxtRetypenewpwd
            // 
            this.TxtRetypenewpwd.Location = new System.Drawing.Point(181, 94);
            this.TxtRetypenewpwd.Name = "TxtRetypenewpwd";
            this.TxtRetypenewpwd.Size = new System.Drawing.Size(234, 20);
            this.TxtRetypenewpwd.TabIndex = 5;
            this.TxtRetypenewpwd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtRetypenewpwd_KeyPress);
            // 
            // BtnCancel
            // 
            this.BtnCancel.BackColor = System.Drawing.SystemColors.Control;
            this.BtnCancel.Location = new System.Drawing.Point(332, 155);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(83, 27);
            this.BtnCancel.TabIndex = 7;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // DeleteFormsRichTxtDelforms
            // 
            this.DeleteFormsRichTxtDelforms.BackColor = System.Drawing.Color.WhiteSmoke;
            this.DeleteFormsRichTxtDelforms.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DeleteFormsRichTxtDelforms.Location = new System.Drawing.Point(-2, -1);
            this.DeleteFormsRichTxtDelforms.Name = "DeleteFormsRichTxtDelforms";
            this.DeleteFormsRichTxtDelforms.ReadOnly = true;
            this.DeleteFormsRichTxtDelforms.Size = new System.Drawing.Size(515, 29);
            this.DeleteFormsRichTxtDelforms.TabIndex = 57;
            this.DeleteFormsRichTxtDelforms.Text = "           CHANGE PASSWORD REQUEST";
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Black;
            this.pictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox2.Location = new System.Drawing.Point(-2, 28);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(515, 1);
            this.pictureBox2.TabIndex = 58;
            this.pictureBox2.TabStop = false;
            // 
            // Changepwd
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(513, 194);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.DeleteFormsRichTxtDelforms);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.LblRetypenewpwd);
            this.Controls.Add(this.TxtRetypenewpwd);
            this.Controls.Add(this.LblNewpwd);
            this.Controls.Add(this.LblOldpwd);
            this.Controls.Add(this.TxtNewpwd);
            this.Controls.Add(this.TxtOldpwd);
            this.Controls.Add(this.BtnChangepwd);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Changepwd";
            this.Text = "Change Password";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnChangepwd;
        private System.Windows.Forms.TextBox TxtOldpwd;
        private System.Windows.Forms.TextBox TxtNewpwd;
        private System.Windows.Forms.Label LblOldpwd;
        private System.Windows.Forms.Label LblNewpwd;
        private System.Windows.Forms.Label LblRetypenewpwd;
        private System.Windows.Forms.TextBox TxtRetypenewpwd;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.RichTextBox DeleteFormsRichTxtDelforms;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}