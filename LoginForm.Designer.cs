using System;   //Exception
using System.Data;
using System.Data.OracleClient; //OracleCon and etc.
using System.Windows.Forms;     //MessageBox and etc.

namespace Deo_Module
{
    partial class LoginForm
    {
        // Required designer variable.        
        //Data members only.

        private Update_Database gl_dbaseclsobj = new Update_Database();
        
        /*****************************************************************************************/
        /**********************************Now my Functions.**************************************/
        void Login_verifyfun()
         {
            //Project_Connection con = new Project_Connection();            
            string login_status=null;
            string job_status=null;
            string emp_name = null;
            while (true)
             {//Infinite Loop, Untill BREAK STATMENT CALL.
                if (TxtLoginname.Text == "" || TxtPassword.Text == "")
                {
                    MessageBox.Show("Login name or Passowrd can not be Blanked.", "Empty Login/Password", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtPassword.Text = null;
                    TxtLoginname.Text = null;
                    LPlsenterloginame.Show();
                    LPlsenterpwd.Show();
                    break;
                }
                else
                {
                    try
                    {                        
                        gl_dbaseclsobj.cls_con.Open();
                        OracleCommand cmd = new OracleCommand();
                        cmd.Connection = gl_dbaseclsobj.cls_con;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "Procedure_verifylogin";

                        OracleParameter p1 = new OracleParameter("Pv_Login", OracleType.Number);
                        p1.Direction = ParameterDirection.Input;
                        p1.Value = int.Parse(TxtLoginname.Text);
                        cmd.Parameters.Add(p1);

                        OracleParameter p2 = new OracleParameter("Pv_Password", OracleType.VarChar);
                        p2.Direction = ParameterDirection.Input;
                        p2.Value = TxtPassword.Text.ToUpper();
                        cmd.Parameters.Add(p2);

                        OracleParameter p3 = new OracleParameter("Pv_Empname", OracleType.VarChar, 25);
                        p3.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p3);

                        OracleParameter p4 = new OracleParameter("Pv_Job_Id", OracleType.VarChar, 20);
                        p4.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p4);

                        OracleParameter p5 = new OracleParameter("Pv_Status", OracleType.VarChar, 5);
                        p5.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p5);

                        cmd.ExecuteNonQuery();
                        emp_name = p3.Value.ToString();
                        job_status = p4.Value.ToString();
                        login_status = p5.Value.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TxtPassword.Text = null;
                        TxtLoginname.Text = null;                    
                    }
                    finally
                    {
                        gl_dbaseclsobj.cls_con.Close();
                    }

                    //You can delete from start to End.
                    //Start 12-July-2007
                    if (login_status == null)
                    {
                        MessageBox.Show("Connection Problem Found.", "Account Status", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TxtPassword.Text = null;
                        TxtLoginname.Text = null;                    
                    }
                    else
                    {//You can delete from start to End.
                     //End   12-July-2007

                        switch (login_status.ToUpper())
                        {
                            case "ALLOW":
                                {
                                    switch (job_status)
                                    {
                                        case "SCANNER":
                                            {
                                                Safeinfo.loginid = int.Parse(TxtLoginname.Text);    //Saving for feature display.                 
                                                Safeinfo.password = TxtPassword.Text.ToUpper();     //Saving for Pwd matching at 'changepwd' Forms display.
                                                Safeinfo.status = job_status;                       //Saving for feature display.
                                                Safeinfo.Empname = emp_name;                        //Saving for feature display.                          

                                                Scanner newform_obj = new Scanner();               //To display create form's obj.
                                                newform_obj.Show();                                //Show new form.
                                                TxtPassword.Text = null;
                                                TxtLoginname.Text = null;
                                                this.Hide();
                                                break;
                                            }
                                        case "DATA ENTRY OPERATOR":
                                            {
                                                Safeinfo.loginid = int.Parse(TxtLoginname.Text);    //Saving for feature display.                 
                                                Safeinfo.password = TxtPassword.Text.ToUpper();     //Saving for Pwd matching at 'changepwd' Forms display.
                                                Safeinfo.status = job_status;                       //Saving for feature display.
                                                Safeinfo.Empname = emp_name;                        //Saving for feature display.                 

                                                Dataentry newform_obj = new Dataentry();            //To display create form's obj.
                                                newform_obj.Show();                                 //Show new form.                        this.Hide();                                    
                                                TxtPassword.Text = null;
                                                TxtLoginname.Text = null;
                                                this.Hide();
                                                break;
                                            }
                                        case "HR":
                                            {
                                                Safeinfo.loginid = int.Parse(TxtLoginname.Text);    //Saving for feature display.                 
                                                Safeinfo.password = TxtPassword.Text.ToUpper();     //Saving for Pwd matching at 'changepwd' Forms display.
                                                Safeinfo.status = job_status;                       //Saving for feature display.
                                                Safeinfo.Empname = emp_name;                        //Saving for feature display.                 

                                                OA_PS_RFID.RFTag_Loader newform_obj = new OA_PS_RFID.RFTag_Loader(); //Its mean Load DEO Form for HR purpose.                                                
                                                newform_obj.Show();                                 //Show new form.                        this.Hide();
                                                TxtPassword.Text = null;
                                                TxtLoginname.Text = null;
                                                this.Hide();
                                                break;

                                                //Safeinfo.loginid = int.Parse(TxtLoginname.Text);    //Saving for feature display.                 
                                                //Safeinfo.password = TxtPassword.Text.ToUpper();     //Saving for Pwd matching at 'changepwd' Forms display.
                                                //Safeinfo.status = job_status;                       //Saving for feature display.
                                                //Safeinfo.Empname = emp_name;                        //Saving for feature display.                 

                                                //Dataentry newform_obj = new Dataentry();            //To display create form's obj.
                                                //newform_obj.Show();                                 //Show new form.                        this.Hide();
                                                //TxtPassword.Text = null;
                                                //TxtLoginname.Text = null;
                                                //this.Hide();
                                                //break;
                                            }
                                        case "ADMINISTRATOR":
                                            {
                                                Safeinfo.loginid = int.Parse(TxtLoginname.Text);
                                                Safeinfo.password = TxtPassword.Text.ToUpper();     //Saving for Pwd matching at 'changepwd' Forms display.
                                                Safeinfo.status = job_status;                       //Saving for feature display.
                                                Safeinfo.Empname = emp_name;                        //Saving for feature display.                 

                                                Admin newform_obj = new Admin();                    //To display create form's obj.
                                                newform_obj.Show();                                 //Show new form.                        this.Hide();
                                                TxtPassword.Text = null;
                                                TxtLoginname.Text = null;
                                                this.Hide();
                                                break;
                                            }
                                        case "RECEPTIONIST":
                                            {
                                                Safeinfo.loginid = int.Parse(TxtLoginname.Text);
                                                Safeinfo.password = TxtPassword.Text.ToUpper();     //Saving for Pwd matching at 'changepwd' Forms display.
                                                Safeinfo.status = job_status;                       //Saving for feature display.
                                                Safeinfo.Empname = emp_name;                        //Saving for feature display.                 

                                                OA_PS_RFID.RFReader newform_obj = new OA_PS_RFID.RFReader();                    //To display create form's obj.
                                                newform_obj.Show();                                 //Show new form.                        this.Hide();
                                                TxtPassword.Text = null;
                                                TxtLoginname.Text = null;
                                                this.Hide();
                                                break;
                                            }
                                    }//End JOB_STATUS Switch.
                                    break;
                                }//End ALLOW CASE.
                            case "LOCK":
                                {
                                    MessageBox.Show("Your account is Temporary locked.\nPlease Contact your Administrator.", "Account Status", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    TxtLoginname.Text = null;
                                    TxtPassword.Text = null;
                                    break;
                                }
                            case "SORRY":
                                {
                                    MessageBox.Show("Login Name/Password denied.", "Account Status", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    LPlsenterloginame.Show();
                                    LPlsenterpwd.Show();
                                    TxtPassword.Text = null;
                                    TxtLoginname.Text = null;
                                    break;
                                }
                            case "":
                                {
                                    MessageBox.Show("Login Name is not valid.", "Account Status", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    LPlsenterloginame.Show();
                                    LPlsenterpwd.Show();
                                    TxtPassword.Text = null;
                                    TxtLoginname.Text = null;
                                    break;
                                }
                        }//End Login_STATUS Switch.                
                    }//End else
                }//End else.
                break;
            }//End of While
        }//Func End.

    
        /*****************************************************************************************/
        /*****************************************************************************************/
        /// </summary>
        /// <param name="disposing"></param>
        /// 

        private System.ComponentModel.IContainer components = null;
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
            this.LPlsenterpwd = new System.Windows.Forms.Label();
            this.LPlsenterloginame = new System.Windows.Forms.Label();
            this.LPassword = new System.Windows.Forms.Label();
            this.LLoginname = new System.Windows.Forms.Label();
            this.BtnChangepwd = new System.Windows.Forms.Button();
            this.TxtPassword = new System.Windows.Forms.TextBox();
            this.TxtLoginname = new System.Windows.Forms.TextBox();
            this.BConnect = new System.Windows.Forms.Button();
            this.BExit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.connectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LblBetaversion = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.TxtLoginSrvsname = new System.Windows.Forms.TextBox();
            this.LblLoginSrvsname = new System.Windows.Forms.Label();
            this.TxtDBServerName = new System.Windows.Forms.TextBox();
            this.LblServerName = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // LPlsenterpwd
            // 
            this.LPlsenterpwd.AutoSize = true;
            this.LPlsenterpwd.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.LPlsenterpwd.ForeColor = System.Drawing.Color.Red;
            this.LPlsenterpwd.Location = new System.Drawing.Point(390, 149);
            this.LPlsenterpwd.Name = "LPlsenterpwd";
            this.LPlsenterpwd.Size = new System.Drawing.Size(116, 13);
            this.LPlsenterpwd.TabIndex = 15;
            this.LPlsenterpwd.Text = "Please Enter Password";
            // 
            // LPlsenterloginame
            // 
            this.LPlsenterloginame.AutoSize = true;
            this.LPlsenterloginame.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.LPlsenterloginame.ForeColor = System.Drawing.Color.Red;
            this.LPlsenterloginame.Location = new System.Drawing.Point(390, 125);
            this.LPlsenterloginame.Name = "LPlsenterloginame";
            this.LPlsenterloginame.Size = new System.Drawing.Size(127, 13);
            this.LPlsenterloginame.TabIndex = 14;
            this.LPlsenterloginame.Text = "Please Enter Login Name";
            // 
            // LPassword
            // 
            this.LPassword.AutoSize = true;
            this.LPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LPassword.Location = new System.Drawing.Point(92, 149);
            this.LPassword.Name = "LPassword";
            this.LPassword.Size = new System.Drawing.Size(61, 13);
            this.LPassword.TabIndex = 5;
            this.LPassword.Text = "Password";
            // 
            // LLoginname
            // 
            this.LLoginname.AutoSize = true;
            this.LLoginname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LLoginname.Location = new System.Drawing.Point(92, 125);
            this.LLoginname.Name = "LLoginname";
            this.LLoginname.Size = new System.Drawing.Size(74, 13);
            this.LLoginname.TabIndex = 3;
            this.LLoginname.Text = "Login Name";
            // 
            // BtnChangepwd
            // 
            this.BtnChangepwd.Location = new System.Drawing.Point(275, 230);
            this.BtnChangepwd.Name = "BtnChangepwd";
            this.BtnChangepwd.Size = new System.Drawing.Size(102, 25);
            this.BtnChangepwd.TabIndex = 12;
            this.BtnChangepwd.Text = "Change Password";
            this.BtnChangepwd.UseVisualStyleBackColor = true;
            this.BtnChangepwd.Click += new System.EventHandler(this.BtnChangepwd_Click);
            // 
            // TxtPassword
            // 
            this.TxtPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtPassword.Location = new System.Drawing.Point(200, 145);
            this.TxtPassword.Name = "TxtPassword";
            this.TxtPassword.PasswordChar = '*';
            this.TxtPassword.Size = new System.Drawing.Size(177, 21);
            this.TxtPassword.TabIndex = 6;
            this.TxtPassword.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtPassword_KeyPress);
            // 
            // TxtLoginname
            // 
            this.TxtLoginname.Location = new System.Drawing.Point(200, 122);
            this.TxtLoginname.Name = "TxtLoginname";
            this.TxtLoginname.Size = new System.Drawing.Size(177, 20);
            this.TxtLoginname.TabIndex = 4;
            this.TxtLoginname.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtLoginname_KeyPress);
            // 
            // BConnect
            // 
            this.BConnect.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BConnect.ForeColor = System.Drawing.SystemColors.ControlText;
            this.BConnect.Location = new System.Drawing.Point(200, 230);
            this.BConnect.Name = "BConnect";
            this.BConnect.Size = new System.Drawing.Size(72, 25);
            this.BConnect.TabIndex = 7;
            this.BConnect.Text = "Connect";
            this.BConnect.UseVisualStyleBackColor = true;
            this.BConnect.Click += new System.EventHandler(this.connectToolStripMenuItem_Click);
            // 
            // BExit
            // 
            this.BExit.Location = new System.Drawing.Point(303, 261);
            this.BExit.Name = "BExit";
            this.BExit.Size = new System.Drawing.Size(73, 24);
            this.BExit.TabIndex = 13;
            this.BExit.Text = "Exit";
            this.BExit.UseVisualStyleBackColor = true;
            this.BExit.Click += new System.EventHandler(this.button3_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(77, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(401, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "ONLINE ATTENDANCE AND PAYROLL SYSTEM";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(185, 68);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(146, 20);
            this.label3.TabIndex = 2;
            this.label3.Text = "THROUGH RFID";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Black;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Location = new System.Drawing.Point(0, 94);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(543, 1);
            this.pictureBox1.TabIndex = 52;
            this.pictureBox1.TabStop = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(538, 24);
            this.menuStrip1.TabIndex = 20;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.connectToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // connectToolStripMenuItem
            // 
            this.connectToolStripMenuItem.Name = "connectToolStripMenuItem";
            this.connectToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.connectToolStripMenuItem.Text = "Connect";
            this.connectToolStripMenuItem.Click += new System.EventHandler(this.connectToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
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
            // LblBetaversion
            // 
            this.LblBetaversion.AutoSize = true;
            this.LblBetaversion.ForeColor = System.Drawing.Color.Red;
            this.LblBetaversion.Location = new System.Drawing.Point(8, 272);
            this.LblBetaversion.Name = "LblBetaversion";
            this.LblBetaversion.Size = new System.Drawing.Size(143, 13);
            this.LblBetaversion.TabIndex = 18;
            this.LblBetaversion.Text = "Alpha Release: 21-Aug-2007";
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Black;
            this.pictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox2.Location = new System.Drawing.Point(-2, 24);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(540, 2);
            this.pictureBox2.TabIndex = 55;
            this.pictureBox2.TabStop = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(8, 294);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 13);
            this.label4.TabIndex = 19;
            this.label4.Text = "Version 1.0";
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(0, 26);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(538, 24);
            this.label5.TabIndex = 0;
            this.label5.Text = "                             MIA ENTERPRISES";
            // 
            // TxtLoginSrvsname
            // 
            this.TxtLoginSrvsname.Location = new System.Drawing.Point(200, 174);
            this.TxtLoginSrvsname.Name = "TxtLoginSrvsname";
            this.TxtLoginSrvsname.Size = new System.Drawing.Size(177, 20);
            this.TxtLoginSrvsname.TabIndex = 8;
            this.TxtLoginSrvsname.Text = "Orcl9i";
            this.TxtLoginSrvsname.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtLoginSrvsname_KeyPress_1);
            // 
            // LblLoginSrvsname
            // 
            this.LblLoginSrvsname.AutoSize = true;
            this.LblLoginSrvsname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblLoginSrvsname.Location = new System.Drawing.Point(92, 177);
            this.LblLoginSrvsname.Name = "LblLoginSrvsname";
            this.LblLoginSrvsname.Size = new System.Drawing.Size(86, 13);
            this.LblLoginSrvsname.TabIndex = 9;
            this.LblLoginSrvsname.Text = "Service Name";
            // 
            // TxtDBServerName
            // 
            this.TxtDBServerName.Location = new System.Drawing.Point(200, 204);
            this.TxtDBServerName.Name = "TxtDBServerName";
            this.TxtDBServerName.Size = new System.Drawing.Size(177, 20);
            this.TxtDBServerName.TabIndex = 11;
            this.TxtDBServerName.Text = "Pavilion-211178";
            this.TxtDBServerName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtDBServerName_KeyPress);
            // 
            // LblServerName
            // 
            this.LblServerName.AutoSize = true;
            this.LblServerName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblServerName.Location = new System.Drawing.Point(92, 207);
            this.LblServerName.Name = "LblServerName";
            this.LblServerName.Size = new System.Drawing.Size(102, 13);
            this.LblServerName.TabIndex = 10;
            this.LblServerName.Text = "Database Server";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(453, 283);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(73, 24);
            this.button1.TabIndex = 56;
            this.button1.Text = "test";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // LoginForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(538, 317);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.LblServerName);
            this.Controls.Add(this.TxtDBServerName);
            this.Controls.Add(this.LblLoginSrvsname);
            this.Controls.Add(this.TxtLoginSrvsname);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.LblBetaversion);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.LPassword);
            this.Controls.Add(this.BConnect);
            this.Controls.Add(this.LLoginname);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.BExit);
            this.Controls.Add(this.LPlsenterpwd);
            this.Controls.Add(this.LPlsenterloginame);
            this.Controls.Add(this.BtnChangepwd);
            this.Controls.Add(this.TxtPassword);
            this.Controls.Add(this.TxtLoginname);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "LoginForm";
            this.Text = "M.I.A Login";
            this.Activated += new System.EventHandler(this.LoginForm_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.LoginForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label LPlsenterpwd;
        private System.Windows.Forms.Label LPlsenterloginame;
        private System.Windows.Forms.Label LPassword;
        private System.Windows.Forms.Label LLoginname;
        private System.Windows.Forms.Button BtnChangepwd;
        private System.Windows.Forms.TextBox TxtPassword;
        private System.Windows.Forms.TextBox TxtLoginname;
        private System.Windows.Forms.Button BConnect;
        private System.Windows.Forms.Button BExit;
        private System.Windows.Forms.Label label1;
        private Label label3;
        private PictureBox pictureBox1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem connectToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private Label LblBetaversion;
        private PictureBox pictureBox2;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripMenuItem aboutToolStripMenuItem;
        private Label label4;
        private Label label5;
        private TextBox TxtLoginSrvsname;
        private Label LblLoginSrvsname;
        private TextBox TxtDBServerName;
        private Label LblServerName;
        private Button button1;

    }
}

