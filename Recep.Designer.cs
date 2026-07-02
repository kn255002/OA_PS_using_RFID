using System;                   //Exception
using System.Data;
using System.Data.OracleClient; //OracleCon and etc.
using System.Windows.Forms;     //MessageBox and etc.
using System.IO;                //For files/Steams.
using System.Drawing;           //For Pic/Images.
using System.Runtime.InteropServices;
using SpeechLib;
using System.Threading;

namespace Deo_Module
{
    partial class Recep
    {               
        private Update_Database gl_dbaseclsobj = new Update_Database();
        private SpVoice Bollo_naa;
        

        private static Pgrdrv2.CReaderCtl rdr1 = new Pgrdrv2.CReaderCtl();
        object downloaded_data;


        //private Thread thread = new Thread(new ThreadStart(ThreadFunc));
        //private OA_PS_RFID.DelegateClass.first_delegate delgate_obj = new OA_PS_RFID.DelegateClass.first_delegate(delegate_handlerfun);
        
        //private void downloadedrfid_fun()
        //{//This func run a thread.
        //    thread.IsBackground = true;
        //    thread.Start();//Here we start Thread. to read PORT READER.
        //    LabelThreadresult.Text= var_rfid;
        //}
        /// /*******************************************************************************/
        //public static string delegate_handlerfun(string ss)
        //{
        //    return "Hello " + ss + "!";
        //}
        
        public bool rfdisp_fun2(string strrfid)
        {
            int rfid_id = Convert.ToInt32(strrfid);
            string srch_rftagid = "Select * From Employee Where Rfid_id=" + rfid_id + " And Formstatus_id=6";
            gl_dbaseclsobj.Show_Datafun(srch_rftagid, 0);
            bool recordExist = gl_dbaseclsobj.cls_dr.Read();
            if (recordExist)
            {//If record found then Display into particular fields.
                int temp_formid = int.Parse(gl_dbaseclsobj.cls_dr.GetDecimal(0).ToString());
                LblName.Text = gl_dbaseclsobj.cls_dr.GetString(1);
                LblBps.Text = gl_dbaseclsobj.cls_dr.GetDecimal(20).ToString();
                LblTimestamp.Text = RecepCtlDate.Value.Day.ToString() + "-" + RecepCtlDate.Value.Month.ToString() + "-" + RecepCtlDate.Value.Year.ToString() + " " + RecepCtlDate.Value.TimeOfDay.Hours.ToString() + " : " + RecepCtlDate.Value.TimeOfDay.Minutes.ToString() + " : " + RecepCtlDate.Value.TimeOfDay.Seconds.ToString();

                {   //Picture loading code.
                    OracleLob blob = gl_dbaseclsobj.cls_dr.GetOracleLob(29);      //Column # .      Emp_image
                    Byte[] BLOBData = new Byte[blob.Length];
                    //Read blob data into byte array
                    int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));
                    //Get the primitive byte data into in-memory data stream
                    MemoryStream stmBLOBData = new MemoryStream(BLOBData);
                    //LOADING INTO PICBOX1                            
                    ReceptCtlPicture.Image = Image.FromStream(stmBLOBData);
                    //MessageBox.Show("Now i'm reseting it.");
                    ReceptCtlPicture.SizeMode = PictureBoxSizeMode.StretchImage;
                }
                gl_dbaseclsobj.Show_Datafun(srch_rftagid, 1);//Now close database connection.
                //Now here below Load Department name.
                gl_dbaseclsobj.Dml_UpdateAdapterfun("Select Dept_name from Department Where dept_id=(Select dept_id from Employee where Rfid_id='" + rfid_id + "')");
                LblDepartment.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();

                try
                {
                    gl_dbaseclsobj.cls_con.Open();
                    OracleCommand cmd = new OracleCommand();
                    cmd.Connection = gl_dbaseclsobj.cls_con;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "Mark_Attendance";

                    //Sending Formdid...
                    OracleParameter p1_formid = new OracleParameter("Pv_Rfid", OracleType.VarChar, 8);
                    p1_formid.Direction = ParameterDirection.Input;
                    p1_formid.Value = strrfid;//Data that thread have downloaded from Reader.
                    cmd.Parameters.Add(p1_formid);

                    OracleParameter p2_formid = new OracleParameter("Answer", OracleType.VarChar, 4);
                    p2_formid.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(p2_formid);
                    cmd.ExecuteNonQuery();
                    RecepLblinout.Text = p2_formid.Value.ToString();
                }
                catch (OracleException oex) { MessageBox.Show(oex.Message, "Oracle Error", MessageBoxButtons.OK, MessageBoxIcon.Stop); }
                finally
                {
                    gl_dbaseclsobj.cls_con.Close();
                }
                existance_statusfun(LblName.Text, RecepLblinout.Text, rfid_id);
                return true;
            }//End if.
            else
            {   //Reader HAVE READ RF TAGE. BUT RECORD IS NOT IN OUR DATABASE. SO DON'T DO ANY THING.
                return false;
            }
        }
        //public static void ThreadFunc()
        //{
        //    object downloaded_data;            
        //    //this.Enabled = false; 

        //    downloaded_data = rdr1.getData("02", Convert.ToByte(1), "");

        //    if (downloaded_data.ToString() == "/")
        //    {
        //        {
        //            //MessageBox.Show("Reader is Empty.", "Reader Response", MessageBoxButtons.OK, MessageBoxIcon.Error);//                    
        //            var_rfid = "error";
        //        }
        //    }
        //    else
        //    {//Data Found. Now Upgrade Dbase status.
        //        var_rfid=downloaded_data.ToString();
        //    }
        //}//Threadfunc end.        
        void existance_statusfun(string name, string status, int this_rfid)
        {
            switch (status)
            {
                case "IN":
                    {
                        Bollo_naa.Speak(name + ", has Log on.", SpeechLib.SpeechVoiceSpeakFlags.SVSFDefault);
                        break;
                    }
                case "OUT":
                    {
                        Bollo_naa.Speak(name + ", has Log out.", SpeechLib.SpeechVoiceSpeakFlags.SVSFDefault);                                            
                        break;
                    }
            }//Switch OFF.
        }
        /***********************************************************************************/

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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Recep));
            this.ReceptCtlPicture = new System.Windows.Forms.PictureBox();
            this.RecepTxtRfid = new System.Windows.Forms.TextBox();
            this.RecepLblName = new System.Windows.Forms.Label();
            this.RecepLblDept = new System.Windows.Forms.Label();
            this.RecepLblBps = new System.Windows.Forms.Label();
            this.RecepLblTempinsert = new System.Windows.Forms.Label();
            this.RecepLblPic = new System.Windows.Forms.Label();
            this.RecepLblTimestamp = new System.Windows.Forms.Label();
            this.RecepCtlDate = new System.Windows.Forms.DateTimePicker();
            this.RecepLblSystemdate = new System.Windows.Forms.Label();
            this.RecepBtnExit = new System.Windows.Forms.Button();
            this.ReceptBtnBack2main = new System.Windows.Forms.Button();
            this.RecepLblLogin = new System.Windows.Forms.Label();
            this.RecepLblJobid = new System.Windows.Forms.Label();
            this.RecepLblSlash = new System.Windows.Forms.Label();
            this.RecepBtnAtten = new System.Windows.Forms.Button();
            this.RecepBtnForcedout = new System.Windows.Forms.Button();
            this.RecepLblinout = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mainToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LblReader_status = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.Recption_daily_attend = new System.Windows.Forms.Label();
            this.LblName = new System.Windows.Forms.Label();
            this.LblBps = new System.Windows.Forms.Label();
            this.LblDepartment = new System.Windows.Forms.Label();
            this.LblTimestamp = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.ReceptCtlPicture)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ReceptCtlPicture
            // 
            this.ReceptCtlPicture.BackColor = System.Drawing.Color.LightGray;
            this.ReceptCtlPicture.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ReceptCtlPicture.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.ReceptCtlPicture.Location = new System.Drawing.Point(675, 124);
            this.ReceptCtlPicture.Name = "ReceptCtlPicture";
            this.ReceptCtlPicture.Size = new System.Drawing.Size(327, 412);
            this.ReceptCtlPicture.TabIndex = 0;
            this.ReceptCtlPicture.TabStop = false;
            // 
            // RecepTxtRfid
            // 
            this.RecepTxtRfid.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.RecepTxtRfid.Location = new System.Drawing.Point(723, 542);
            this.RecepTxtRfid.Name = "RecepTxtRfid";
            this.RecepTxtRfid.Size = new System.Drawing.Size(79, 20);
            this.RecepTxtRfid.TabIndex = 16;
            // 
            // RecepLblName
            // 
            this.RecepLblName.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblName.Location = new System.Drawing.Point(18, 148);
            this.RecepLblName.Name = "RecepLblName";
            this.RecepLblName.Size = new System.Drawing.Size(92, 27);
            this.RecepLblName.TabIndex = 1;
            this.RecepLblName.Text = "Name";
            // 
            // RecepLblDept
            // 
            this.RecepLblDept.AutoSize = true;
            this.RecepLblDept.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblDept.Location = new System.Drawing.Point(18, 378);
            this.RecepLblDept.Name = "RecepLblDept";
            this.RecepLblDept.Size = new System.Drawing.Size(136, 26);
            this.RecepLblDept.TabIndex = 5;
            this.RecepLblDept.Text = "Department";
            // 
            // RecepLblBps
            // 
            this.RecepLblBps.AutoSize = true;
            this.RecepLblBps.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblBps.Location = new System.Drawing.Point(18, 242);
            this.RecepLblBps.Name = "RecepLblBps";
            this.RecepLblBps.Size = new System.Drawing.Size(53, 26);
            this.RecepLblBps.TabIndex = 3;
            this.RecepLblBps.Text = "Bps";
            // 
            // RecepLblTempinsert
            // 
            this.RecepLblTempinsert.AutoSize = true;
            this.RecepLblTempinsert.ForeColor = System.Drawing.Color.Black;
            this.RecepLblTempinsert.Location = new System.Drawing.Point(677, 545);
            this.RecepLblTempinsert.Name = "RecepLblTempinsert";
            this.RecepLblTempinsert.Size = new System.Drawing.Size(40, 13);
            this.RecepLblTempinsert.TabIndex = 15;
            this.RecepLblTempinsert.Text = "Tag ID";
            // 
            // RecepLblPic
            // 
            this.RecepLblPic.AutoSize = true;
            this.RecepLblPic.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblPic.Location = new System.Drawing.Point(671, 101);
            this.RecepLblPic.Name = "RecepLblPic";
            this.RecepLblPic.Size = new System.Drawing.Size(65, 20);
            this.RecepLblPic.TabIndex = 14;
            this.RecepLblPic.Text = "Picture";
            // 
            // RecepLblTimestamp
            // 
            this.RecepLblTimestamp.AutoSize = true;
            this.RecepLblTimestamp.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblTimestamp.Location = new System.Drawing.Point(18, 499);
            this.RecepLblTimestamp.Name = "RecepLblTimestamp";
            this.RecepLblTimestamp.Size = new System.Drawing.Size(140, 26);
            this.RecepLblTimestamp.TabIndex = 7;
            this.RecepLblTimestamp.Text = "Time Stamp";
            // 
            // RecepCtlDate
            // 
            this.RecepCtlDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepCtlDate.Location = new System.Drawing.Point(188, 693);
            this.RecepCtlDate.Name = "RecepCtlDate";
            this.RecepCtlDate.Size = new System.Drawing.Size(366, 26);
            this.RecepCtlDate.TabIndex = 13;
            this.RecepCtlDate.ValueChanged += new System.EventHandler(this.RecepCtlDate_ValueChanged);
            // 
            // RecepLblSystemdate
            // 
            this.RecepLblSystemdate.AutoSize = true;
            this.RecepLblSystemdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblSystemdate.Location = new System.Drawing.Point(18, 693);
            this.RecepLblSystemdate.Name = "RecepLblSystemdate";
            this.RecepLblSystemdate.Size = new System.Drawing.Size(149, 26);
            this.RecepLblSystemdate.TabIndex = 12;
            this.RecepLblSystemdate.Text = "System Date";
            // 
            // RecepBtnExit
            // 
            this.RecepBtnExit.Location = new System.Drawing.Point(910, 695);
            this.RecepBtnExit.Name = "RecepBtnExit";
            this.RecepBtnExit.Size = new System.Drawing.Size(70, 23);
            this.RecepBtnExit.TabIndex = 0;
            this.RecepBtnExit.Text = "Exit";
            this.RecepBtnExit.UseVisualStyleBackColor = true;
            this.RecepBtnExit.Click += new System.EventHandler(this.RecepBtnExit_Click);
            // 
            // ReceptBtnBack2main
            // 
            this.ReceptBtnBack2main.Location = new System.Drawing.Point(909, 696);
            this.ReceptBtnBack2main.Name = "ReceptBtnBack2main";
            this.ReceptBtnBack2main.Size = new System.Drawing.Size(79, 23);
            this.ReceptBtnBack2main.TabIndex = 17;
            this.ReceptBtnBack2main.Text = "Main";
            this.ReceptBtnBack2main.UseVisualStyleBackColor = true;
            this.ReceptBtnBack2main.Click += new System.EventHandler(this.button1_Click);
            // 
            // RecepLblLogin
            // 
            this.RecepLblLogin.AutoSize = true;
            this.RecepLblLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblLogin.Location = new System.Drawing.Point(745, 7);
            this.RecepLblLogin.Name = "RecepLblLogin";
            this.RecepLblLogin.Size = new System.Drawing.Size(48, 17);
            this.RecepLblLogin.TabIndex = 9;
            this.RecepLblLogin.Text = "Login";
            // 
            // RecepLblJobid
            // 
            this.RecepLblJobid.AutoSize = true;
            this.RecepLblJobid.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblJobid.Location = new System.Drawing.Point(881, 9);
            this.RecepLblJobid.Name = "RecepLblJobid";
            this.RecepLblJobid.Size = new System.Drawing.Size(52, 17);
            this.RecepLblJobid.TabIndex = 11;
            this.RecepLblJobid.Text = "Job id";
            // 
            // RecepLblSlash
            // 
            this.RecepLblSlash.AutoSize = true;
            this.RecepLblSlash.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblSlash.Location = new System.Drawing.Point(871, 9);
            this.RecepLblSlash.Name = "RecepLblSlash";
            this.RecepLblSlash.Size = new System.Drawing.Size(13, 17);
            this.RecepLblSlash.TabIndex = 10;
            this.RecepLblSlash.Text = "/";
            // 
            // RecepBtnAtten
            // 
            this.RecepBtnAtten.Location = new System.Drawing.Point(723, 568);
            this.RecepBtnAtten.Name = "RecepBtnAtten";
            this.RecepBtnAtten.Size = new System.Drawing.Size(79, 23);
            this.RecepBtnAtten.TabIndex = 24;
            this.RecepBtnAtten.Text = "Attendance";
            this.RecepBtnAtten.UseVisualStyleBackColor = true;
            this.RecepBtnAtten.Click += new System.EventHandler(this.RecepBtnAtten_Click);
            // 
            // RecepBtnForcedout
            // 
            this.RecepBtnForcedout.Location = new System.Drawing.Point(826, 64);
            this.RecepBtnForcedout.Name = "RecepBtnForcedout";
            this.RecepBtnForcedout.Size = new System.Drawing.Size(132, 23);
            this.RecepBtnForcedout.TabIndex = 26;
            this.RecepBtnForcedout.Text = "Forced Out";
            this.RecepBtnForcedout.UseVisualStyleBackColor = true;
            this.RecepBtnForcedout.Visible = false;
            this.RecepBtnForcedout.Click += new System.EventHandler(this.RecepBtnForcedout_Click);
            // 
            // RecepLblinout
            // 
            this.RecepLblinout.BackColor = System.Drawing.Color.LavenderBlush;
            this.RecepLblinout.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecepLblinout.ForeColor = System.Drawing.Color.Crimson;
            this.RecepLblinout.Location = new System.Drawing.Point(350, 273);
            this.RecepLblinout.Name = "RecepLblinout";
            this.RecepLblinout.Size = new System.Drawing.Size(185, 52);
            this.RecepLblinout.TabIndex = 27;
            this.RecepLblinout.Text = "STATUS";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Location = new System.Drawing.Point(601, 254);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(68, 71);
            this.pictureBox1.TabIndex = 28;
            this.pictureBox1.TabStop = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1030, 24);
            this.menuStrip1.TabIndex = 29;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem,
            this.mainToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(96, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // mainToolStripMenuItem
            // 
            this.mainToolStripMenuItem.Name = "mainToolStripMenuItem";
            this.mainToolStripMenuItem.Size = new System.Drawing.Size(96, 22);
            this.mainToolStripMenuItem.Text = "Main";
            this.mainToolStripMenuItem.Click += new System.EventHandler(this.mainToolStripMenuItem_Click);
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
            // LblReader_status
            // 
            this.LblReader_status.AutoSize = true;
            this.LblReader_status.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblReader_status.ForeColor = System.Drawing.Color.Red;
            this.LblReader_status.Location = new System.Drawing.Point(352, 81);
            this.LblReader_status.Name = "LblReader_status";
            this.LblReader_status.Size = new System.Drawing.Size(306, 24);
            this.LblReader_status.TabIndex = 60;
            this.LblReader_status.Text = "Reader Initializing in Progress...";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 2000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Recption_daily_attend
            // 
            this.Recption_daily_attend.BackColor = System.Drawing.Color.WhiteSmoke;
            this.Recption_daily_attend.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Recption_daily_attend.Location = new System.Drawing.Point(0, 24);
            this.Recption_daily_attend.Name = "Recption_daily_attend";
            this.Recption_daily_attend.Size = new System.Drawing.Size(1030, 24);
            this.Recption_daily_attend.TabIndex = 61;
            this.Recption_daily_attend.Text = "                                                           MIA DAILY ATTENDANCE";
            // 
            // LblName
            // 
            this.LblName.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.LblName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LblName.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblName.Location = new System.Drawing.Point(188, 144);
            this.LblName.Name = "LblName";
            this.LblName.Size = new System.Drawing.Size(276, 35);
            this.LblName.TabIndex = 62;
            // 
            // LblBps
            // 
            this.LblBps.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.LblBps.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LblBps.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblBps.Location = new System.Drawing.Point(188, 236);
            this.LblBps.Name = "LblBps";
            this.LblBps.Size = new System.Drawing.Size(75, 32);
            this.LblBps.TabIndex = 63;
            // 
            // LblDepartment
            // 
            this.LblDepartment.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.LblDepartment.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LblDepartment.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblDepartment.Location = new System.Drawing.Point(188, 374);
            this.LblDepartment.Name = "LblDepartment";
            this.LblDepartment.Size = new System.Drawing.Size(276, 35);
            this.LblDepartment.TabIndex = 64;
            // 
            // LblTimestamp
            // 
            this.LblTimestamp.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.LblTimestamp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LblTimestamp.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblTimestamp.Location = new System.Drawing.Point(188, 495);
            this.LblTimestamp.Name = "LblTimestamp";
            this.LblTimestamp.Size = new System.Drawing.Size(276, 35);
            this.LblTimestamp.TabIndex = 65;
            // 
            // Recep
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(1030, 748);
            this.Controls.Add(this.LblTimestamp);
            this.Controls.Add(this.LblDepartment);
            this.Controls.Add(this.LblBps);
            this.Controls.Add(this.LblName);
            this.Controls.Add(this.Recption_daily_attend);
            this.Controls.Add(this.LblReader_status);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.RecepLblinout);
            this.Controls.Add(this.RecepBtnForcedout);
            this.Controls.Add(this.RecepBtnAtten);
            this.Controls.Add(this.RecepLblSlash);
            this.Controls.Add(this.RecepLblJobid);
            this.Controls.Add(this.RecepLblLogin);
            this.Controls.Add(this.RecepLblSystemdate);
            this.Controls.Add(this.RecepCtlDate);
            this.Controls.Add(this.RecepLblTimestamp);
            this.Controls.Add(this.RecepLblPic);
            this.Controls.Add(this.RecepLblTempinsert);
            this.Controls.Add(this.RecepLblBps);
            this.Controls.Add(this.RecepLblDept);
            this.Controls.Add(this.RecepLblName);
            this.Controls.Add(this.RecepTxtRfid);
            this.Controls.Add(this.ReceptCtlPicture);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.ReceptBtnBack2main);
            this.Controls.Add(this.RecepBtnExit);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "Recep";
            this.Text = "M.I.A Daily Attendance";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.Recep_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Recep_FormClosing);
            this.Load += new System.EventHandler(this.Recep_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ReceptCtlPicture)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.PictureBox ReceptCtlPicture;
        private System.Windows.Forms.TextBox RecepTxtRfid;
        private System.Windows.Forms.Label RecepLblName;
        private System.Windows.Forms.Label RecepLblDept;
        private System.Windows.Forms.Label RecepLblBps;
        private System.Windows.Forms.Label RecepLblTempinsert;
        private System.Windows.Forms.Label RecepLblPic;
        private System.Windows.Forms.Label RecepLblTimestamp;
        private System.Windows.Forms.DateTimePicker RecepCtlDate;
        private System.Windows.Forms.Label RecepLblSystemdate;
        private System.Windows.Forms.Button RecepBtnExit;
        private System.Windows.Forms.Button ReceptBtnBack2main;
        private System.Windows.Forms.Label RecepLblLogin;
        private System.Windows.Forms.Label RecepLblJobid;
        private System.Windows.Forms.Label RecepLblSlash;
        private Button RecepBtnAtten;
        private Button RecepBtnForcedout;
        private Label RecepLblinout;
        private PictureBox pictureBox1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripMenuItem aboutToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private Label LblReader_status;
        private System.Windows.Forms.Timer timer1;
        private Label Recption_daily_attend;
        private Label LblName;
        private Label LblBps;
        private Label LblDepartment;
        private Label LblTimestamp;
        private ToolStripMenuItem mainToolStripMenuItem;
    }
}