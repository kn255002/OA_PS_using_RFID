using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;


//using System.Data.OracleClient;
namespace Deo_Module
{
    partial class Scanner
    {
        // Required designer variable.
        //
        //
        //Data members only.
        private System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();

        private string strImageName1, strImageName2, strTemp1, strTemp2;
        private int form_id1, form_id2;

        static string strOrcl9i = "Data Source=(DESCRIPTION="
             + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + Safeinfo.machinename + ")(PORT=1521)))"
             + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" + Safeinfo.servicename + ")));"
             + "User Id=Pk_ocp ;Password=me;";
        OracleConnection con = new OracleConnection(strOrcl9i);

        private Point crop_down;
        private int crop_width, crop_height;
       
        /*****************************************************************************************/
        /**********************************Now my Functions.**************************************/
        void Hide_Control()
        {
            TxtFormimage.Text ="";
            Picbox2.Image = null;
            TxtEmpimage.Text ="";
            Picbox1.Image = null;
            BtnSave.Enabled = false;
            TxtPrivewformid.Text = null;
            TxtScannedby.Hide();
            LblScannedby.Hide();
            form_id1 = form_id2 = 0;
            strImageName1 = strImageName2 = strTemp1 = strTemp2 = null;

            TxtCropXaxis.Text = TxtCropYaxis.Text = "";
            BtnCrop.Enabled = false;
            TxtHeight.Text = "";
            TxtWidth.Text = "";
            con.Close();

        }
        int conv_strtoint(string takeString, bool browse_check)
        {
            takeString = takeString.Substring(takeString.Length - 10);//Take last 10 Char of a file.        
            takeString = takeString.Substring(1, 5);

            if (takeString.StartsWith("P") && browse_check == true)   //'True' Represent It is 1st Browser Button.
            {
                //Conv from string1 to Integer1.
                TxtEmpimage.Text = takeString;
                strTemp1 = takeString.Substring(1, 4);               //Take only Integer from 10 Char.  
                //form_id1 = int.Parse(strTemp1);                      //Convert 04 Character into Integer.  
                return 1;                                            //1 Represent, Result is Only for 1st Browser Button.
            }

            else if (takeString.StartsWith("F") && browse_check == false)//'false' Represent It is 2nd Browser Button.
            {
                //Conv from string1 to Integer1.
                TxtFormimage.Text = takeString;
                strTemp2 = takeString.Substring(1, 4);              //Take only Integer from 10 Char.  
                form_id2 = int.Parse(strTemp2);                     //Convert 04 Character into Integer. 
                //if (form_id1 == form_id2) { BtnSave.Enabled = true; }
                return 2;                                            //2 Represent, Result is Only for 2nd Browser Button.
            }
            else
            {
                MessageBox.Show("Please Choose Valid Image Name. Start with 'F'","Image Selection Problem",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return 0;
            }
        }
        void delete_imagefun(string dele_path1, string dele_path2)
        {
            try
            {
                //MessageBox.Show("I'm going to delete Physical File.");
                Picbox1.Image.Dispose();
                Picbox1.Image = null;
                string myfile = dele_path1;
                File.Delete(myfile);

                Picbox2.Image.Dispose();
                Picbox2.Image = null;
                myfile = dele_path2;
                File.Delete(myfile);
                //--MessageBox.Show("Images has been deleted.","Images Deleted",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
            catch (IOException) { MessageBox.Show(" Image/File Could not deleted.","File I/O Error",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
        }
        int Save_imageintodbfun()    //Save both Images into database.
        {
            try
            {
                con.Open();
                //Change the default cursor to 'WaitCursor'(an HourGlass)        
                this.Cursor = Cursors.WaitCursor;

                //If curAdId is null then insert record
                //if (curAdID == "") 

                //To fill Dataset and update datasource.
                OracleDataAdapter pAdapter = new OracleDataAdapter("Select * from Employee", con);
                DataSet pDataSet = new DataSet("Employee");
                pAdapter.FillSchema(pDataSet, SchemaType.Source, "Employee");

                //For automatically generating commands to make changes to database through DataSet
                OracleCommandBuilder pCmdBldr = new OracleCommandBuilder(pAdapter);

                pAdapter.Fill(pDataSet, "Employee");

                //AddWithKey sets the Primary Key information to complete the 
                //schema information
                pAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;

                //Create a new row in the DataTable contained in the DataSet
                DataRow pRow = pDataSet.Tables["Employee"].NewRow();
                //Scanner will save only 05 Column's data.
                pRow["Form_id"] = form_id2;                 //Setting Column Value.
                pRow["FORMSTATUS_ID"] = 0;                  //Setting Column Value.
                pRow["SCANNED_BY"] = Safeinfo.loginid;      //Setting Column Value.
                pRow["PASSWORD"] = form_id2;                //Setting Column Value.
                pRow["ACCOUNT_STATUS"] = 0;                 //Setting Column Value.


                //ADD Image is added.
                if (strImageName1 == null)
                {
                    strImageName1 = "c:/picbox1.jpeg";
                    Picbox1.Image.Save(strImageName1.ToUpper(), System.Drawing.Imaging.ImageFormat.Jpeg);
                    //Read Access to the file chosed by 'Browse' button.
                    FileStream fs = new FileStream(strImageName1.ToUpper(), FileMode.Open, FileAccess.Read);   //ok

                    //Create a byte array of file stream length
                    byte[] ImageData = new byte[fs.Length];                                           //ok 

                    //Read block of bytes from stream into byte array
                    fs.Read(ImageData, 0, System.Convert.ToInt32(fs.Length));                         //ok
                    fs.Close();

                    //Assigning the byte array containing image data.
                    pRow["Emp_image"] = ImageData;

                }
                else if (strImageName1 != "")
                {
                    //Read Access to the file chosed by 'Browse' button.
                    FileStream fs = new FileStream(@strImageName1, FileMode.Open, FileAccess.Read);   //ok

                    //Create a byte array of file stream length
                    byte[] ImageData = new byte[fs.Length];                                           //ok 

                    //Read block of bytes from stream into byte array
                    fs.Read(ImageData, 0, System.Convert.ToInt32(fs.Length));                         //ok
                    fs.Close();

                    //Assigning the byte array containing image data.
                    pRow["Emp_image"] = ImageData;
                }//Employee own pic into database.
                /************************************************************/
                /************************************************************/
                if (strImageName2 != "")
                {
                    //Read Access to the file chosed by 'Browse' button.
                    FileStream fs = new FileStream(@strImageName2, FileMode.Open, FileAccess.Read);   //ok

                    //Create a byte array of file stream length
                    byte[] ImageData = new byte[fs.Length];                                           //ok 

                    //Read block of bytes from stream into byte array
                    fs.Read(ImageData, 0, System.Convert.ToInt32(fs.Length));                         //ok
                    fs.Close();

                    //Assigning the byte array containing image data.
                    pRow["Form_image"] = ImageData;  //Column# 03.
                }//Employee Form image into database.
                else
                {
                    MessageBox.Show("Without Images a Row is being added into Database.","Addition Caution",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                }

                //Adding the 'printmediaRow' to the DataSet
                pDataSet.Tables["Employee"].Rows.Add(pRow);
                pAdapter.Update(pDataSet, "Employee");
                MessageBox.Show("One record successfully added.", "Record Successful Added", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //--MessageBox.Show("Time Taken in Seconds:" + watch.ElapsedMilliseconds / 1000 + "\nTime Taken in Milliseconds:" + watch.ElapsedMilliseconds, "Time Consumed");
                this.Cursor = Cursors.Default;
                return 1;

            }//OFF try block.
                catch (DataException oex)             
                {
                    MessageBox.Show(oex.Message.ToString(),"Oracle Database Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return 0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    return 0;
                }
                finally
                {
                    //MessageBox.Show("I'm going to close db connection.");
                    con.Close();
                }
            }//OFF Save_imageintodbfun.
        void draw_fun(Point crop_up)
        {
            Picbox2.Refresh();
            crop_width = crop_up.X - crop_down.Y;
            crop_height = crop_up.Y - crop_down.Y;            
            Size mysize=new Size(crop_width,crop_height);
            Rectangle myrect = new Rectangle(crop_down, mysize);
//            Graphics dc = Picbox2.CreateGraphics();
            Pen RedPen = new Pen(Color.Red, 3);
            Picbox2.CreateGraphics().DrawRectangle(RedPen, myrect);
            GC.Collect();   //Garabase Collection. Now start Explictly Garbase collection.            
        }

       /*****************************************************************************************/
       /*****************************************************************************************/
       //
       //
       //






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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Scanner));
            this.BtnBrowseform = new System.Windows.Forms.Button();
            this.TxtFormimage = new System.Windows.Forms.TextBox();
            this.LblFormno = new System.Windows.Forms.Label();
            this.BtnSave = new System.Windows.Forms.Button();
            this.Picbox2 = new System.Windows.Forms.PictureBox();
            this.BtnLogout = new System.Windows.Forms.Button();
            this.TxtLoginname = new System.Windows.Forms.TextBox();
            this.TxtLoginas = new System.Windows.Forms.TextBox();
            this.LblLoginname = new System.Windows.Forms.Label();
            this.LblLoginas = new System.Windows.Forms.Label();
            this.BtnClear = new System.Windows.Forms.Button();
            this.Picbox1 = new System.Windows.Forms.PictureBox();
            this.BtnBrowsemp = new System.Windows.Forms.Button();
            this.BtnDelete = new System.Windows.Forms.Button();
            this.LblEmpPicture = new System.Windows.Forms.Label();
            this.TxtEmpimage = new System.Windows.Forms.TextBox();
            this.BtnPrivew = new System.Windows.Forms.Button();
            this.TxtPrivewformid = new System.Windows.Forms.TextBox();
            this.LblPrivewformno = new System.Windows.Forms.Label();
            this.TxtScannedby = new System.Windows.Forms.TextBox();
            this.LblScannedby = new System.Windows.Forms.Label();
            this.BtnMain = new System.Windows.Forms.Button();
            this.LblCropXaxis = new System.Windows.Forms.Label();
            this.TxtCropXaxis = new System.Windows.Forms.TextBox();
            this.BtnCrop = new System.Windows.Forms.Button();
            this.LblCropYaxis = new System.Windows.Forms.Label();
            this.TxtCropYaxis = new System.Windows.Forms.TextBox();
            this.LblHeight = new System.Windows.Forms.Label();
            this.TxtHeight = new System.Windows.Forms.TextBox();
            this.LblWidth = new System.Windows.Forms.Label();
            this.TxtWidth = new System.Windows.Forms.TextBox();
            this.CkbChoosepic = new System.Windows.Forms.CheckBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.backToMainToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.privewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.indexToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.searchToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.Picbox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Picbox1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnBrowseform
            // 
            this.BtnBrowseform.Location = new System.Drawing.Point(689, 676);
            this.BtnBrowseform.Name = "BtnBrowseform";
            this.BtnBrowseform.Size = new System.Drawing.Size(67, 21);
            this.BtnBrowseform.TabIndex = 0;
            this.BtnBrowseform.Text = "Browse";
            this.BtnBrowseform.UseVisualStyleBackColor = true;
            this.BtnBrowseform.Click += new System.EventHandler(this.BtnBrowseform_Click_1);
            // 
            // TxtFormimage
            // 
            this.TxtFormimage.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtFormimage.Location = new System.Drawing.Point(517, 676);
            this.TxtFormimage.Name = "TxtFormimage";
            this.TxtFormimage.ReadOnly = true;
            this.TxtFormimage.Size = new System.Drawing.Size(166, 20);
            this.TxtFormimage.TabIndex = 1;
            // 
            // LblFormno
            // 
            this.LblFormno.AutoSize = true;
            this.LblFormno.Location = new System.Drawing.Point(464, 680);
            this.LblFormno.Name = "LblFormno";
            this.LblFormno.Size = new System.Drawing.Size(47, 13);
            this.LblFormno.TabIndex = 2;
            this.LblFormno.Text = "Form No";
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(783, 676);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(67, 21);
            this.BtnSave.TabIndex = 8;
            this.BtnSave.Text = "Save";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BSave_Click);
            // 
            // Picbox2
            // 
            this.Picbox2.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.Picbox2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("Picbox2.BackgroundImage")));
            this.Picbox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.Picbox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.Picbox2.InitialImage = null;
            this.Picbox2.Location = new System.Drawing.Point(4, 27);
            this.Picbox2.Name = "Picbox2";
            this.Picbox2.Size = new System.Drawing.Size(846, 643);
            this.Picbox2.TabIndex = 5;
            this.Picbox2.TabStop = false;
            this.Picbox2.WaitOnLoad = true;
            this.Picbox2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Picbox2_MouseDown);
            this.Picbox2.MouseMove += new System.Windows.Forms.MouseEventHandler(this.Picbox2_MouseMove);
            this.Picbox2.MouseUp += new System.Windows.Forms.MouseEventHandler(this.Picbox2_MouseUp);
            // 
            // BtnLogout
            // 
            this.BtnLogout.Location = new System.Drawing.Point(944, 649);
            this.BtnLogout.Name = "BtnLogout";
            this.BtnLogout.Size = new System.Drawing.Size(67, 21);
            this.BtnLogout.TabIndex = 11;
            this.BtnLogout.Text = "Log out";
            this.BtnLogout.UseVisualStyleBackColor = true;
            this.BtnLogout.Click += new System.EventHandler(this.BExit_Click);
            // 
            // TxtLoginname
            // 
            this.TxtLoginname.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtLoginname.ForeColor = System.Drawing.Color.Black;
            this.TxtLoginname.Location = new System.Drawing.Point(922, 357);
            this.TxtLoginname.Name = "TxtLoginname";
            this.TxtLoginname.ReadOnly = true;
            this.TxtLoginname.Size = new System.Drawing.Size(95, 20);
            this.TxtLoginname.TabIndex = 18;
            // 
            // TxtLoginas
            // 
            this.TxtLoginas.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtLoginas.ForeColor = System.Drawing.Color.Black;
            this.TxtLoginas.Location = new System.Drawing.Point(922, 380);
            this.TxtLoginas.Name = "TxtLoginas";
            this.TxtLoginas.ReadOnly = true;
            this.TxtLoginas.Size = new System.Drawing.Size(95, 20);
            this.TxtLoginas.TabIndex = 20;
            // 
            // LblLoginname
            // 
            this.LblLoginname.AutoSize = true;
            this.LblLoginname.ForeColor = System.Drawing.SystemColors.ControlText;
            this.LblLoginname.Location = new System.Drawing.Point(856, 361);
            this.LblLoginname.Name = "LblLoginname";
            this.LblLoginname.Size = new System.Drawing.Size(64, 13);
            this.LblLoginname.TabIndex = 17;
            this.LblLoginname.Text = "Login Name";
            // 
            // LblLoginas
            // 
            this.LblLoginas.AutoSize = true;
            this.LblLoginas.ForeColor = System.Drawing.SystemColors.ControlText;
            this.LblLoginas.Location = new System.Drawing.Point(856, 383);
            this.LblLoginas.Name = "LblLoginas";
            this.LblLoginas.Size = new System.Drawing.Size(48, 13);
            this.LblLoginas.TabIndex = 19;
            this.LblLoginas.Text = "Login As";
            // 
            // BtnClear
            // 
            this.BtnClear.Location = new System.Drawing.Point(878, 330);
            this.BtnClear.Name = "BtnClear";
            this.BtnClear.Size = new System.Drawing.Size(67, 21);
            this.BtnClear.TabIndex = 9;
            this.BtnClear.Text = "Clear";
            this.BtnClear.UseVisualStyleBackColor = true;
            this.BtnClear.Click += new System.EventHandler(this.BtnClear_Click);
            // 
            // Picbox1
            // 
            this.Picbox1.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.Picbox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.Picbox1.Location = new System.Drawing.Point(905, 27);
            this.Picbox1.Name = "Picbox1";
            this.Picbox1.Size = new System.Drawing.Size(112, 126);
            this.Picbox1.TabIndex = 11;
            this.Picbox1.TabStop = false;
            // 
            // BtnBrowsemp
            // 
            this.BtnBrowsemp.Location = new System.Drawing.Point(902, 209);
            this.BtnBrowsemp.Name = "BtnBrowsemp";
            this.BtnBrowsemp.Size = new System.Drawing.Size(67, 21);
            this.BtnBrowsemp.TabIndex = 6;
            this.BtnBrowsemp.Text = "Browse";
            this.BtnBrowsemp.UseVisualStyleBackColor = true;
            this.BtnBrowsemp.Visible = false;
            this.BtnBrowsemp.Click += new System.EventHandler(this.BtnBrowsemp_Click_1);
            // 
            // BtnDelete
            // 
            this.BtnDelete.Location = new System.Drawing.Point(950, 330);
            this.BtnDelete.Name = "BtnDelete";
            this.BtnDelete.Size = new System.Drawing.Size(67, 21);
            this.BtnDelete.TabIndex = 29;
            this.BtnDelete.Text = "Delete";
            this.BtnDelete.UseVisualStyleBackColor = true;
            this.BtnDelete.Click += new System.EventHandler(this.BtnDelete_Click);
            // 
            // LblEmpPicture
            // 
            this.LblEmpPicture.AutoSize = true;
            this.LblEmpPicture.Location = new System.Drawing.Point(899, 156);
            this.LblEmpPicture.Name = "LblEmpPicture";
            this.LblEmpPicture.Size = new System.Drawing.Size(89, 13);
            this.LblEmpPicture.TabIndex = 3;
            this.LblEmpPicture.Text = "Employee Picture";
            // 
            // TxtEmpimage
            // 
            this.TxtEmpimage.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtEmpimage.Location = new System.Drawing.Point(905, 170);
            this.TxtEmpimage.Name = "TxtEmpimage";
            this.TxtEmpimage.ReadOnly = true;
            this.TxtEmpimage.Size = new System.Drawing.Size(112, 20);
            this.TxtEmpimage.TabIndex = 4;
            // 
            // BtnPrivew
            // 
            this.BtnPrivew.Location = new System.Drawing.Point(950, 288);
            this.BtnPrivew.Name = "BtnPrivew";
            this.BtnPrivew.Size = new System.Drawing.Size(67, 21);
            this.BtnPrivew.TabIndex = 13;
            this.BtnPrivew.Text = "Privew";
            this.BtnPrivew.UseVisualStyleBackColor = true;
            this.BtnPrivew.Click += new System.EventHandler(this.BtnPrivew_Click);
            // 
            // TxtPrivewformid
            // 
            this.TxtPrivewformid.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtPrivewformid.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtPrivewformid.Location = new System.Drawing.Point(942, 236);
            this.TxtPrivewformid.Name = "TxtPrivewformid";
            this.TxtPrivewformid.Size = new System.Drawing.Size(75, 20);
            this.TxtPrivewformid.TabIndex = 12;
            this.TxtPrivewformid.Leave += new System.EventHandler(this.TxtPrivewformid_Leave);
            this.TxtPrivewformid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtPrivewformid_KeyPress);
            // 
            // LblPrivewformno
            // 
            this.LblPrivewformno.AutoSize = true;
            this.LblPrivewformno.Location = new System.Drawing.Point(869, 239);
            this.LblPrivewformno.Name = "LblPrivewformno";
            this.LblPrivewformno.Size = new System.Drawing.Size(47, 13);
            this.LblPrivewformno.TabIndex = 16;
            this.LblPrivewformno.Text = "Form No";
            // 
            // TxtScannedby
            // 
            this.TxtScannedby.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtScannedby.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtScannedby.Location = new System.Drawing.Point(942, 262);
            this.TxtScannedby.Name = "TxtScannedby";
            this.TxtScannedby.ReadOnly = true;
            this.TxtScannedby.Size = new System.Drawing.Size(75, 20);
            this.TxtScannedby.TabIndex = 15;
            // 
            // LblScannedby
            // 
            this.LblScannedby.AutoSize = true;
            this.LblScannedby.ForeColor = System.Drawing.SystemColors.ControlText;
            this.LblScannedby.Location = new System.Drawing.Point(869, 265);
            this.LblScannedby.Name = "LblScannedby";
            this.LblScannedby.Size = new System.Drawing.Size(64, 13);
            this.LblScannedby.TabIndex = 14;
            this.LblScannedby.Text = "Scanned by";
            // 
            // BtnMain
            // 
            this.BtnMain.Location = new System.Drawing.Point(937, 627);
            this.BtnMain.Name = "BtnMain";
            this.BtnMain.Size = new System.Drawing.Size(79, 21);
            this.BtnMain.TabIndex = 10;
            this.BtnMain.Text = "Main";
            this.BtnMain.UseVisualStyleBackColor = true;
            this.BtnMain.Click += new System.EventHandler(this.BtnBacktomain_Click);
            // 
            // LblCropXaxis
            // 
            this.LblCropXaxis.AutoSize = true;
            this.LblCropXaxis.Location = new System.Drawing.Point(858, 434);
            this.LblCropXaxis.Name = "LblCropXaxis";
            this.LblCropXaxis.Size = new System.Drawing.Size(17, 13);
            this.LblCropXaxis.TabIndex = 21;
            this.LblCropXaxis.Text = "X:";
            // 
            // TxtCropXaxis
            // 
            this.TxtCropXaxis.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtCropXaxis.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtCropXaxis.Location = new System.Drawing.Point(895, 430);
            this.TxtCropXaxis.Name = "TxtCropXaxis";
            this.TxtCropXaxis.Size = new System.Drawing.Size(38, 20);
            this.TxtCropXaxis.TabIndex = 22;
            this.TxtCropXaxis.Enter += new System.EventHandler(this.TxtCropXaxis_Enter);
            this.TxtCropXaxis.Leave += new System.EventHandler(this.TxtCropXaxis_Leave);
            this.TxtCropXaxis.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtCropXaxis_KeyPress);
            this.TxtCropXaxis.TextChanged += new System.EventHandler(this.TxtCropXaxis_TextChanged);
            // 
            // BtnCrop
            // 
            this.BtnCrop.Enabled = false;
            this.BtnCrop.Location = new System.Drawing.Point(950, 492);
            this.BtnCrop.Name = "BtnCrop";
            this.BtnCrop.Size = new System.Drawing.Size(67, 21);
            this.BtnCrop.TabIndex = 7;
            this.BtnCrop.Text = "Crop";
            this.BtnCrop.UseVisualStyleBackColor = true;
            this.BtnCrop.Click += new System.EventHandler(this.BtnCrop_Click);
            // 
            // LblCropYaxis
            // 
            this.LblCropYaxis.AutoSize = true;
            this.LblCropYaxis.Location = new System.Drawing.Point(941, 432);
            this.LblCropYaxis.Name = "LblCropYaxis";
            this.LblCropYaxis.Size = new System.Drawing.Size(17, 13);
            this.LblCropYaxis.TabIndex = 25;
            this.LblCropYaxis.Text = "Y:";
            // 
            // TxtCropYaxis
            // 
            this.TxtCropYaxis.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtCropYaxis.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtCropYaxis.Location = new System.Drawing.Point(979, 427);
            this.TxtCropYaxis.Name = "TxtCropYaxis";
            this.TxtCropYaxis.Size = new System.Drawing.Size(38, 20);
            this.TxtCropYaxis.TabIndex = 26;
            this.TxtCropYaxis.Enter += new System.EventHandler(this.TxtCropYaxis_Enter);
            this.TxtCropYaxis.Leave += new System.EventHandler(this.TxtCropYaxis_Leave);
            this.TxtCropYaxis.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtCropYaxis_KeyPress);
            this.TxtCropYaxis.TextChanged += new System.EventHandler(this.TxtCropYaxis_TextChanged);
            // 
            // LblHeight
            // 
            this.LblHeight.AutoSize = true;
            this.LblHeight.Location = new System.Drawing.Point(941, 458);
            this.LblHeight.Name = "LblHeight";
            this.LblHeight.Size = new System.Drawing.Size(38, 13);
            this.LblHeight.TabIndex = 27;
            this.LblHeight.Text = "Height";
            // 
            // TxtHeight
            // 
            this.TxtHeight.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtHeight.Enabled = false;
            this.TxtHeight.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtHeight.Location = new System.Drawing.Point(979, 453);
            this.TxtHeight.Name = "TxtHeight";
            this.TxtHeight.Size = new System.Drawing.Size(38, 20);
            this.TxtHeight.TabIndex = 28;
            // 
            // LblWidth
            // 
            this.LblWidth.AutoSize = true;
            this.LblWidth.Location = new System.Drawing.Point(858, 460);
            this.LblWidth.Name = "LblWidth";
            this.LblWidth.Size = new System.Drawing.Size(35, 13);
            this.LblWidth.TabIndex = 23;
            this.LblWidth.Text = "Width";
            // 
            // TxtWidth
            // 
            this.TxtWidth.BackColor = System.Drawing.Color.WhiteSmoke;
            this.TxtWidth.Enabled = false;
            this.TxtWidth.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtWidth.Location = new System.Drawing.Point(895, 456);
            this.TxtWidth.Name = "TxtWidth";
            this.TxtWidth.Size = new System.Drawing.Size(38, 20);
            this.TxtWidth.TabIndex = 24;
            // 
            // CkbChoosepic
            // 
            this.CkbChoosepic.AutoSize = true;
            this.CkbChoosepic.Location = new System.Drawing.Point(902, 192);
            this.CkbChoosepic.Name = "CkbChoosepic";
            this.CkbChoosepic.Size = new System.Drawing.Size(98, 17);
            this.CkbChoosepic.TabIndex = 5;
            this.CkbChoosepic.Text = "Choose Picture";
            this.CkbChoosepic.UseVisualStyleBackColor = true;
            this.CkbChoosepic.Visible = false;
            this.CkbChoosepic.CheckStateChanged += new System.EventHandler(this.CkbChoosepic_CheckStateChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1028, 24);
            this.menuStrip1.TabIndex = 30;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.toolStripSeparator2,
            this.backToMainToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "&File";
            this.fileToolStripMenuItem.MouseEnter += new System.EventHandler(this.fileToolStripMenuItem_MouseEnter);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.openToolStripMenuItem.Text = "Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveToolStripMenuItem.Image")));
            this.saveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.saveToolStripMenuItem.Text = "&Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(133, 6);
            // 
            // backToMainToolStripMenuItem
            // 
            this.backToMainToolStripMenuItem.Name = "backToMainToolStripMenuItem";
            this.backToMainToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.backToMainToolStripMenuItem.Text = "Main";
            this.backToMainToolStripMenuItem.Click += new System.EventHandler(this.backToMainToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.clearToolStripMenuItem,
            this.privewToolStripMenuItem,
            this.deleteToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.toolsToolStripMenuItem.Text = "&Tools";
            // 
            // clearToolStripMenuItem
            // 
            this.clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            this.clearToolStripMenuItem.Size = new System.Drawing.Size(106, 22);
            this.clearToolStripMenuItem.Text = "Clear";
            this.clearToolStripMenuItem.Click += new System.EventHandler(this.clearToolStripMenuItem_Click);
            // 
            // privewToolStripMenuItem
            // 
            this.privewToolStripMenuItem.Name = "privewToolStripMenuItem";
            this.privewToolStripMenuItem.Size = new System.Drawing.Size(106, 22);
            this.privewToolStripMenuItem.Text = "Privew";
            this.privewToolStripMenuItem.Click += new System.EventHandler(this.privewToolStripMenuItem_Click);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(106, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.contentsToolStripMenuItem,
            this.indexToolStripMenuItem,
            this.searchToolStripMenuItem,
            this.toolStripSeparator5,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // contentsToolStripMenuItem
            // 
            this.contentsToolStripMenuItem.Name = "contentsToolStripMenuItem";
            this.contentsToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.contentsToolStripMenuItem.Text = "&Contents";
            // 
            // indexToolStripMenuItem
            // 
            this.indexToolStripMenuItem.Name = "indexToolStripMenuItem";
            this.indexToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.indexToolStripMenuItem.Text = "&Index";
            // 
            // searchToolStripMenuItem
            // 
            this.searchToolStripMenuItem.Name = "searchToolStripMenuItem";
            this.searchToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.searchToolStripMenuItem.Text = "&Search";
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(115, 6);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.aboutToolStripMenuItem.Text = "&About...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Black;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Location = new System.Drawing.Point(0, 23);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1028, 2);
            this.pictureBox1.TabIndex = 54;
            this.pictureBox1.TabStop = false;
            // 
            // Scanner
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(1028, 746);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.CkbChoosepic);
            this.Controls.Add(this.LblHeight);
            this.Controls.Add(this.TxtHeight);
            this.Controls.Add(this.LblWidth);
            this.Controls.Add(this.TxtWidth);
            this.Controls.Add(this.LblCropYaxis);
            this.Controls.Add(this.TxtCropYaxis);
            this.Controls.Add(this.LblCropXaxis);
            this.Controls.Add(this.TxtCropXaxis);
            this.Controls.Add(this.BtnCrop);
            this.Controls.Add(this.BtnMain);
            this.Controls.Add(this.LblScannedby);
            this.Controls.Add(this.TxtScannedby);
            this.Controls.Add(this.LblPrivewformno);
            this.Controls.Add(this.TxtPrivewformid);
            this.Controls.Add(this.BtnPrivew);
            this.Controls.Add(this.TxtEmpimage);
            this.Controls.Add(this.LblEmpPicture);
            this.Controls.Add(this.BtnDelete);
            this.Controls.Add(this.BtnBrowsemp);
            this.Controls.Add(this.Picbox1);
            this.Controls.Add(this.BtnClear);
            this.Controls.Add(this.LblLoginname);
            this.Controls.Add(this.LblLoginas);
            this.Controls.Add(this.TxtLoginname);
            this.Controls.Add(this.BtnLogout);
            this.Controls.Add(this.TxtLoginas);
            this.Controls.Add(this.Picbox2);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.LblFormno);
            this.Controls.Add(this.TxtFormimage);
            this.Controls.Add(this.BtnBrowseform);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Scanner";
            this.Text = "M.I.A Scaner";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.Scanner_Activated);
            ((System.ComponentModel.ISupportInitialize)(this.Picbox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Picbox1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnBrowseform;
        private System.Windows.Forms.TextBox TxtFormimage;
        private System.Windows.Forms.Label LblFormno;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.PictureBox Picbox2;
        private System.Windows.Forms.Button BtnLogout;
        private System.Windows.Forms.TextBox TxtLoginname;
        private System.Windows.Forms.TextBox TxtLoginas;
        private System.Windows.Forms.Label LblLoginname;
        private System.Windows.Forms.Label LblLoginas;
        private System.Windows.Forms.Button BtnClear;
        private System.Windows.Forms.PictureBox Picbox1;
        private System.Windows.Forms.Button BtnBrowsemp;
        private System.Windows.Forms.Button BtnDelete;
        private System.Windows.Forms.Label LblEmpPicture;
        private System.Windows.Forms.TextBox TxtEmpimage;
        private System.Windows.Forms.Button BtnPrivew;
        private System.Windows.Forms.TextBox TxtPrivewformid;
        private System.Windows.Forms.Label LblPrivewformno;
        private System.Windows.Forms.TextBox TxtScannedby;
        private System.Windows.Forms.Label LblScannedby;
        private Button BtnMain;
        private Label LblCropXaxis;
        private TextBox TxtCropXaxis;
        private Button BtnCrop;
        private Label LblCropYaxis;
        private TextBox TxtCropYaxis;
        private Label LblHeight;
        private TextBox TxtHeight;
        private Label LblWidth;
        private TextBox TxtWidth;
        private CheckBox CkbChoosepic;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem saveToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripMenuItem toolsToolStripMenuItem;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripMenuItem contentsToolStripMenuItem;
        private ToolStripMenuItem indexToolStripMenuItem;
        private ToolStripMenuItem searchToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator5;
        private ToolStripMenuItem aboutToolStripMenuItem;
        private ToolStripMenuItem backToMainToolStripMenuItem;
        private ToolStripMenuItem privewToolStripMenuItem;
        private ToolStripMenuItem clearToolStripMenuItem;
        private ToolStripMenuItem openToolStripMenuItem;
        private ToolStripMenuItem deleteToolStripMenuItem;
        private PictureBox pictureBox1;
    }
}