using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;

namespace Deo_Module
{
    partial class Dataentry
    {

        
        //--------------------------------------------For DataMembers-------------------------------------------
        //Data members only.
        private Stopwatch watch = new System.Diagnostics.Stopwatch();
        private Update_Database gl_dbaseobj = new Update_Database(); //This object will be use for Accessing Update_Database class.
        static int form_id;         //When a Formid updated then save here. For next time accessing such particular form.
        //bool record_exist = false;
        
        //---------------------------For Drawing.---------------
        // Actually here we works on How to draw Rectangel according to control on image.
        // 
        private Pen mypen = new Pen(Color.Red, 4);
        Point myctllocation = new Point();
        Size myctlsize = new Size();
        //--------------------------------------------For Encryption-------------------------------------------
        private Encryption pk_ocpencrypt = new Encryption();

        /****************************************Now Method Implementation.*************************************/
        

        void set_clearfields()
        {
            form_id = 0;
            TxtFormid.Text = "";
            TxtName.Text = "";
            TxtFathername.Text = "";
            TxtCnic.Text = "";

            TxtDob.Text = "";
            CmbDay.Text = "";
            CmbMon.Text = "";
            CmbYear.Text = "";

            TxtGender.Text = "";
            CmbGender.Text = "";

            TxtReligion.Text = "";
            CmbReligion.Text = "";

            TxtDomicile.Text = "";
            CmbDomicile.Text = "";

            Txtmail.Text = "";
            TxtPhoneno.Text = "";
            TxtAccno.Text = "";
            TxtBankbranch.Text = "";
            TxtRftagno.Text = "";

            TxtEmployeetype.Text = "";
            CmbEmployeetype.Text = "";
            TxtTempaddress.Text = "";
            TxtPermaddress.Text = "";

            TxtDegreename.Text = "";
            TxtGrade.Text = "";
            CmbPassyear.Text = "";
            TxtMajor.Text="";
            TxtInstitute.Text = "";


            
            PictureBox2.Image = null;
            PictureBox1.Image = null;
            BtnSave.Enabled = false;
            BtnSkip.Enabled = false;           
        }
        void set_formfun(int choose_formid, int set_formid)
        {
            try
            {
                string query = "Select Form_id,Formstatus_id from Employee where  formstatus_id=" + choose_formid;
                gl_dbaseobj.Dml_UpdateAdapterfun(query);
                if (gl_dbaseobj.cls_dataset.Tables[0].Rows.Count >= 1)//Its mean if you found rows. KEEP in mind IT CAN BE MORETHAN ONE.
                {
                    if (Safeinfo.status=="DATA ENTRY OPERATOR" ||Safeinfo.purpose_deo_hr_receptionst == 0)//IF you are NOT HR or If you are DEO then Run this.
                    {
                        disable_ctlfun(false);//Enable Controls bcz Data is AVAILABLE FOR 'DEO' to ENTRY PROCESS. HA HAHHAAHAHA YES!!!! WORKING FINE.
                    }

                    BtnRefersh.Enabled = false;     //its mean you've found a record, So immediate disable Refresh button.
                    BtnSave.Enabled = true;
                    BtnSkip.Enabled = true;
                    TxtRftagno.Enabled = true;
                    LblRftagno.Enabled = true;                    

                    //HERE I TRY TO MAKE IT QUEUE. BY JUST CHOOSING FIRST AVAILABLE RECORD IN THE TABLE[0].ROWS[0].
                    //
                    form_id = int.Parse(gl_dbaseobj.cls_dataset.Tables[0].Rows[0][0].ToString()); //Saving form_id into a Global Variable.
                    //NOW IAM GOING TO LCOK THIS ROW/FORMID.
                    string str_lock_form="Select * from Employee Where form_id="+form_id+" for Update";
                    gl_dbaseobj.Show_Datafun(str_lock_form, 0);//Run query for LOCK this form, and keep OPEN connection.
                    gl_dbaseobj.Show_Datafun(str_lock_form, 1);//DOn't run query. Just close connection.
                   
                    switch (choose_formid)
                    {
                        case 0: //Its mean ur DEO OR Administrator with safeinfo.purpose_deo_hr=0
                            {
                                gl_dbaseobj.cls_dataset.Tables[0].Rows[0][1] = set_formid;    //Here i set this 01st row at particular status.
                                gl_dbaseobj.cls_odadptr.Update(gl_dbaseobj.cls_dataset);        //Here i update Database.                                
                                //NOW I RELEASE ALL LOCKS.
                                break;
                            }
                        case 2://Its mean ur HR  OR Administrator with safeinfo.purpose_deo_hr=1
                            {
                                gl_dbaseobj.cls_dataset.Tables[0].Rows[0][1] = set_formid;    //Here i set this 01st row at particular status.
                                gl_dbaseobj.cls_odadptr.Update(gl_dbaseobj.cls_dataset);        //Here i update Database.                                
                                break;
                            }
                    }//End switch.
                }//If block.                              
                else
                {
                    MessageBox.Show("Please Re-try after few seconds.\nData Queue is empty.", "Empty Queue ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    watch.Reset();
                    watch.Start();
                    form_id = 0;    //bcz form not found.
                    TxtRftagno.Enabled = false;
                    BtnSave.Enabled = false;
                    BtnSkip.Enabled = false;
                    BtnRefersh.Enabled = true;     //its mean you do not have found a record, So immediate Enable Refresh button for feature time.
                    disable_ctlfun(true);//Disable controls because Data is not available for Entry process.
                }//ene else.                
            }
            catch (OracleException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
            }
        }        
        void fetch_datafun(string job_status, bool want2_skip)
        {
            watch.Reset();
            watch.Start();
            watch.Stop();
            if (want2_skip == false)            //Just to bring+set A record,as here check_skip=false.
            { //I DO NOT WANT TO SKIP FORM.

                if (job_status.ToUpper() == "HR" || Safeinfo.purpose_deo_hr_receptionst== 1)//Now 1st verify status and then set_formfunc();
                {//Its mean Form is not going to be skip, Rather form is goin to be picked by HR.
                    BtnTag_addition.Show();     //Bcz its HR.And its ALLOWED to go back on Tag load form/page.
                    set_formfun(2, 3);         //Choose forms where formstatus_id=2 and set a particular form to formstatus_id=3                
                }                              //2=ready to HR, 3= picked by HR.
                else if (job_status.ToUpper() == "DATA ENTRY OPERATOR" || Safeinfo.purpose_deo_hr_receptionst == 0)
                {//Its mean Form is not going to be skip, Rather form is goin to be picked by deo.
                    BtnTag_addition.Hide();     //Bcz its not HR.And its not allowd to go back on Tag load form/page.
                    set_formfun(0, 1);         //Choose forms where formstatus_id=0 and set a particular form to formstatus_id=1
                }                              //0=ready to DEO, 1= picked by DEO.
            }
            else//Just to Skip record, as here check_skip=true.
            {//YES I WANT TO SKIP FORM.
                if (job_status.ToUpper() == "HR"|| Safeinfo.purpose_deo_hr_receptionst== 1)
                {
                    set4_deo_scanfun(0);         //Set this form=0, Ready for DEO.
                }
                else if (job_status.ToUpper() == "DATA ENTRY OPERATOR" || Safeinfo.purpose_deo_hr_receptionst == 0)
                {
                    set4_deo_scanfun(-1);         //Set this form=-1, Ready for Scanner.
                }
            }
        }
        void load_recfun()
        {
                //string choosequery = "Select *  from Employee where form_id=" + form_id;  //Choose AGAIN THIS PARTICULAR FORM.
                string choosequery = "Select Form_id,Ename,Father_name,cnic,To_Char(Date_birth,'DD-MON-RRRR'),Gender,Religion,Domicile,Email,Phone_no,Bank_acc,Bank_Branch,Formstatus_id,Scanned_by,Deo_by,Hr_by,Admin_by,Password,Account_Status,Rfid_id,Bps,Job_id,Mgr_id,Dept_id,Hire_date,Joining_date,Emp_type,Temp_Address,Perm_Address,Emp_image,Form_image from Employee where form_id=" + form_id;  //Choose AGAIN THIS PARTICULAR FORM.                
                gl_dbaseobj.Show_Datafun(choosequery, 0);
                bool data_found = gl_dbaseobj.cls_dr.Read();
                if (data_found)
                {
                    if (Safeinfo.status == "DATA ENTRY OPERATOR" || Safeinfo.purpose_deo_hr_receptionst== 0)
                    {
                        TxtFormid.Text = gl_dbaseobj.cls_dr.GetDecimal(0).ToString(); //Display form_id into Textbox which id captured.                        
                        {   //Now Displaying Picture.Emp_Image.
                                                        
                            //MessageBox.Show("Employee Image Loading, Now");
                            OracleLob blob = gl_dbaseobj.cls_dr.GetOracleLob(29);
                            Byte[] BLOBData = new Byte[blob.Length];

                            //Read blob data into byte array
                            int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));

                            //Get the primitive byte data into in-memory data stream
                            System.IO.MemoryStream stmBLOBData = new System.IO.MemoryStream(BLOBData);

                            //LOADING INTO PICBOX1
                            PictureBox1.Image = Image.FromStream(stmBLOBData);
                            //MessageBox.Show("Now i'm reseting it.");
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                        }
                        {   //Now Displaying Picture.form_image
                                                        
                            //MessageBox.Show("Form's image Loading, Now");
                            OracleLob blob = gl_dbaseobj.cls_dr.GetOracleLob(30);   //Column # 30.      form_image
                            Byte[] BLOBData = new Byte[blob.Length];

                            //Read blob data into byte array
                            int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));

                            //Get the primitive byte data into in-memory data stream
                            System.IO.MemoryStream stmBLOBData = new System.IO.MemoryStream(BLOBData);

                            //LOADING INTO PICBOX1
                            PictureBox2.Image = Image.FromStream(stmBLOBData);
                            //MessageBox.Show("Now i'm reseting it.");
                            PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                        }
                        //Now enter normal column values.
                    }//Now OFF, Deo Data displayed.

                    else if (Safeinfo.status == "HR" || Safeinfo.purpose_deo_hr_receptionst == 1)
                    {//Here you show data for HR.

                        TxtFormid.Text = gl_dbaseobj.cls_dr.GetDecimal(0).ToString(); //Display form_id into Textbox which id captured.                        
                        TxtName.Text = gl_dbaseobj.cls_dr.GetString(1).ToUpper();
                        TxtFathername.Text = gl_dbaseobj.cls_dr.GetString(2).ToUpper();
                        TxtCnic.Text = gl_dbaseobj.cls_dr.GetDecimal(3).ToString();
                        TxtDob.Text = gl_dbaseobj.cls_dr.GetString(4).ToString();
                        //gender.(5)                            
                        if (gl_dbaseobj.cls_dr.GetString(5).ToUpper() == "M")
                        {
                            TxtGender.Text = "MALE";
                        }
                        else
                        {
                            TxtGender.Text = "FEMALE";
                        }
                        TxtReligion.Text = gl_dbaseobj.cls_dr.GetString(6).ToUpper();
                        TxtDomicile.Text = gl_dbaseobj.cls_dr.GetString(7).ToUpper();
                        Txtmail.Text = gl_dbaseobj.cls_dr.GetString(8).ToUpper();
                        TxtPhoneno.Text = gl_dbaseobj.cls_dr.GetDecimal(9).ToString();
                        TxtAccno.Text = gl_dbaseobj.cls_dr.GetDecimal(10).ToString();
                        TxtBankbranch.Text = gl_dbaseobj.cls_dr.GetString(11).ToUpper();
                        //--Formstatus_id.show_clsobj.dr.GetString(12).ToUpper();
                        //--Scanned_by.show_clsobj.dr.GetString(13).ToUpper();
                        //--Deo_by.show_clsobj.dr.GetString(14).ToUpper();
                        //--Hr_by.show_clsobj.dr.GetString(15).ToUpper();
                        //--Admin_by.show_clsobj.dr.GetString(16).ToUpper();
                        //--Password.show_clsobj.dr.GetString(17).ToUpper();
                        //--Account_status.show_clsobj.dr.GetString(18).ToUpper();
                        //TxtRftagno.Text=show_clsobj.dr.GetString(19).ToUpper();
                        //--Bps.show_clsobj.dr.GetString(20).ToUpper();
                        //--Job_id.show_clsobj.dr.GetString(21).ToUpper();
                        //--Mgr_id.show_clsobj.dr.GetString(22).ToUpper();
                        //--Dept_id.show_clsobj.dr.GetString(23).ToUpper();
                        //--Hire_date.show_clsobj.dr.GetString(24).ToUpper();                                                
                        //--Joining_date.show_clsobj.dr.GetString(25).ToUpper();                                                
                        if (gl_dbaseobj.cls_dr.GetString(26).ToUpper() == "C")
                        {
                            TxtEmployeetype.Text = "CONTRACT";
                        }
                        else
                        {
                            TxtEmployeetype.Text = "PERMANENT";
                        }
                        TxtTempaddress.Text = gl_dbaseobj.cls_dr.GetString(27).ToUpper();
                        TxtPermaddress.Text = gl_dbaseobj.cls_dr.GetString(28).ToUpper();

                        {   //Now Displaying Picture.   //Column # 29. Emp_Image.
                            
                            //MessageBox.Show("Employee Image Loading, Now");
                            OracleLob blob = gl_dbaseobj.cls_dr.GetOracleLob(29);
                            Byte[] BLOBData = new Byte[blob.Length];

                            //Read blob data into byte array
                            int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));

                            //Get the primitive byte data into in-memory data stream
                            System.IO.MemoryStream stmBLOBData = new System.IO.MemoryStream(BLOBData);

                            //LOADING INTO PICBOX1
                            PictureBox1.Image = Image.FromStream(stmBLOBData);
                            //MessageBox.Show("Now i'm reseting it.");
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                        }
                        {   //--Form_Image.show_clsobj.dr.GetString(30).ToUpper();                        
                            OracleLob blob = gl_dbaseobj.cls_dr.GetOracleLob(30);  //It is form_image
                            Byte[] BLOBData = new Byte[blob.Length];

                            //Read blob data into byte array
                            int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));

                            //Get the primitive byte data into in-memory data stream
                            System.IO.MemoryStream stmBLOBData = new System.IO.MemoryStream(BLOBData);

                            //LOADING INTO PICBOX1                            
                            PictureBox2.Image = Image.FromStream(stmBLOBData);
                            //MessageBox.Show("Now i'm reseting it.");
                            PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                        }
                        //14-Aug-2007***********************************Now Display Qualification data.**************************
                        string query4qualifcation = "Select Deg_Name,Pass_Year,Div_Grade,Majors,Institute From Qualification where form_id=" + form_id;  //Choose AGAIN THIS PARTICULAR FORM.                
                        gl_dbaseobj.Show_Datafun(query4qualifcation, 0);
                        if (gl_dbaseobj.cls_dr.Read())
                        {
                            TxtDegreename.Text = gl_dbaseobj.cls_dr[0].ToString();//.GetString(0).ToString().ToUpper();  //Display Deg_name from Qualifcation Table.
                            CmbPassyear.Text = gl_dbaseobj.cls_dr.GetDateTime(1).ToString();            //Display Pass_Year from Qualifcation Table.
                            TxtGrade.Text = gl_dbaseobj.cls_dr.GetString(2).ToString().ToUpper();       //Display Div_Grade from Qualifcation Table.
                            TxtMajor.Text = gl_dbaseobj.cls_dr.GetString(3).ToString().ToUpper();       //Display Majors from Qualifcation Table.
                            TxtInstitute.Text = gl_dbaseobj.cls_dr.GetString(4).ToString().ToUpper();   //Display Institute from Qualifcation Table.
                        }                        
                        
                    }//Now OFF, HR data displayed.
                    else 
                    {//If it is neither DEO or HR Then rase Error.
                        MessageBox.Show("Your are not Valid User to access forms.");
                    }
                }//End if
                else 
                {//Hey! why didn't u pick form.
                 //  MessageBox.Show("Some body took this form fist, Your Request was in the Queue.","" ,MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                gl_dbaseobj.Show_Datafun(choosequery, 1);//close conn.
                watch.Stop();   //Stop Time.
                if (BtnRefersh.Enabled != true)
                {
                    DeoLblTimeinsec.Text = "Time Load in Sec:" + watch.ElapsedMilliseconds / 1000 + " And in Millisec:" + watch.ElapsedMilliseconds;
                }
        }//End Fun.
        void set4_deo_scanfun(int v_setstatus)
        {
                if (v_setstatus == 0)  //Set for DEO.
                {
                    string query_set4deo = "Update Employee set Hr_by=" + Safeinfo.loginid + ", formstatus_id="
                                                                + 0
                                                                + " where form_id=" + form_id;
                    int effected_rows = gl_dbaseobj.Dml_Updatefun(query_set4deo);                 //Set this Formstatus_id=0.
                    //if effected_rows!=1 then throw exception.
                }
                else if (v_setstatus == -1) //Set for Scanner to delete form.
                {
                    string query_set4scanner = "Update Employee set Deo_By="+Safeinfo.loginid+", formstatus_id="
                                                                  + -1
                                                                  + " where form_id=" + form_id;
                    int effected_rows = gl_dbaseobj.Dml_Updatefun(query_set4scanner);                 //Set this Formstatus_id=-1.
                    //if effected_rows!=1 then throw exception.
                }            
        }
        protected void DEO_fun()
        {
            //Hide Dob controls for DEO.
            TxtDob.Visible = false;

            CmbDay.Visible = true;
            CmbMon.Visible = true;
            CmbYear.Visible = true;

            LblDeodd.Visible = true;
            LblDeoMon.Visible = true;
            LblDeoYear.Visible = true;
            
            //Hide  Gender controls for DEO.
            TxtGender.Visible = false;
            //Hide  Religion controls for DEO.
            TxtReligion.Visible = false;
            //Hide  Domicile controls for DEO.
            TxtDomicile.Visible = false;
            //Hide  Employeetype controls for DEO.
            TxtEmployeetype.Visible = false;
            //Hide  Extra controls for DEO.
            TxtRftagno.Hide();
            LblRftagno.Hide();
            TxtPwd.Hide();
            LblPwd.Hide();
            BtnEncrypt.Hide();
            domainUpDown1.Hide();
            LblAeskey.Hide();

            //Show DEO Gender controls.
            CmbGender.Visible = true;
            //Show DEO Religion controls.
            CmbReligion.Visible = true;
            //Show DEO Domicile controls.
            CmbDomicile.Visible = true;
            //Show DEO Emptype controls.
            CmbEmployeetype.Visible = true;
            
            if (Safeinfo.purpose_deo_hr_receptionst == 0)
            {
                BtnDeoScanMain.Show();//Show me Back going button. (Only Enable for Administrator).                                        
            }
            else
            {
                BtnDeoScanMain.Hide();
            }

            /********************************************************************************************************/
            fetch_datafun(Safeinfo.status.ToUpper(), false);   //I'm Not skiping Rather fetch record so 2nd arg=false
            load_recfun();                                    //Show me that record which u've setted.  
        }//Deo_fun ends.
        protected void HR_fun()
        {
            //Hide DEO Dob controls.
            CmbDay.Visible = false;
            CmbMon.Visible = false;
            CmbYear.Visible = false;
            LblDeodd.Visible = false;
            LblDeoMon.Visible = false;
            LblDeoYear.Visible = false;
            //Hide DEO Gender controls.
            CmbGender.Visible = false;
            //Hide DEO Religion controls.
            CmbReligion.Visible = false;
            //Hide DEO Domicile controls.
            CmbDomicile.Visible = false;
            //Hide DEO Emptype controls.
            CmbEmployeetype.Visible = false;

            //Show Dob controls for HR.
            TxtDob.Visible = true;
            //Show Gender controls for HR.
            TxtGender.Visible = true;
            //Show Religion controls for HR.
            TxtReligion.Visible = true;
            //Show Domicile controls for HR.
            TxtDomicile.Visible = true;
            //Show Domicile controls for HR.
            TxtEmployeetype.Visible = true;
            //Show Extra Control 'Rftagno'
            TxtRftagno.Enabled = true;
            LblRftagno.Enabled = true;
            LblAeskey.Show();
            

            if (Safeinfo.purpose_deo_hr_receptionst == 1)
            {
                BtnDeoScanMain.Show();//Show me Back going button. (Only Enable for Administrator).                                        
            }
            else
            {
                BtnDeoScanMain.Hide();
            }
            disable_ctlfun(true);//Disable controls bcz I am HR.                
            /********************************************************************************************************/
            fetch_datafun(Safeinfo.status.ToUpper(), false);   //I'm Not skiping Rather fetch record so 2nd arg=false
            load_recfun();                                    //Show me that record which u've setted.  
            /********************************************************************************************************/
        }//hr_fun ends.
        int  save_recfun(int setformstatus)               //This is for DEO save function.
        {//Here is in below, Function overloading concept.
            try
            {//First it will set Date_Birth, Be carefull you can have Fast logic.
                //
                //
                //only_4_dob.Dml_Updatefun(
                string final_date = null;  //It will concatenate below sdd,smon,syyyy.then you can use final_date as date.
                string sdd = Convert.ToString(CmbDay.SelectedItem);//ok
                string smon = Convert.ToString(CmbMon.SelectedItem);//ok                
                string syyyy = Convert.ToString(CmbYear.SelectedItem);//ok.                    
                final_date = sdd + "-" + smon + "-" + syyyy;
                string update_only_datebirth = "Update Pk_ocp.Employee set Date_Birth= To_date('" + final_date.ToUpper() + "'," + "'DD-MON-RRRR') where Form_id=" + form_id;                    
                
                if (gl_dbaseobj.Dml_Updatefun(update_only_datebirth) == 1)
                   {
                       //Now save data into Employee Table.
                       gl_dbaseobj.Dml_UpdateAdapterfun("Select * From Employee Where form_id=" + form_id);
                       gl_dbaseobj.cls_dataset.Tables[0].Rows[0][1] = TxtName.Text.ToUpper();
                       gl_dbaseobj.cls_dataset.Tables[0].Rows[0][2] = TxtFathername.Text.ToUpper();
                       gl_dbaseobj.cls_dataset.Tables[0].Rows[0][3] = Convert.ToInt64(TxtCnic.Text);
                        //save_clsobj.cls_dataset.Tables[0].Rows[0][4] = OracleDateTime.Parse("21-NOV-1978");   //Date_birth.
                        string match_string0 = "MALE";
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][5] = (match_string0 == Convert.ToString(CmbGender.SelectedItem)) ? "M" : "F";
                        switch (Convert.ToString(CmbReligion.SelectedItem))
                        {
                            case "ISLAM":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][6] = "ISLAM";
                                    break;
                                }
                            case "CHRISTIAN":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][6] = "CHRISTIAN";
                                    break;
                                }
                            case "HINDU":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][6] = "HINDU";
                                    break;
                                }
                            case "SIKH":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][6] = "SIKH";
                                    break;
                                }
                            case "BUDH":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][6] = "BUDH";
                                    break;
                                }
                            case "QADIANI/AHMEDI":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][6] = "QADIANI/AHMEDI";
                                    break;
                                }/*
                          * default:
                          * {
                          * save_clsobj.cls_dataset.Tables[0].Rows[0][6] = "What ever u wrote";
                          * break;
                          * }*/
                        }//Religion switch end.
                        switch (Convert.ToString(CmbDomicile.SelectedItem))
                        {
                            case "FATA":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][7] = "FATA";
                                    break;
                                }
                            case "NWFP":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][7] = "NWFP";
                                    break;
                                }
                            case "PUNJAB":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][7] = "PUNJAB";
                                    break;
                                }
                            case "SINDH":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][7] = "SINDH";
                                    break;
                                }
                            case "BALOCHISTAN":
                                {
                                    gl_dbaseobj.cls_dataset.Tables[0].Rows[0][7] = "BALOCHISTAN";
                                    break;
                                }/*
                          * default:
                          * {
                          * save_clsobj.cls_dataset.Tables[0].Rows[0][7] = "What ever u wrote";
                          * break;
                          * }*/
                        }//Domicile switch end.

                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][8] = Txtmail.Text.ToUpper();
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][9] = Convert.ToInt64(TxtPhoneno.Text);
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][10] = int.Parse(TxtAccno.Text);
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][11] = TxtBankbranch.Text.ToUpper();
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][12] = setformstatus;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][13] = scannedby;
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][14] = Safeinfo.loginid;//Deo_by
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][15] = Safeinfo.hr;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][16] = Safeinfo.Admin;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][17] = Password
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][18] = Account_status;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][19] = Rfid_id
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][20] = bps;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][21] = job_id;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][22] = mgr_id;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][23] = dept_id;
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][24] = Hire_date; when you assigned job,bps then you hired.
                        //gl_dbaseobj.cls_dataset.Tables[0].Rows[0][25] = Joining_date; when you assigned job,bps then you set it.
                        match_string0 = "CONTRACT";
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][26] = (match_string0 == Convert.ToString(CmbEmployeetype.SelectedItem)) ? "C" : "P";
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][27] = TxtTempaddress.Text;
                        gl_dbaseobj.cls_dataset.Tables[0].Rows[0][28] = TxtPermaddress.Text;
                        gl_dbaseobj.cls_odadptr.Update(gl_dbaseobj.cls_dataset);
                    //******************************************************************************************//
                    //Now Insert Row into Qualification Table.
                    String strinsrtquery="Insert Into Qualification Values("+ form_id+",'"+TxtDegreename.Text+"',Sysdate,'"+TxtGrade.Text+"','"+TxtMajor.Text+"','"+TxtInstitute.Text+"')";
                    gl_dbaseobj.Dml_UpdateAdapterfun(strinsrtquery);
                    return (1);//Its mean record successfully saved.
                }       //Date_Birth and other data enterd.
                else 
                {   
                    string back_query = "Update employee set formstatus_id=0 where form_id=" + form_id;
                    if (gl_dbaseobj.Dml_Updatefun(back_query) == 1)
                    {
                        MessageBox.Show("Form is going to set back.");
                    }
                    set_clearfields();
                    return (0);//Its mean record Can not 
                }
            }//Try block end.
            catch (OracleException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);                
                string back_query = "Update employee set formstatus_id=0 where form_id=" + form_id;
                if (gl_dbaseobj.Dml_Updatefun(back_query) == 1)
                {
                    MessageBox.Show("Form is going to set back.");
                }
                set_clearfields();                
                return (0);//Its mean record Can not be saved.
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"General Error Found",MessageBoxButtons.OK,MessageBoxIcon.Error);
                string back_query = "Update employee set formstatus_id=0 where form_id=" + form_id;
                if (gl_dbaseobj.Dml_Updatefun(back_query) == 1)
                {
                    MessageBox.Show("Form is going to set back.");
                }
                set_clearfields();
                return (0);//Its mean record Can not be saved.
            }
            //save_clsobj.cls_dataset.Tables[0].Rows[0][27] = Emp_pic;
            //save_clsobj.cls_dataset.Tables[0].Rows[0][27] = Form_pci
        }//End save 4 deo func.
        int  save_recfun(int formstatus_id, int rftag_no) //This is for HR save function.
        {//Here is Upper,Function overloading concept.
           
            //Here we assign Hiredate and RFTAG No.
            string update_query = "Update Employee set formstatus_id=" + formstatus_id +
                                        ",Hr_by=" + Safeinfo.loginid +
                                        ",rfid_id=" + rftag_no +
                                        ",Hire_date= Sysdate" +
                                        " where form_id=" + form_id;
            if (gl_dbaseobj.Dml_Updatefun(update_query) == 1)
            {
                //Here you INSERT RECORD INTO ATTENDANCE TABLE. Bcz This Person is now going to JOIN his job.
                string insert_attendance = "Insert into Attendance Values(" + form_id + ",SYSDATE,'FIRST TIME JOIN',00,00,00)";
                gl_dbaseobj.Dml_Updatefun(insert_attendance);
                return 1;
            }
            else 
            {
                string back_query = "Update employee set formstatus_id=2 where form_id=" + form_id;
                if (gl_dbaseobj.Dml_Updatefun(back_query) == 1)
                {
                    //MessageBox.Show("Form is going to set back.");
                }
                return 0;
            }            
        }
        void disp_rectanglefun(Point mostleft, Size rectanglesize)
        {
            if (TxtFormid.Text != "")
            {
                PictureBox2.Refresh();
                Rectangle myrect = new Rectangle(mostleft, rectanglesize);
                Graphics dc = PictureBox2.CreateGraphics();
                Pen RedPen = new Pen(Color.Red, 5);
                dc.DrawRectangle(RedPen, myrect);
            }        
        }
        void disable_ctlfun(bool ctl_enable_true_false)
        {
            if (Safeinfo.status == "HR" || Safeinfo.purpose_deo_hr_receptionst == 1)
            {//Bcz I'm HR i ALWAYS want to see DISABLE CONTROLS.
                TxtName.ReadOnly = true;
                TxtFathername.ReadOnly = true;
                TxtCnic.ReadOnly = true;
                TxtDob.ReadOnly = true;
                CmbDay.Enabled = false;
                CmbMon.Enabled = false;
                CmbYear.Enabled = false;
                CmbGender.Enabled = false;
                CmbReligion.Enabled = false;
                CmbDomicile.Enabled = false;
                Txtmail.ReadOnly = true;                
                TxtPhoneno.ReadOnly = true;
                TxtAccno.ReadOnly = true;
                TxtBankbranch.ReadOnly = true;
                TxtEmployeetype.ReadOnly = true;
                TxtTempaddress.ReadOnly = true;
                TxtPermaddress.ReadOnly = true;    
                //Disable also Qualification Controls.
                TxtDegreename.ReadOnly = true;
                CmbPassyear.Enabled=false;
                TxtGrade.ReadOnly = true;
                TxtMajor.ReadOnly = true;
                TxtInstitute.ReadOnly = true;                
            }

            else //you want to Enable controls.
            {
                if (ctl_enable_true_false)//'TRUE' mean you want to disable controls.
                {
                    TxtName.Enabled = false;
                    TxtFathername.Enabled = false;
                    TxtCnic.Enabled = false;
                    CmbDay.Enabled = false;
                    CmbMon.Enabled = false;
                    CmbYear.Enabled = false;
                    CmbGender.Enabled = false;
                    CmbReligion.Enabled = false;
                    CmbDomicile.Enabled = false;
                    Txtmail.Enabled = false;
                    TxtPhoneno.Enabled = false;
                    TxtAccno.Enabled = false;
                    TxtBankbranch.Enabled = false;
                    CmbEmployeetype.Enabled = false;
                    TxtTempaddress.Enabled = false;
                    TxtPermaddress.Enabled = false;
                    //Disable Qualification Controls.
                    TxtDegreename.Enabled = false;
                    CmbPassyear.Enabled = false;
                    TxtGrade.Enabled = false;
                    TxtMajor.Enabled = false;
                    TxtInstitute.Enabled = false;

                }
                else//If it gave 'False', then its mean you want to ENABLE controls.
                {
                    TxtName.ReadOnly = false;
                    TxtFathername.ReadOnly = false;
                    TxtCnic.ReadOnly = false;
                    CmbDay.Enabled = true;
                    CmbMon.Enabled = true;
                    CmbYear.Enabled = true;
                    CmbGender.Enabled = true;
                    CmbReligion.Enabled = true;
                    CmbDomicile.Enabled = true;
                    Txtmail.ReadOnly = false;
                    TxtPhoneno.ReadOnly = false;
                    TxtAccno.ReadOnly = false;
                    TxtBankbranch.ReadOnly = false;
                    CmbEmployeetype.Enabled = true;
                    TxtTempaddress.ReadOnly = false;
                    TxtPermaddress.ReadOnly = false;


                    //Enable Qualification Controls.
                    TxtDegreename.Enabled = true;
                    CmbPassyear.Enabled = true;
                    TxtGrade.Enabled = true;
                    TxtMajor.Enabled = true;
                    TxtInstitute.Enabled = true;

                }
            }
        
        }

        bool Deo_validate_field()
        {
            Boolean should_i_save = true;

            {//Here is in below, Function overloading concept.
                
                    //Now Verify Field's data, before saving.
                    //--Verify Name.
                if (TxtName.Text.Length < 3 || TxtName.Text.Length > 25)
                {
                    LblName.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LblName.ForeColor = Color.Black;
                    
                }
                    //--Verify Father Name.
                if (TxtFathername.Text.Length < 3 || TxtFathername.Text.Length > 25)
                {
                    LblFathername.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LblFathername.ForeColor = Color.Black;                    
                }
                    //--Verify Cnic.
                if (TxtCnic.Text.Length < 13 || TxtCnic.Text.Length > 13)
                {
                    LCnic.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LCnic.ForeColor = Color.Black;                    
                }
                    //--Verify Date of Birth.
                if (CmbDay.Text == "" || CmbMon.Text == "" || CmbYear.Text == "")
                {
                    LblDob.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LblDob.ForeColor = Color.Black;                    
                }
                    //--Verify Gender.
                if (CmbGender.Text == "")
                {
                    LGender.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LGender.ForeColor = Color.Black;
                }
                //--Verify Religion.
                if (CmbReligion.Text == "")
                {
                    LReligion.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LReligion.ForeColor = Color.Black;
                }
                    //--Verify Domicile.
                if (CmbDomicile.Text == "")
                {
                    LDomicile.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LDomicile.ForeColor = Color.Black;
                }
                //--Verify Email.
                if (Txtmail.Text.Length<7 )
                {
                    LEmail.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else if (Txtmail.Text.Contains("@") == false)
                {
                    LEmail.ForeColor = Color.Red;
                    should_i_save = false;
                }
                else
                {
                    LEmail.ForeColor = Color.Black;                       
                }
                //--Verify Cell #.
                if (TxtPhoneno.Text.Length < 7 || TxtPhoneno.Text.Length > 11)
                    {
                        LPhoneno.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LPhoneno.ForeColor = Color.Black;                        
                    }
                    //--Verify Bank Account #.
                    if (TxtAccno.Text.Length < 5 || TxtPhoneno.Text.Length > 10)
                    {
                        LblAccno.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LblAccno.ForeColor = Color.Black;                        
                    }
                    //--Verify Bank Branch.
                    if (TxtBankbranch.Text.Length < 5 || TxtBankbranch.Text.Length > 40)
                    {
                        LblBankbranch.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LblBankbranch.ForeColor = Color.Black;                        
                    }
                    //--Verify Employee Type.
                    if (CmbEmployeetype.Text == "")
                    {
                        LEmployeetype.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LEmployeetype.ForeColor = Color.Black;                       
                    }
                    //--Verify Degree Name.
                    if (TxtDegreename.Text.Length < 2 || TxtDegreename.Text.Length > 10)
                    {
                        LblDegreename.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LblDegreename.ForeColor = Color.Black;                        
                    }
                    //--Verify Pass year.
                    if (CmbPassyear.Text == "")
                    {
                        LblPassyear.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LblPassyear.ForeColor = Color.Black;                        
                    }
                    //--Verify grade.
                    if (TxtGrade.Text == "")
                    {
                        LblGrade.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LblGrade.ForeColor = Color.Black;                        
                    }
                    //--Verify Major.
                    if (TxtMajor.Text == "")
                    {
                        LblMajor.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LblMajor.ForeColor = Color.Black;                        
                    }
                    //--Verify Institute.
                    if (TxtInstitute.Text == "")
                    {
                        LblInstitute.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LblInstitute.ForeColor = Color.Black;                        
                    }

                    //--Verify Temp address.
                    if (TxtTempaddress.Text.Length < 5 || TxtTempaddress.Text.Length > 60)
                    {
                        LTempaddress.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LTempaddress.ForeColor = Color.Black;                        
                    }
                    //--Verify Perm Address.
                    if (TxtPermaddress.Text.Length < 5 || TxtPermaddress.Text.Length > 60)
                    {
                        LPermaddress.ForeColor = Color.Red;
                        should_i_save = false;
                    }
                    else
                    {
                        LPermaddress.ForeColor = Color.Black;
                        
                    }

                    return (should_i_save);
            }
        }
        //int Hr_validate_field(int rfid_id)
        //{ //Check RFTAG_ID exist in Database.

        //        Deo_Module.Update_Database Temp_obj = new Update_Database();
        //        if (Temp_obj.Record_Foundfun("Select * from Tag_ids Where Tag_id=" + rfid_id) == 1)
        //        { //ITs mean record found against this tag_id and record_foundfun() return(1))
        //            LblRftagno.ForeColor = Color.Black;
        //            return (1);

        //        }
        //        else
        //        {//ITs mean Against this Tag_id Record not found and record_foundfun() return(0))                   
        //            LblRftagno.ForeColor = Color.Red;
        //            return (0);
        //        }
        //}







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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Dataentry));
            this.LblRftagno = new System.Windows.Forms.Label();
            this.TxtRftagno = new System.Windows.Forms.TextBox();
            this.LEmployeetype = new System.Windows.Forms.Label();
            this.CmbEmployeetype = new System.Windows.Forms.ComboBox();
            this.CmbYear = new System.Windows.Forms.ComboBox();
            this.CmbMon = new System.Windows.Forms.ComboBox();
            this.CmbDay = new System.Windows.Forms.ComboBox();
            this.CmbDomicile = new System.Windows.Forms.ComboBox();
            this.CmbReligion = new System.Windows.Forms.ComboBox();
            this.LGender = new System.Windows.Forms.Label();
            this.CmbGender = new System.Windows.Forms.ComboBox();
            this.PictureBox2 = new System.Windows.Forms.PictureBox();
            this.LPermaddress = new System.Windows.Forms.Label();
            this.TxtPermaddress = new System.Windows.Forms.TextBox();
            this.LTempaddress = new System.Windows.Forms.Label();
            this.TxtTempaddress = new System.Windows.Forms.TextBox();
            this.LPhoneno = new System.Windows.Forms.Label();
            this.TxtPhoneno = new System.Windows.Forms.TextBox();
            this.LReligion = new System.Windows.Forms.Label();
            this.LCnic = new System.Windows.Forms.Label();
            this.TxtCnic = new System.Windows.Forms.TextBox();
            this.LEmail = new System.Windows.Forms.Label();
            this.LDomicile = new System.Windows.Forms.Label();
            this.Txtmail = new System.Windows.Forms.TextBox();
            this.LblDob = new System.Windows.Forms.Label();
            this.LblFathername = new System.Windows.Forms.Label();
            this.LblName = new System.Windows.Forms.Label();
            this.TxtFathername = new System.Windows.Forms.TextBox();
            this.TxtName = new System.Windows.Forms.TextBox();
            this.BtnLogout = new System.Windows.Forms.Button();
            this.BtnSave = new System.Windows.Forms.Button();
            this.LLoginname = new System.Windows.Forms.Label();
            this.LLoginas = new System.Windows.Forms.Label();
            this.TxtLoginname = new System.Windows.Forms.TextBox();
            this.TxtLoginas = new System.Windows.Forms.TextBox();
            this.BtnSkip = new System.Windows.Forms.Button();
            this.TxtFormid = new System.Windows.Forms.TextBox();
            this.LblFormid = new System.Windows.Forms.Label();
            this.BtnRefersh = new System.Windows.Forms.Button();
            this.BtnDeoScanMain = new System.Windows.Forms.Button();
            this.LblAccno = new System.Windows.Forms.Label();
            this.TxtAccno = new System.Windows.Forms.TextBox();
            this.LblBankbranch = new System.Windows.Forms.Label();
            this.TxtBankbranch = new System.Windows.Forms.TextBox();
            this.LblDeodd = new System.Windows.Forms.Label();
            this.LblDeoMon = new System.Windows.Forms.Label();
            this.LblDeoYear = new System.Windows.Forms.Label();
            this.LblCursorlocation = new System.Windows.Forms.Label();
            this.DeoLblTimeinsec = new System.Windows.Forms.Label();
            this.TxtDob = new System.Windows.Forms.TextBox();
            this.TxtGender = new System.Windows.Forms.TextBox();
            this.TxtReligion = new System.Windows.Forms.TextBox();
            this.TxtDomicile = new System.Windows.Forms.TextBox();
            this.TxtEmployeetype = new System.Windows.Forms.TextBox();
            this.PictureBox1 = new System.Windows.Forms.PictureBox();
            this.BtnEncrypt = new System.Windows.Forms.Button();
            this.domainUpDown1 = new System.Windows.Forms.DomainUpDown();
            this.TxtPwd = new System.Windows.Forms.TextBox();
            this.LblPwd = new System.Windows.Forms.Label();
            this.LblAeskey = new System.Windows.Forms.Label();
            this.LblDegreename = new System.Windows.Forms.Label();
            this.TxtDegreename = new System.Windows.Forms.TextBox();
            this.LblPassyear = new System.Windows.Forms.Label();
            this.CmbPassyear = new System.Windows.Forms.ComboBox();
            this.LblGrade = new System.Windows.Forms.Label();
            this.TxtGrade = new System.Windows.Forms.TextBox();
            this.LblMajor = new System.Windows.Forms.Label();
            this.TxtMajor = new System.Windows.Forms.TextBox();
            this.LblInstitute = new System.Windows.Forms.Label();
            this.TxtInstitute = new System.Windows.Forms.TextBox();
            this.BtnTag_addition = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // LblRftagno
            // 
            this.LblRftagno.AutoSize = true;
            this.LblRftagno.ForeColor = System.Drawing.Color.Black;
            this.LblRftagno.Location = new System.Drawing.Point(734, 709);
            this.LblRftagno.Name = "LblRftagno";
            this.LblRftagno.Size = new System.Drawing.Size(53, 13);
            this.LblRftagno.TabIndex = 46;
            this.LblRftagno.Text = "RF Tag #";
            // 
            // TxtRftagno
            // 
            this.TxtRftagno.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtRftagno.Location = new System.Drawing.Point(787, 705);
            this.TxtRftagno.Name = "TxtRftagno";
            this.TxtRftagno.Size = new System.Drawing.Size(77, 20);
            this.TxtRftagno.TabIndex = 47;
            
            this.TxtRftagno.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtRftagno_KeyPress);
            this.TxtRftagno.TextChanged += new System.EventHandler(this.TxtRftagno_TextChanged);
            // 
            // LEmployeetype
            // 
            this.LEmployeetype.AutoSize = true;
            this.LEmployeetype.Location = new System.Drawing.Point(698, 639);
            this.LEmployeetype.Name = "LEmployeetype";
            this.LEmployeetype.Size = new System.Drawing.Size(80, 13);
            this.LEmployeetype.TabIndex = 30;
            this.LEmployeetype.Text = "Employee Type";
            // 
            // CmbEmployeetype
            // 
            this.CmbEmployeetype.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbEmployeetype.FormattingEnabled = true;
            this.CmbEmployeetype.Items.AddRange(new object[] {
            "CONTRACT",
            "PERMANENT"});
            this.CmbEmployeetype.Location = new System.Drawing.Point(781, 635);
            this.CmbEmployeetype.Name = "CmbEmployeetype";
            this.CmbEmployeetype.Size = new System.Drawing.Size(83, 21);
            this.CmbEmployeetype.TabIndex = 32;
            this.CmbEmployeetype.Leave += new System.EventHandler(this.CmbEmployeetype_Leave);
            this.CmbEmployeetype.Enter += new System.EventHandler(this.CmbEmployeetype_Enter);
            // 
            // CmbYear
            // 
            this.CmbYear.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbYear.FormattingEnabled = true;
            this.CmbYear.Items.AddRange(new object[] {
            "1970",
            "1971",
            "1972",
            "1973",
            "1974",
            "1975",
            "1976",
            "1977",
            "1978",
            "1979",
            "1980",
            "1981",
            "1982",
            "1983",
            "1984",
            "1985",
            "1986",
            "1987",
            "1988",
            "1989",
            "1990"});
            this.CmbYear.Location = new System.Drawing.Point(814, 587);
            this.CmbYear.Name = "CmbYear";
            this.CmbYear.Size = new System.Drawing.Size(50, 21);
            this.CmbYear.TabIndex = 14;
            this.CmbYear.Leave += new System.EventHandler(this.CmbYear_Leave);
            // 
            // CmbMon
            // 
            this.CmbMon.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbMon.FormattingEnabled = true;
            this.CmbMon.Items.AddRange(new object[] {
            "Jan",
            "Feb",
            "Mar",
            "Apr",
            "May",
            "Jun",
            "Jul",
            "Aug",
            "Sep",
            "Oct",
            "Nov",
            "Dec"});
            this.CmbMon.Location = new System.Drawing.Point(724, 587);
            this.CmbMon.Name = "CmbMon";
            this.CmbMon.Size = new System.Drawing.Size(51, 21);
            this.CmbMon.TabIndex = 12;
            this.CmbMon.Leave += new System.EventHandler(this.CmbMon_Leave);
            // 
            // CmbDay
            // 
            this.CmbDay.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbDay.FormattingEnabled = true;
            this.CmbDay.Items.AddRange(new object[] {
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26",
            "27",
            "28",
            "29",
            "30",
            "31"});
            this.CmbDay.Location = new System.Drawing.Point(650, 587);
            this.CmbDay.Name = "CmbDay";
            this.CmbDay.Size = new System.Drawing.Size(42, 21);
            this.CmbDay.TabIndex = 10;
            this.CmbDay.Leave += new System.EventHandler(this.CmbDay_Leave);
            this.CmbDay.Enter += new System.EventHandler(this.CmbDay_Enter);
            // 
            // CmbDomicile
            // 
            this.CmbDomicile.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbDomicile.FormattingEnabled = true;
            this.CmbDomicile.Items.AddRange(new object[] {
            "FATA",
            "NWFP",
            "PUNJAB",
            "SINDH",
            "BALOCHISTAN"});
            this.CmbDomicile.Location = new System.Drawing.Point(464, 611);
            this.CmbDomicile.Name = "CmbDomicile";
            this.CmbDomicile.Size = new System.Drawing.Size(92, 21);
            this.CmbDomicile.TabIndex = 21;
            this.CmbDomicile.Leave += new System.EventHandler(this.CmbDomicile_Leave);
            this.CmbDomicile.Enter += new System.EventHandler(this.CmbDomicile_Enter);
            // 
            // CmbReligion
            // 
            this.CmbReligion.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbReligion.FormattingEnabled = true;
            this.CmbReligion.Items.AddRange(new object[] {
            "ISLAM",
            "CHRISTIAN",
            "HINDU",
            "SIKH",
            "BUDH",
            "QADIANI/AHMEDI"});
            this.CmbReligion.Location = new System.Drawing.Point(248, 611);
            this.CmbReligion.Name = "CmbReligion";
            this.CmbReligion.Size = new System.Drawing.Size(159, 21);
            this.CmbReligion.TabIndex = 19;
            this.CmbReligion.Leave += new System.EventHandler(this.CmbReligion_Leave);
            this.CmbReligion.Enter += new System.EventHandler(this.CmbReligion_Enter);
            // 
            // LGender
            // 
            this.LGender.AutoSize = true;
            this.LGender.Location = new System.Drawing.Point(5, 615);
            this.LGender.Name = "LGender";
            this.LGender.Size = new System.Drawing.Size(42, 13);
            this.LGender.TabIndex = 15;
            this.LGender.Text = "Gender";
            // 
            // CmbGender
            // 
            this.CmbGender.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbGender.FormattingEnabled = true;
            this.CmbGender.Items.AddRange(new object[] {
            "MALE",
            "FEMALE"});
            this.CmbGender.Location = new System.Drawing.Point(49, 611);
            this.CmbGender.Name = "CmbGender";
            this.CmbGender.Size = new System.Drawing.Size(123, 21);
            this.CmbGender.TabIndex = 17;
            this.CmbGender.Leave += new System.EventHandler(this.CmbGender_Leave);
            this.CmbGender.Enter += new System.EventHandler(this.CmbGender_Enter);
            // 
            // PictureBox2
            // 
            this.PictureBox2.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.PictureBox2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("PictureBox2.BackgroundImage")));
            this.PictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.PictureBox2.Location = new System.Drawing.Point(2, 15);
            this.PictureBox2.Name = "PictureBox2";
            this.PictureBox2.Size = new System.Drawing.Size(862, 571);
            this.PictureBox2.TabIndex = 56;
            this.PictureBox2.TabStop = false;
            this.PictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(this.PictureBox1_MouseMove);
            // 
            // LPermaddress
            // 
            this.LPermaddress.AutoSize = true;
            this.LPermaddress.Location = new System.Drawing.Point(414, 685);
            this.LPermaddress.Name = "LPermaddress";
            this.LPermaddress.Size = new System.Drawing.Size(72, 13);
            this.LPermaddress.TabIndex = 44;
            this.LPermaddress.Text = "Perm Address";
            // 
            // TxtPermaddress
            // 
            this.TxtPermaddress.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtPermaddress.Location = new System.Drawing.Point(491, 681);
            this.TxtPermaddress.Name = "TxtPermaddress";
            this.TxtPermaddress.Size = new System.Drawing.Size(373, 20);
            this.TxtPermaddress.TabIndex = 45;
            this.TxtPermaddress.Enter += new System.EventHandler(this.TxtPermaddress_Enter);
            this.TxtPermaddress.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtPermaddress_KeyPress);
            // 
            // LTempaddress
            // 
            this.LTempaddress.AutoSize = true;
            this.LTempaddress.Location = new System.Drawing.Point(5, 688);
            this.LTempaddress.Name = "LTempaddress";
            this.LTempaddress.Size = new System.Drawing.Size(75, 13);
            this.LTempaddress.TabIndex = 42;
            this.LTempaddress.Text = "Temp Address";
            // 
            // TxtTempaddress
            // 
            this.TxtTempaddress.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtTempaddress.Location = new System.Drawing.Point(100, 684);
            this.TxtTempaddress.Name = "TxtTempaddress";
            this.TxtTempaddress.Size = new System.Drawing.Size(307, 20);
            this.TxtTempaddress.TabIndex = 43;
            this.TxtTempaddress.Enter += new System.EventHandler(this.TxtTempaddress_Enter);
            this.TxtTempaddress.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtTempaddress_KeyPress);
            // 
            // LPhoneno
            // 
            this.LPhoneno.AutoSize = true;
            this.LPhoneno.Location = new System.Drawing.Point(5, 639);
            this.LPhoneno.Name = "LPhoneno";
            this.LPhoneno.Size = new System.Drawing.Size(41, 13);
            this.LPhoneno.TabIndex = 24;
            this.LPhoneno.Text = "Cell No";
            // 
            // TxtPhoneno
            // 
            this.TxtPhoneno.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtPhoneno.Location = new System.Drawing.Point(49, 635);
            this.TxtPhoneno.Name = "TxtPhoneno";
            this.TxtPhoneno.Size = new System.Drawing.Size(123, 20);
            this.TxtPhoneno.TabIndex = 25;
            this.TxtPhoneno.Enter += new System.EventHandler(this.TxtPhoneno_Enter);
            this.TxtPhoneno.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtPhoneno_KeyPress);
            // 
            // LReligion
            // 
            this.LReligion.AutoSize = true;
            this.LReligion.Location = new System.Drawing.Point(179, 615);
            this.LReligion.Name = "LReligion";
            this.LReligion.Size = new System.Drawing.Size(45, 13);
            this.LReligion.TabIndex = 18;
            this.LReligion.Text = "Religion";
            // 
            // LCnic
            // 
            this.LCnic.AutoSize = true;
            this.LCnic.Location = new System.Drawing.Point(414, 592);
            this.LCnic.Name = "LCnic";
            this.LCnic.Size = new System.Drawing.Size(28, 13);
            this.LCnic.TabIndex = 6;
            this.LCnic.Text = "Cnic";
            // 
            // TxtCnic
            // 
            this.TxtCnic.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtCnic.Location = new System.Drawing.Point(442, 588);
            this.TxtCnic.Name = "TxtCnic";
            this.TxtCnic.Size = new System.Drawing.Size(114, 20);
            this.TxtCnic.TabIndex = 7;
            this.TxtCnic.Enter += new System.EventHandler(this.TxtCnic_Enter);
            this.TxtCnic.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtCnic_KeyPress);
            // 
            // LEmail
            // 
            this.LEmail.AutoSize = true;
            this.LEmail.Location = new System.Drawing.Point(560, 615);
            this.LEmail.Name = "LEmail";
            this.LEmail.Size = new System.Drawing.Size(32, 13);
            this.LEmail.TabIndex = 22;
            this.LEmail.Text = "Email";
            // 
            // LDomicile
            // 
            this.LDomicile.AutoSize = true;
            this.LDomicile.Location = new System.Drawing.Point(414, 615);
            this.LDomicile.Name = "LDomicile";
            this.LDomicile.Size = new System.Drawing.Size(47, 13);
            this.LDomicile.TabIndex = 20;
            this.LDomicile.Text = "Domicile";
            // 
            // Txtmail
            // 
            this.Txtmail.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.Txtmail.Location = new System.Drawing.Point(650, 611);
            this.Txtmail.Name = "Txtmail";
            this.Txtmail.Size = new System.Drawing.Size(214, 20);
            this.Txtmail.TabIndex = 23;
            this.Txtmail.Enter += new System.EventHandler(this.Txtmail_Enter);
            this.Txtmail.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Txtmail_KeyPress);
            // 
            // LblDob
            // 
            this.LblDob.AutoSize = true;
            this.LblDob.Location = new System.Drawing.Point(560, 591);
            this.LblDob.Name = "LblDob";
            this.LblDob.Size = new System.Drawing.Size(69, 13);
            this.LblDob.TabIndex = 8;
            this.LblDob.Text = "Date of Birth ";
            // 
            // LblFathername
            // 
            this.LblFathername.AutoSize = true;
            this.LblFathername.Location = new System.Drawing.Point(179, 592);
            this.LblFathername.Name = "LblFathername";
            this.LblFathername.Size = new System.Drawing.Size(68, 13);
            this.LblFathername.TabIndex = 4;
            this.LblFathername.Text = "Father Name";
            // 
            // LblName
            // 
            this.LblName.AutoSize = true;
            this.LblName.Location = new System.Drawing.Point(5, 592);
            this.LblName.Name = "LblName";
            this.LblName.Size = new System.Drawing.Size(35, 13);
            this.LblName.TabIndex = 2;
            this.LblName.Text = "Name";
            // 
            // TxtFathername
            // 
            this.TxtFathername.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtFathername.Location = new System.Drawing.Point(248, 588);
            this.TxtFathername.Name = "TxtFathername";
            this.TxtFathername.Size = new System.Drawing.Size(160, 20);
            this.TxtFathername.TabIndex = 5;
            this.TxtFathername.Enter += new System.EventHandler(this.TxtFathername_Enter);            
            this.TxtFathername.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtFathername_KeyPress);
            // 
            // TxtName
            // 
            this.TxtName.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtName.Location = new System.Drawing.Point(49, 588);
            this.TxtName.MaxLength = 40;
            this.TxtName.Name = "TxtName";
            this.TxtName.Size = new System.Drawing.Size(123, 20);
            this.TxtName.TabIndex = 3;
            this.TxtName.Enter += new System.EventHandler(this.TxtName_Enter);
            this.TxtName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtName_KeyPress);
            // 
            // BtnLogout
            // 
            this.BtnLogout.Location = new System.Drawing.Point(915, 671);
            this.BtnLogout.Name = "BtnLogout";
            this.BtnLogout.Size = new System.Drawing.Size(67, 21);
            this.BtnLogout.TabIndex = 54;
            this.BtnLogout.Text = "Log out";
            this.BtnLogout.UseVisualStyleBackColor = true;
            this.BtnLogout.Click += new System.EventHandler(this.button1_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(871, 525);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(67, 21);
            this.BtnSave.TabIndex = 51;
            this.BtnSave.Text = "Save";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // LLoginname
            // 
            this.LLoginname.AutoSize = true;
            this.LLoginname.ForeColor = System.Drawing.SystemColors.ControlText;
            this.LLoginname.Location = new System.Drawing.Point(865, 374);
            this.LLoginname.Name = "LLoginname";
            this.LLoginname.Size = new System.Drawing.Size(64, 13);
            this.LLoginname.TabIndex = 60;
            this.LLoginname.Text = "Login Name";
            // 
            // LLoginas
            // 
            this.LLoginas.AutoSize = true;
            this.LLoginas.ForeColor = System.Drawing.SystemColors.ControlText;
            this.LLoginas.Location = new System.Drawing.Point(865, 396);
            this.LLoginas.Name = "LLoginas";
            this.LLoginas.Size = new System.Drawing.Size(48, 13);
            this.LLoginas.TabIndex = 59;
            this.LLoginas.Text = "Login As";
            // 
            // TxtLoginname
            // 
            this.TxtLoginname.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtLoginname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtLoginname.ForeColor = System.Drawing.Color.Black;
            this.TxtLoginname.Location = new System.Drawing.Point(930, 371);
            this.TxtLoginname.Name = "TxtLoginname";
            this.TxtLoginname.ReadOnly = true;
            this.TxtLoginname.Size = new System.Drawing.Size(87, 20);
            this.TxtLoginname.TabIndex = 63;
            // 
            // TxtLoginas
            // 
            this.TxtLoginas.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtLoginas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtLoginas.ForeColor = System.Drawing.Color.Black;
            this.TxtLoginas.Location = new System.Drawing.Point(915, 392);
            this.TxtLoginas.Name = "TxtLoginas";
            this.TxtLoginas.ReadOnly = true;
            this.TxtLoginas.Size = new System.Drawing.Size(102, 20);
            this.TxtLoginas.TabIndex = 62;
            // 
            // BtnSkip
            // 
            this.BtnSkip.Location = new System.Drawing.Point(950, 525);
            this.BtnSkip.Name = "BtnSkip";
            this.BtnSkip.Size = new System.Drawing.Size(67, 21);
            this.BtnSkip.TabIndex = 52;
            this.BtnSkip.Text = "Skip";
            this.BtnSkip.UseVisualStyleBackColor = true;
            this.BtnSkip.Click += new System.EventHandler(this.BSkip_Click);
            // 
            // TxtFormid
            // 
            this.TxtFormid.AcceptsReturn = true;
            this.TxtFormid.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtFormid.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtFormid.Location = new System.Drawing.Point(930, 133);
            this.TxtFormid.Name = "TxtFormid";
            this.TxtFormid.ReadOnly = true;
            this.TxtFormid.Size = new System.Drawing.Size(87, 20);
            this.TxtFormid.TabIndex = 1;
            // 
            // LblFormid
            // 
            this.LblFormid.AutoSize = true;
            this.LblFormid.Location = new System.Drawing.Point(887, 136);
            this.LblFormid.Name = "LblFormid";
            this.LblFormid.Size = new System.Drawing.Size(42, 13);
            this.LblFormid.TabIndex = 0;
            this.LblFormid.Text = "Form Id";
            // 
            // BtnRefersh
            // 
            this.BtnRefersh.Location = new System.Drawing.Point(950, 427);
            this.BtnRefersh.Name = "BtnRefersh";
            this.BtnRefersh.Size = new System.Drawing.Size(67, 21);
            this.BtnRefersh.TabIndex = 64;
            this.BtnRefersh.Text = "Refresh";
            this.BtnRefersh.UseVisualStyleBackColor = true;
            this.BtnRefersh.Click += new System.EventHandler(this.BtnRefersh_Click);
            // 
            // BtnDeoScanMain
            // 
            this.BtnDeoScanMain.Location = new System.Drawing.Point(915, 646);
            this.BtnDeoScanMain.Name = "BtnDeoScanMain";
            this.BtnDeoScanMain.Size = new System.Drawing.Size(67, 21);
            this.BtnDeoScanMain.TabIndex = 53;
            this.BtnDeoScanMain.Text = "Main";
            this.BtnDeoScanMain.UseVisualStyleBackColor = true;
            this.BtnDeoScanMain.Click += new System.EventHandler(this.BtnDeoScanBack2main_Click);
            // 
            // LblAccno
            // 
            this.LblAccno.AutoSize = true;
            this.LblAccno.Location = new System.Drawing.Point(179, 639);
            this.LblAccno.Name = "LblAccno";
            this.LblAccno.Size = new System.Drawing.Size(92, 13);
            this.LblAccno.TabIndex = 26;
            this.LblAccno.Text = "Bank Account No";
            // 
            // TxtAccno
            // 
            this.TxtAccno.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtAccno.Location = new System.Drawing.Point(274, 635);
            this.TxtAccno.Name = "TxtAccno";
            this.TxtAccno.Size = new System.Drawing.Size(133, 20);
            this.TxtAccno.TabIndex = 27;
            this.TxtAccno.Enter += new System.EventHandler(this.TxtAccno_Enter);
            this.TxtAccno.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtAccno_KeyPress);
            // 
            // LblBankbranch
            // 
            this.LblBankbranch.AutoSize = true;
            this.LblBankbranch.Location = new System.Drawing.Point(414, 639);
            this.LblBankbranch.Name = "LblBankbranch";
            this.LblBankbranch.Size = new System.Drawing.Size(69, 13);
            this.LblBankbranch.TabIndex = 28;
            this.LblBankbranch.Text = "Bank Branch";
            // 
            // TxtBankbranch
            // 
            this.TxtBankbranch.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtBankbranch.Location = new System.Drawing.Point(485, 635);
            this.TxtBankbranch.Name = "TxtBankbranch";
            this.TxtBankbranch.Size = new System.Drawing.Size(207, 20);
            this.TxtBankbranch.TabIndex = 29;
            this.TxtBankbranch.Enter += new System.EventHandler(this.TxtBankbranch_Enter);
            this.TxtBankbranch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtBankbranch_KeyPress);
            // 
            // LblDeodd
            // 
            this.LblDeodd.AutoSize = true;
            this.LblDeodd.Location = new System.Drawing.Point(627, 591);
            this.LblDeodd.Name = "LblDeodd";
            this.LblDeodd.Size = new System.Drawing.Size(23, 13);
            this.LblDeodd.TabIndex = 9;
            this.LblDeodd.Text = "DD";
            // 
            // LblDeoMon
            // 
            this.LblDeoMon.AutoSize = true;
            this.LblDeoMon.Location = new System.Drawing.Point(695, 591);
            this.LblDeoMon.Name = "LblDeoMon";
            this.LblDeoMon.Size = new System.Drawing.Size(28, 13);
            this.LblDeoMon.TabIndex = 11;
            this.LblDeoMon.Text = "Mon";
            // 
            // LblDeoYear
            // 
            this.LblDeoYear.AutoSize = true;
            this.LblDeoYear.Location = new System.Drawing.Point(778, 591);
            this.LblDeoYear.Name = "LblDeoYear";
            this.LblDeoYear.Size = new System.Drawing.Size(29, 13);
            this.LblDeoYear.TabIndex = 13;
            this.LblDeoYear.Text = "Year";
            // 
            // LblCursorlocation
            // 
            this.LblCursorlocation.AutoSize = true;
            this.LblCursorlocation.Location = new System.Drawing.Point(769, 28);
            this.LblCursorlocation.Name = "LblCursorlocation";
            this.LblCursorlocation.Size = new System.Drawing.Size(81, 13);
            this.LblCursorlocation.TabIndex = 55;
            this.LblCursorlocation.Text = "Cursor Location";
            // 
            // DeoLblTimeinsec
            // 
            this.DeoLblTimeinsec.AutoSize = true;
            this.DeoLblTimeinsec.Location = new System.Drawing.Point(683, 568);
            this.DeoLblTimeinsec.Name = "DeoLblTimeinsec";
            this.DeoLblTimeinsec.Size = new System.Drawing.Size(176, 13);
            this.DeoLblTimeinsec.TabIndex = 56;
            this.DeoLblTimeinsec.Text = "Time to Load in Sec: And in Millsec:";
            // 
            // TxtDob
            // 
            this.TxtDob.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtDob.Location = new System.Drawing.Point(650, 587);
            this.TxtDob.Name = "TxtDob";
            this.TxtDob.Size = new System.Drawing.Size(214, 20);
            this.TxtDob.TabIndex = 15;
            this.TxtDob.Visible = false;
            this.TxtDob.Enter += new System.EventHandler(this.TxtDob_Enter);
            // 
            // TxtGender
            // 
            this.TxtGender.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtGender.Location = new System.Drawing.Point(49, 611);
            this.TxtGender.Name = "TxtGender";
            this.TxtGender.ReadOnly = true;
            this.TxtGender.Size = new System.Drawing.Size(123, 20);
            this.TxtGender.TabIndex = 17;
            this.TxtGender.Visible = false;
            this.TxtGender.Enter += new System.EventHandler(this.TxtGender_Enter);
            // 
            // TxtReligion
            // 
            this.TxtReligion.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtReligion.Location = new System.Drawing.Point(247, 611);
            this.TxtReligion.Name = "TxtReligion";
            this.TxtReligion.ReadOnly = true;
            this.TxtReligion.Size = new System.Drawing.Size(160, 20);
            this.TxtReligion.TabIndex = 19;
            this.TxtReligion.Visible = false;
            this.TxtReligion.Enter += new System.EventHandler(this.TxtReligion_Enter);
            // 
            // TxtDomicile
            // 
            this.TxtDomicile.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtDomicile.Location = new System.Drawing.Point(464, 611);
            this.TxtDomicile.Name = "TxtDomicile";
            this.TxtDomicile.ReadOnly = true;
            this.TxtDomicile.Size = new System.Drawing.Size(92, 20);
            this.TxtDomicile.TabIndex = 21;
            this.TxtDomicile.Visible = false;
            this.TxtDomicile.Enter += new System.EventHandler(this.TxtDomicile_Enter);
            // 
            // TxtEmployeetype
            // 
            this.TxtEmployeetype.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtEmployeetype.Location = new System.Drawing.Point(781, 635);
            this.TxtEmployeetype.Name = "TxtEmployeetype";
            this.TxtEmployeetype.ReadOnly = true;
            this.TxtEmployeetype.Size = new System.Drawing.Size(83, 20);
            this.TxtEmployeetype.TabIndex = 31;
            this.TxtEmployeetype.Visible = false;
            this.TxtEmployeetype.Enter += new System.EventHandler(this.TxtEmployeetype_Enter);
            // 
            // PictureBox1
            // 
            this.PictureBox1.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.PictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("PictureBox1.BackgroundImage")));
            this.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.PictureBox1.Location = new System.Drawing.Point(905, 1);
            this.PictureBox1.Name = "PictureBox1";
            this.PictureBox1.Size = new System.Drawing.Size(112, 126);
            this.PictureBox1.TabIndex = 63;
            this.PictureBox1.TabStop = false;
            // 
            // BtnEncrypt
            // 
            this.BtnEncrypt.Location = new System.Drawing.Point(870, 503);
            this.BtnEncrypt.Name = "BtnEncrypt";
            this.BtnEncrypt.Size = new System.Drawing.Size(67, 21);
            this.BtnEncrypt.TabIndex = 50;
            this.BtnEncrypt.Text = "Encrypt";
            this.BtnEncrypt.UseVisualStyleBackColor = true;
            this.BtnEncrypt.Click += new System.EventHandler(this.BtnEncrypt_Click);
            // 
            // domainUpDown1
            // 
            this.domainUpDown1.Items.Add("128");
            this.domainUpDown1.Items.Add("192");
            this.domainUpDown1.Items.Add("256");
            this.domainUpDown1.Location = new System.Drawing.Point(924, 480);
            this.domainUpDown1.Name = "domainUpDown1";
            this.domainUpDown1.ReadOnly = true;
            this.domainUpDown1.Size = new System.Drawing.Size(66, 20);
            this.domainUpDown1.TabIndex = 49;
            this.domainUpDown1.Text = "128";
            // 
            // TxtPwd
            // 
            this.TxtPwd.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtPwd.Location = new System.Drawing.Point(924, 454);
            this.TxtPwd.Name = "TxtPwd";
            this.TxtPwd.Size = new System.Drawing.Size(92, 20);
            this.TxtPwd.TabIndex = 48;
            // 
            // LblPwd
            // 
            this.LblPwd.AutoSize = true;
            this.LblPwd.Location = new System.Drawing.Point(865, 457);
            this.LblPwd.Name = "LblPwd";
            this.LblPwd.Size = new System.Drawing.Size(53, 13);
            this.LblPwd.TabIndex = 58;
            this.LblPwd.Text = "Password";
            // 
            // LblAeskey
            // 
            this.LblAeskey.AutoSize = true;
            this.LblAeskey.Location = new System.Drawing.Point(865, 482);
            this.LblAeskey.Name = "LblAeskey";
            this.LblAeskey.Size = new System.Drawing.Size(49, 13);
            this.LblAeskey.TabIndex = 57;
            this.LblAeskey.Text = "AES Key";
            // 
            // LblDegreename
            // 
            this.LblDegreename.AutoSize = true;
            this.LblDegreename.Location = new System.Drawing.Point(5, 662);
            this.LblDegreename.Name = "LblDegreename";
            this.LblDegreename.Size = new System.Drawing.Size(73, 13);
            this.LblDegreename.TabIndex = 32;
            this.LblDegreename.Text = "Degree Name";
            // 
            // TxtDegreename
            // 
            this.TxtDegreename.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtDegreename.Location = new System.Drawing.Point(86, 658);
            this.TxtDegreename.Name = "TxtDegreename";
            this.TxtDegreename.Size = new System.Drawing.Size(86, 20);
            this.TxtDegreename.TabIndex = 33;
            this.TxtDegreename.Enter += new System.EventHandler(this.TxtDegreename_Enter);
            this.TxtDegreename.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtDegreename_KeyPress);
            // 
            // LblPassyear
            // 
            this.LblPassyear.AutoSize = true;
            this.LblPassyear.Location = new System.Drawing.Point(179, 662);
            this.LblPassyear.Name = "LblPassyear";
            this.LblPassyear.Size = new System.Drawing.Size(55, 13);
            this.LblPassyear.TabIndex = 34;
            this.LblPassyear.Text = "Pass Year";
            // 
            // CmbPassyear
            // 
            this.CmbPassyear.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.CmbPassyear.FormattingEnabled = true;
            this.CmbPassyear.Items.AddRange(new object[] {
            "1970",
            "1971",
            "1972",
            "1973",
            "1974",
            "1975",
            "1976",
            "1977",
            "1978",
            "1979",
            "1980",
            "1981",
            "1982",
            "1983",
            "1984",
            "1985",
            "1986",
            "1987",
            "1988",
            "1989",
            "1990",
            "1991",
            "1992",
            "1993",
            "1994",
            "1995",
            "1996",
            "1997",
            "1998",
            "1999",
            "2001",
            "2002",
            "2003",
            "2004",
            "2005",
            "2006",
            "2007",
            "2008",
            "2009",
            "2010",
            "2011",
            "2012"});
            this.CmbPassyear.Location = new System.Drawing.Point(247, 658);
            this.CmbPassyear.Name = "CmbPassyear";
            this.CmbPassyear.Size = new System.Drawing.Size(50, 21);
            this.CmbPassyear.TabIndex = 35;
            // 
            // LblGrade
            // 
            this.LblGrade.AutoSize = true;
            this.LblGrade.Location = new System.Drawing.Point(300, 662);
            this.LblGrade.Name = "LblGrade";
            this.LblGrade.Size = new System.Drawing.Size(36, 13);
            this.LblGrade.TabIndex = 36;
            this.LblGrade.Text = "Grade";
            // 
            // TxtGrade
            // 
            this.TxtGrade.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtGrade.Location = new System.Drawing.Point(337, 658);
            this.TxtGrade.Name = "TxtGrade";
            this.TxtGrade.Size = new System.Drawing.Size(70, 20);
            this.TxtGrade.TabIndex = 37;
            this.TxtGrade.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtGrade_KeyPress);
            // 
            // LblMajor
            // 
            this.LblMajor.AutoSize = true;
            this.LblMajor.Location = new System.Drawing.Point(414, 662);
            this.LblMajor.Name = "LblMajor";
            this.LblMajor.Size = new System.Drawing.Size(33, 13);
            this.LblMajor.TabIndex = 38;
            this.LblMajor.Text = "Major";
            // 
            // TxtMajor
            // 
            this.TxtMajor.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtMajor.Location = new System.Drawing.Point(464, 658);
            this.TxtMajor.Name = "TxtMajor";
            this.TxtMajor.Size = new System.Drawing.Size(92, 20);
            this.TxtMajor.TabIndex = 39;
            this.TxtMajor.Enter += new System.EventHandler(this.TxtMajor_Enter);
            this.TxtMajor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtMajor_KeyPress);
            // 
            // LblInstitute
            // 
            this.LblInstitute.AutoSize = true;
            this.LblInstitute.Location = new System.Drawing.Point(640, 662);
            this.LblInstitute.Name = "LblInstitute";
            this.LblInstitute.Size = new System.Drawing.Size(44, 13);
            this.LblInstitute.TabIndex = 40;
            this.LblInstitute.Text = "Institute";
            // 
            // TxtInstitute
            // 
            this.TxtInstitute.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.TxtInstitute.Location = new System.Drawing.Point(684, 658);
            this.TxtInstitute.Name = "TxtInstitute";
            this.TxtInstitute.Size = new System.Drawing.Size(180, 20);
            this.TxtInstitute.TabIndex = 41;
            this.TxtInstitute.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtInstitute_KeyPress);
            // 
            // BtnTag_addition
            // 
            this.BtnTag_addition.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnTag_addition.Location = new System.Drawing.Point(2, 1);
            this.BtnTag_addition.Name = "BtnTag_addition";
            this.BtnTag_addition.Size = new System.Drawing.Size(78, 20);
            this.BtnTag_addition.TabIndex = 65;
            this.BtnTag_addition.Text = "Tag Addition";
            this.BtnTag_addition.UseVisualStyleBackColor = true;
            this.BtnTag_addition.Click += new System.EventHandler(this.BtnTag_addition_Click);
            // 
            // Dataentry
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(1028, 746);
            this.Controls.Add(this.BtnTag_addition);
            this.Controls.Add(this.LblInstitute);
            this.Controls.Add(this.TxtInstitute);
            this.Controls.Add(this.LblMajor);
            this.Controls.Add(this.TxtMajor);
            this.Controls.Add(this.LblGrade);
            this.Controls.Add(this.TxtGrade);
            this.Controls.Add(this.LblPassyear);
            this.Controls.Add(this.CmbPassyear);
            this.Controls.Add(this.LblDegreename);
            this.Controls.Add(this.TxtDegreename);
            this.Controls.Add(this.LblAeskey);
            this.Controls.Add(this.LblPwd);
            this.Controls.Add(this.TxtPwd);
            this.Controls.Add(this.domainUpDown1);
            this.Controls.Add(this.BtnEncrypt);
            this.Controls.Add(this.PictureBox1);
            this.Controls.Add(this.DeoLblTimeinsec);
            this.Controls.Add(this.LblCursorlocation);
            this.Controls.Add(this.LblDeoYear);
            this.Controls.Add(this.LblDeoMon);
            this.Controls.Add(this.LblDeodd);
            this.Controls.Add(this.LblBankbranch);
            this.Controls.Add(this.TxtBankbranch);
            this.Controls.Add(this.LblAccno);
            this.Controls.Add(this.TxtAccno);
            this.Controls.Add(this.BtnDeoScanMain);
            this.Controls.Add(this.BtnRefersh);
            this.Controls.Add(this.LblFormid);
            this.Controls.Add(this.TxtFormid);
            this.Controls.Add(this.BtnSkip);
            this.Controls.Add(this.LLoginname);
            this.Controls.Add(this.LLoginas);
            this.Controls.Add(this.TxtLoginname);
            this.Controls.Add(this.TxtLoginas);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.LblRftagno);
            this.Controls.Add(this.TxtRftagno);
            this.Controls.Add(this.LEmployeetype);
            this.Controls.Add(this.CmbMon);
            this.Controls.Add(this.LGender);
            this.Controls.Add(this.PictureBox2);
            this.Controls.Add(this.LPermaddress);
            this.Controls.Add(this.TxtPermaddress);
            this.Controls.Add(this.LTempaddress);
            this.Controls.Add(this.TxtTempaddress);
            this.Controls.Add(this.LPhoneno);
            this.Controls.Add(this.TxtPhoneno);
            this.Controls.Add(this.LReligion);
            this.Controls.Add(this.LCnic);
            this.Controls.Add(this.TxtCnic);
            this.Controls.Add(this.LEmail);
            this.Controls.Add(this.LDomicile);
            this.Controls.Add(this.Txtmail);
            this.Controls.Add(this.LblDob);
            this.Controls.Add(this.LblFathername);
            this.Controls.Add(this.LblName);
            this.Controls.Add(this.TxtFathername);
            this.Controls.Add(this.TxtName);
            this.Controls.Add(this.BtnLogout);
            this.Controls.Add(this.TxtEmployeetype);
            this.Controls.Add(this.CmbEmployeetype);
            this.Controls.Add(this.CmbDay);
            this.Controls.Add(this.TxtDomicile);
            this.Controls.Add(this.CmbDomicile);
            this.Controls.Add(this.TxtReligion);
            this.Controls.Add(this.CmbReligion);
            this.Controls.Add(this.TxtGender);
            this.Controls.Add(this.CmbGender);
            this.Controls.Add(this.CmbYear);
            this.Controls.Add(this.TxtDob);
            this.Name = "Dataentry";
            this.Text = "M.I.A Data Entry";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Dataentry_FormClosing);
            this.Load += new System.EventHandler(this.Dataentry_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label LblRftagno;
        private System.Windows.Forms.TextBox TxtRftagno;
        private System.Windows.Forms.Label LEmployeetype;
        private System.Windows.Forms.ComboBox CmbEmployeetype;
        private System.Windows.Forms.ComboBox CmbYear;
        private System.Windows.Forms.ComboBox CmbMon;
        private System.Windows.Forms.ComboBox CmbDay;
        private System.Windows.Forms.ComboBox CmbDomicile;
        private System.Windows.Forms.ComboBox CmbReligion;
        private System.Windows.Forms.Label LGender;
        private System.Windows.Forms.ComboBox CmbGender;
        private System.Windows.Forms.PictureBox PictureBox2;
        private System.Windows.Forms.Label LPermaddress;
        private System.Windows.Forms.TextBox TxtPermaddress;
        private System.Windows.Forms.Label LTempaddress;
        private System.Windows.Forms.TextBox TxtTempaddress;
        private System.Windows.Forms.Label LPhoneno;
        private System.Windows.Forms.TextBox TxtPhoneno;
        private System.Windows.Forms.Label LReligion;
        private System.Windows.Forms.Label LCnic;
        private System.Windows.Forms.TextBox TxtCnic;
        private System.Windows.Forms.Label LEmail;
        private System.Windows.Forms.Label LDomicile;
        private System.Windows.Forms.TextBox Txtmail;
        private System.Windows.Forms.Label LblDob;
        private System.Windows.Forms.Label LblFathername;
        private System.Windows.Forms.Label LblName;
        private System.Windows.Forms.TextBox TxtFathername;
        private System.Windows.Forms.TextBox TxtName;
        private System.Windows.Forms.Button BtnLogout;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.Label LLoginname;
        private System.Windows.Forms.Label LLoginas;
        private System.Windows.Forms.TextBox TxtLoginname;
        private System.Windows.Forms.TextBox TxtLoginas;
        private System.Windows.Forms.Button BtnSkip;
        private System.Windows.Forms.TextBox TxtFormid;
        private System.Windows.Forms.Label LblFormid;
        private Button BtnRefersh;
        private Button BtnDeoScanMain;
        private Label LblAccno;
        private TextBox TxtAccno;
        private Label LblBankbranch;
        private TextBox TxtBankbranch;
        private Label LblDeodd;
        private Label LblDeoMon;
        private Label LblDeoYear;
        private Label LblCursorlocation;
        private Label DeoLblTimeinsec;
        private TextBox TxtDob;
        private TextBox TxtGender;
        private TextBox TxtReligion;
        private TextBox TxtDomicile;
        private TextBox TxtEmployeetype;
        private PictureBox PictureBox1;
        private Button BtnEncrypt;
        private DomainUpDown domainUpDown1;
        private TextBox TxtPwd;
        private Label LblPwd;
        private Label LblAeskey;
        private Label LblDegreename;
        private TextBox TxtDegreename;
        private Label LblPassyear;
        private ComboBox CmbPassyear;
        private Label LblGrade;
        private TextBox TxtGrade;
        private Label LblMajor;
        private TextBox TxtMajor;
        private Label LblInstitute;
        private TextBox TxtInstitute;
        private Button BtnTag_addition;
    }
}