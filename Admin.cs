using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.IO;

namespace Deo_Module
{
    public partial class Admin : Form
    {
        public Admin()
        {
            InitializeComponent();
            //Constructor pk_ocp.
            LblNullogin.Text = Safeinfo.status;
            LblNullname.Text = Safeinfo.Empname;
            //


            ///ONLY FOR Management/Department CONTROLS
            ///I disable/enable/Hide/show controls here.
            MgmDeptGroupadddeptscale.Hide();        //It hides Mgmt_Dept_Groupof_Departments.
            MgmDeptBtnEdit.Enabled = false;
            ///

            

            ///ONLY FOR Management/Job CONTROLS
            ///I disable/enable/Hide/show controls here.
            MgmJobBtnEdit.Enabled = false;
            MgmJobTxtPermotedby.Text = Safeinfo.Empname.ToString();
            ///



            ///ONLY FOR Management/Attendance CONTROLS
            ///I disable/enable/Hide/show controls here.
            MgmAttGroupDetailview.Hide();   //Initially it hides Group_Detail_Attendance.
            ///
            ///


            ///ONLY FOR Employee/General CONTROLS
            ///I disable/enable/Hide/show controls here.
            EmpGenBtnSave.Enabled = false;   //Initially it Disable Save Button and if record found it will enable it.
            EmpGenTxtUpdateby.Text = Safeinfo.Empname.ToString();

            ///ONLY FOR Employee/Advanced CONTROLS
            ///I disable/enable/Hide/show controls here.            
        }
        
        private void BtnLogout_Click(object sender, EventArgs e)
        {
            //char charemptype = '\0';
            //if (CmbEmployeetype.Text.ToUpper() == "CONTRACT") { charemptype = 'C'; } else { charemptype = 'P'; }
            //This button is available at two TAB Control Pages. (Dept,Emp)
            this.Close();
            LoginForm newobj_form = new LoginForm();
            newobj_form.ShowDialog();
        }
        private void BtnClose_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Should I close it, Tab1?");
        }
        private void Scanner_Click(object sender, EventArgs e)
        {
            Scanner new_obj = new Scanner();
            new_obj.ShowDialog();
            this.Hide();
        }
        private void BtnEdit_Click(object sender, EventArgs e)
        {
            MgmDeptBtnEdit.Enabled = false;  //Now you can't Edit Until you Update/Cancel button.

            MgmDeptGroupadddeptscale.Show();
            MgmDeptTxtDid.ReadOnly = true;  //Now when u Update data PK will not be  editable.
            MgmDeptTxtBps.ReadOnly = true;  //Now when u Update data PK will not be  editable.
            MgmDeptTxtJobid.ReadOnly = true;//Now when u Update data PK will not be  editable.
            MgmDeptBtnSave.Text = "Update Now";
            MgmDeptBtnDelete.Show();

            if (MgmDeptRbDept.Checked == true)
            {   //End hide/show control blocks.
                mgmdept_hideshowfun(false);//Call hide+show function for Update purpose.

                if (MgmDeptGridvu.CurrentRow.Cells[0].Value.ToString() != "")
                {   //Here it choose the row, Where your current mouse is selected. and display its data for Edit.
                    string show_query = "Select * from Department Where Dept_id=" + Convert.ToInt32(MgmDeptGridvu.CurrentRow.Cells[0].Value);
                    gl_dbaseclsobj.Dml_UpdateAdapterfun(show_query);
                    //MgmDeptGridvu.CurrentRow.
                    MgmDeptTxtDid.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();
                    MgmDeptTxtDname.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][1].ToString();
                    MgmDeptTxtMgrid.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2].ToString();
                }
                else
                {
                    MessageBox.Show("Please select a row to edit.", "Edit Denied", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }//Dept ckbox=true else-if end.                
            else if (MgmDeptRbJobscales.Checked == true)
            {
                mgmdept_hideshowfun(false);//Call hide+show function for Update purpose.
                if (MgmDeptGridvu.CurrentRow.Cells[0].Value.ToString() != "")
                {   //Here it choose the row, Where your current mouse is selected. and display its data for Edit/Save.
                    string show_query = "Select * from Scale Where Bps=" + Convert.ToInt32(MgmDeptGridvu.CurrentRow.Cells[0].Value) + " And " + " Job_id='" + Convert.ToString(MgmDeptGridvu.CurrentRow.Cells[1].Value) + "'";
                    //Update_Database show_obj = new Update_Database();
                    gl_dbaseclsobj.Dml_UpdateAdapterfun(show_query);
                    MgmDeptTxtBps.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();
                    MgmDeptTxtJobid.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][1].ToString();
                    MgmDeptTxtMinbpay.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2].ToString();
                    MgmDeptTxtMaxbpay.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][3].ToString();
                    MgmDeptTxtPerIncr.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][4].ToString();
                    MgmDeptTxtHouserent.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][5].ToString();
                    MgmDeptTxtMedicalallow.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][6].ToString();
                    MgmDeptTxtConvallow.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][7].ToString();
                    MgmDeptTxtPhoneallow.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][8].ToString();
                    MgmDeptTxtLatesitallow.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][9].ToString();
                    MgmDeptTxtOtherallow.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][10].ToString();
                    MgmDeptTxtLatededuc.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][11].ToString();
                    MgmDeptTxtOtherdeduc.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][12].ToString();
                }
                else
                {
                    MessageBox.Show("Please select a row to edit.", "Edit Denied", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }//Scale ckbox=true else-if end.
        }//end fun
        private void BtnAddrow_Click(object sender, EventArgs e)
        {   //MessageBox.Show("Now You Can Add Rows.:" + MgmDeptGridvu.AllowUserToAddRows.ToString());

            MgmDeptGroupadddeptscale.Show();
            MgmDeptTxtDid.ReadOnly = false;  //Now when u insert data it will remain editable.
            MgmDeptTxtBps.ReadOnly = false;  //Now when u insert data it will remain editable.
            MgmDeptTxtJobid.ReadOnly = false;//Now when u insert data it will remain editable.
            MgmDeptBtnSave.Text = "Save";
            MgmDeptBtnDelete.Hide();

            mgmdept_hideshowfun(true);
        }
        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            if (MgmDeptRbDept.Checked == true)
            {
                MgmDeptLblDeptscaleview.Text = "Departments detailed view.";
                show_deptfun();
                MgmDeptBtnEdit.Enabled = true;

            }
            else if (MgmDeptRbJobscales.Checked == true)
            {
                MgmDeptLblDeptscaleview.Text = "Job Scale detailed view.";
                //Actually u call here scale table view.
                show_scalefun();
                MgmDeptBtnEdit.Enabled = true;
            }

        }
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Forced For integrity Press F\nForced for NUll Press N");
        }
        private void BtnShowrecord_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Select * from Advanced where status=choose.");
        }        
        private void GenBtnSave_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("Are you sure to save changes.", "Save Record", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                string query = "Select * From Employee Where Rfid_id='" + EmpGenTxtRfid.Text + "'";
                gl_dbaseclsobj.Dml_UpdateAdapterfun(query);
                gl_form_id = int.Parse(gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString());//Save for Important feature,After updation use this id and show same form that you've updated.
                //No need to updated PK. so we are not using Rows[0][0].
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][1] = EmpGentxtEname.Text.ToUpper();    //No need to update Pk.
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2] = EmpGenTxtFname.Text.ToUpper();    //No need to update Pk.
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][3] = int.Parse(EmpGenTxtCnic.Text);    //No need to update Pk.
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][4] = EmpGenTxtDob.Text;                //No need to update Pk.
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][5] = Empgentxtgender;                  //No need to update Pk.
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][6] = EmpGenTxtReligion.Text.ToUpper(); //No need to update Pk.
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][7] = EmpGenTxtDomicile.Text.ToUpper(); //No need to update Pk.                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][8] =   EmpGenTxtEmail.Text.ToUpper();

                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][9] = Convert.ToInt64(EmpGenTxtPhno.Text);
                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][10] = int.Parse(EmpGenTxtBnkacc.Text);
                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][11] = EmpGenTxtBnkbranch.Text.ToUpper();
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][12]=   Formstatus_id;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][13]=   Scanned_by;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][14]=   Deo_by;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][15]=   Hr_by;
                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][16] = Safeinfo.loginid;

                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][17] = EmpGenTxtPwd.Text.ToUpper();
                //Pls Khurram mention password lenght algo in Event handling.

                if (EmpGenRdbEnable.Checked == true)
                {
                    gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][18] = 1;
                }
                else
                {
                    gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][18] = 0;
                }
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][19] = int.Parse(EmpGenTxtRfid.Text);
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][20] = bps;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][21] = job_id;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][22] = mgr_id;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][23] = dept_id;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][24] = Hire_date;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][25] = Joining_date;
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][26] = Emp_Type;

                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][27] = EmpGenTxtTempadd.Text.ToUpper();
                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][28] = EmpGenTxtPermadd.Text.ToUpper();
                if (gl_str != null)
                {//if you've choosed an Image then execute following code and save image into dbase.
                    //gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][29] = Emp_Image;
                    //Here is Emp_image Load code. Shaa bassh. Khurram Nazir.  pk_ocp@yahoo.com
                    byte[] ImageData;
                    //MessageBox.Show("Here path of file.");
                    //Read Access to the file chosed by 'Browse' button.
                    FileStream fs = new FileStream(@gl_str, FileMode.Open, FileAccess.Read);   //ok
                    //Create a byte array of file stream length
                    ImageData = new byte[fs.Length];                                    //ok 
                    //Read block of bytes from stream into byte array
                    fs.Read(ImageData, 0, System.Convert.ToInt32(fs.Length));                  //ok
                    fs.Close();
                    //Assigning the byte array containing image data.
                    //pRow["Emp_image"] = ImageData;  //Column# 03.
                    //Employee own pic into database.

                    ////odadptr.FillSchema(urdataset, SchemaType.Source, "Employee");
                    ////For automatically generating commands to make changes to database through DataSet
                    //OracleCommandBuilder cmdbldr = new OracleCommandBuilder(odadptr);
                    //odadptr.Fill(urdataset, "Employee");
                    DataRow myrow = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0];

                    ////AddWithKey sets the Primary Key information to complete the 
                    //schema information
                    gl_dbaseclsobj.cls_odadptr.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                    myrow.BeginEdit();//Now Start edit.
                    if (ImageData.Length != 0)
                    {
                        myrow["Emp_image"] = ImageData;
                    }
                    else
                    {
                        MessageBox.Show("Image have problem.", "Image Problem Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    myrow.EndEdit();//Now end edit.                    
                }
                // gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][29] = Form_Image;
                gl_dbaseclsobj.cls_odadptr.Update(gl_dbaseclsobj.cls_dataset.Tables[0]);
                gl_str = null; //Set gl_str=null bcz on save button it contain path of *.jpeg
                MessageBox.Show("One Record Updated Successfully.", "Record Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                EmpGenBtnSave.Enabled = false;      //Now again disable it.
                search_Empfun(0, int.Parse(EmpGenTxtSearchrec.Text));//Now again Refresh its result by displaying same forms particulars.                
            }
        }
        private void GenCkformid_CheckStateChanged(object sender, EventArgs e)
        {

            if (gl_chkbox_satuts == true)
            {
                EmpGenCkbFormid.Checked = true;
                EmpGenCkbRfid.Enabled = false;
                EmpGenCkbCnic.Enabled = false;

                gl_searchfor = 0;            //Search according to form_id.
                gl_chkbox_satuts = false; //Setting for next time check.
            }
            else
            {
                EmpGenCkbFormid.Checked = false;
                EmpGenCkbRfid.Enabled = true;
                EmpGenCkbCnic.Enabled = true;

                gl_searchfor = -2;//Reset nothing,Setting for next time check.
                gl_chkbox_satuts = true;//Setting for next time check.
            }
        }
        private void GenCkRfid_CheckStateChanged(object sender, EventArgs e)
        {
            if (gl_chkbox_satuts == true)
            {
                EmpGenCkbRfid.Checked = true;
                EmpGenCkbFormid.Enabled = false;
                EmpGenCkbCnic.Enabled = false;

                gl_searchfor = 0;            //Search according to RFID_id.
                gl_chkbox_satuts = false; //Setting for next time check.
            }
            else
            {
                EmpGenCkbRfid.Checked = false;
                EmpGenCkbFormid.Enabled = true;
                EmpGenCkbCnic.Enabled = true;

                gl_searchfor = -2;//Reset nothing,Setting for next time check.
                gl_chkbox_satuts = true;
            }
        }
        private void GenCkCnic_CheckStateChanged(object sender, EventArgs e)
        {
            if (gl_chkbox_satuts == true)
            {
                EmpGenCkbCnic.Checked = true;
                EmpGenCkbFormid.Enabled = false;
                EmpGenCkbRfid.Enabled = false;

                gl_searchfor = 2;         //Search according to Cnic.
                gl_chkbox_satuts = false; //Setting for next time check.
            }
            else
            {
                EmpGenCkbCnic.Checked = false;
                EmpGenCkbFormid.Enabled = true;
                EmpGenCkbRfid.Enabled = true;

                gl_searchfor = -2;//Reset nothing,Setting for next time check.
                gl_chkbox_satuts = true;//Setting for next time check.
            }
        }
        private void GenBtnSearchrec_Click(object sender, EventArgs e)
        {
            //               1                     &&             1                        &&                 1                  &&          1
            if ((((EmpGenCkbCnic.Checked == false) && (EmpGenCkbFormid.Checked == false)) && (EmpGenCkbRfid.Checked == false)) && (EmpGenTxtSearchrec.Text == ""))
            {
                MessageBox.Show("Please choose valid search option.", "Search Option Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //ok.
            }
            //               1                       &&             1                       &&                  1                &&          1
            else if ((((EmpGenCkbCnic.Checked == false) && (EmpGenCkbFormid.Checked == false)) && (EmpGenCkbRfid.Checked == false)) && (EmpGenTxtSearchrec.Text != ""))
            {
                MessageBox.Show("Please choose valid search option.", "Search Option Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //ok.
            }
            //              1                 ||                    1               ||      1                           &&  1              
            else if ((((EmpGenCkbCnic.Checked == true) || (EmpGenCkbFormid.Checked == true)) || (EmpGenCkbRfid.Checked == true)) && (EmpGenTxtSearchrec.Text == ""))
            {
                MessageBox.Show("Please Enter Valid Information.", "Valid Information Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //ok.
            }

                //                  1                  ||         1                          ||             1                    &&            1
            else if ((((EmpGenCkbCnic.Checked == true) || (EmpGenCkbFormid.Checked == true)) || (EmpGenCkbRfid.Checked == true)) && (EmpGenTxtSearchrec.Text != ""))
            {
                switch (gl_searchfor)
                {
                    case 0:
                        {//Search against a Particular Rf Tag.
                            search_Empfun(0, int.Parse(EmpGenTxtSearchrec.Text));
                            break;
                        }
                    case 1:
                        {//Search against a Particular Form Id.
                            search_Empfun(1, int.Parse(EmpGenTxtSearchrec.Text));
                            break;
                        }                    
                    case 2:
                        {//Search against a Particular Cnic
                            search_Empfun(2, long.Parse(EmpGenTxtSearchrec.Text));
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("Enter Valid Information.");
                            break;
                        }
                }//End switch.
            }//End if-slse.                        
        }
        private void MgmDeptBtnSave_Click(object sender, EventArgs e)
        {
            if (MgmDeptBtnSave.Text == "Save")//Check yourself what is your Lable, If it is 'Save' then mean you want a New row to insert.
            {                
                    if (MgmDeptRbDept.Checked == true)//Insert record for which table? if you checked box for 'Dept' then use this one.
                    {
                        if (MgmDeptTxtDid.Text == "" || MgmDeptTxtDname.Text == ""||MgmDeptTxtMgrid.Text=="")//Dept_id or Dept_Name or Mgr_id can't be null.
                        {
                            MessageBox.Show("Please Enter valid data. Null not Allowed", "Valid data required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {//Now insert record in Department table.
                            DialogResult result0 = MessageBox.Show("Are you Sure to Permanently save changes.", "Update Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result0 == DialogResult.Yes)
                            {
                                string insert_row = "Select * from Department";
                                gl_dbaseclsobj.Dml_UpdateAdapterfun(insert_row);
                                DataRow myrow = gl_dbaseclsobj.cls_dataset.Tables[0].NewRow();
                                myrow["Dept_id"] = int.Parse(MgmDeptTxtDid.Text);
                                myrow["Dept_name"] = MgmDeptTxtDname.Text.ToUpper();
                                myrow["Mgr_id"] = int.Parse(MgmDeptTxtMgrid.Text);
                                
                                //pAdapter.Update(pDataSet, "Employee");
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Add(myrow);
                                gl_dbaseclsobj.cls_odadptr.Update(gl_dbaseclsobj.cls_dataset.Tables[0]);

                                MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                                MgmDeptGroupadddeptscale.Hide();
                                show_deptfun();
                            }
                        }
                    }//Insert into department table OFF.
                    else if (MgmDeptRbJobscales.Checked == true)
                    { //Insert record for which table, if you checked box for 'Scale' then this one.
                    
                        //MessageBox.Show("I'll run save query for Scale table.");
                        if (MgmDeptTxtBps.Text == "" || MgmDeptTxtJobid.Text == "" || MgmDeptTxtMinbpay.Text == "" || MgmDeptTxtMaxbpay.Text == "" ||( Convert.ToInt32(MgmDeptTxtMinbpay.Text) < 3000) || (Convert.ToInt32(MgmDeptTxtMaxbpay.Text) < 3000))//You can't set them Null.
                        {
                            MessageBox.Show("Valid data required.", "Fields validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {   //Now Run Insert Query for Scale table.   
                            if (verify_scalefieldfun())
                            {//Verify Field.
                                if (MgmDeptTxtPerIncr.Text == "" || MgmDeptTxtHouserent.Text == "" || MgmDeptTxtMedicalallow.Text == "" || MgmDeptTxtConvallow.Text == "" || MgmDeptTxtPhoneallow.Text == "" || MgmDeptTxtOtherallow.Text == "" || MgmDeptTxtPhoneallow.Text == "" || MgmDeptTxtLatesitallow.Text == "" || MgmDeptTxtOtherallow.Text == "" || MgmDeptTxtLatededuc.Text == "" || MgmDeptTxtOtherdeduc.Text == "")
                                {
                                    DialogResult result1 = MessageBox.Show("Do you wish default data for Null columns.", "Insertion Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (result1 == DialogResult.Yes)
                                    {
                                        try
                                        {
                                            OracleCommand cmd = new OracleCommand();
                                            gl_dbaseclsobj.cls_con.Open();
                                            cmd.Connection = gl_dbaseclsobj.cls_con;
                                            cmd.CommandType = CommandType.StoredProcedure;
                                            cmd.CommandText = "PROC_INSERT_JOB";

                                            //Sending Bps.
                                            OracleParameter p1_bps = new OracleParameter("Pv_BPS", OracleType.Number);
                                            p1_bps.Direction = ParameterDirection.Input;
                                            p1_bps.Value = int.Parse(MgmDeptTxtBps.Text.ToString());
                                            cmd.Parameters.Add(p1_bps);

                                            //Sending House Allowance True or not.
                                            OracleParameter p2_jobid = new OracleParameter("Pv_JOB_ID", OracleType.VarChar);
                                            p2_jobid.Direction = ParameterDirection.Input;
                                            p2_jobid.Value = MgmDeptTxtJobid.Text.ToUpper();
                                            cmd.Parameters.Add(p2_jobid);
                                            //Sending Convence Allowance True or not.
                                            OracleParameter p4_MIN_B_PAY = new OracleParameter("Pv_MIN_B_PAY", OracleType.Number);
                                            p4_MIN_B_PAY.Direction = ParameterDirection.Input;
                                            p4_MIN_B_PAY.Value = int.Parse(MgmDeptTxtMinbpay.Text.ToString());
                                            cmd.Parameters.Add(p4_MIN_B_PAY);

                                            //Sending Convence Allowance True or not.
                                            OracleParameter p4_MAX_B_PAY = new OracleParameter("Pv_MAX_B_PAY", OracleType.Number);
                                            p4_MAX_B_PAY.Direction = ParameterDirection.Input;
                                            p4_MAX_B_PAY.Value = int.Parse(MgmDeptTxtMaxbpay.Text.ToString());
                                            cmd.Parameters.Add(p4_MAX_B_PAY);

                                            if (MgmDeptTxtPerIncr.Text == "")
                                            {
                                                OracleParameter p4_INCR = new OracleParameter("Pv_INCR", OracleType.Number);
                                                p4_INCR.Direction = ParameterDirection.Input;
                                                p4_INCR.Value = 0;
                                                cmd.Parameters.Add(p4_INCR);
                                            }
                                            else
                                            {
                                                OracleParameter p4_INCR = new OracleParameter("Pv_INCR", OracleType.Number);
                                                p4_INCR.Direction = ParameterDirection.Input;
                                                p4_INCR.Value = int.Parse(MgmDeptTxtPerIncr.Text.ToString());
                                                cmd.Parameters.Add(p4_INCR);
                                            }
                                            if (MgmDeptTxtHouserent.Text == "")
                                            {
                                                OracleParameter p4_HOUSE_RENT = new OracleParameter("Pv_HOUSE_RENT", OracleType.Number);
                                                p4_HOUSE_RENT.Direction = ParameterDirection.Input;
                                                p4_HOUSE_RENT.Value = 0;
                                                cmd.Parameters.Add(p4_HOUSE_RENT);
                                            }
                                            else
                                            {
                                                OracleParameter p4_HOUSE_RENT = new OracleParameter("Pv_HOUSE_RENT", OracleType.Number);
                                                p4_HOUSE_RENT.Direction = ParameterDirection.Input;
                                                p4_HOUSE_RENT.Value = int.Parse(MgmDeptTxtHouserent.Text.ToString());
                                                cmd.Parameters.Add(p4_HOUSE_RENT);
                                            }
                                            if (MgmDeptTxtMedicalallow.Text == "")
                                            {
                                                OracleParameter p4_MEDICAL_ALLOW = new OracleParameter("Pv_MEDICAL_ALLOW", OracleType.Number);
                                                p4_MEDICAL_ALLOW.Direction = ParameterDirection.Input;
                                                p4_MEDICAL_ALLOW.Value = 0;
                                                cmd.Parameters.Add(p4_MEDICAL_ALLOW);
                                            }
                                            else
                                            {
                                                OracleParameter p4_MEDICAL_ALLOW = new OracleParameter("Pv_MEDICAL_ALLOW", OracleType.Number);
                                                p4_MEDICAL_ALLOW.Direction = ParameterDirection.Input;
                                                p4_MEDICAL_ALLOW.Value = int.Parse(MgmDeptTxtMedicalallow.Text.ToString());
                                                cmd.Parameters.Add(p4_MEDICAL_ALLOW);
                                            }
                                            if (MgmDeptTxtConvallow.Text == "")
                                            {
                                                OracleParameter p4_CONVE_ALLOW = new OracleParameter("Pv_CONVE_ALLOW", OracleType.Number);
                                                p4_CONVE_ALLOW.Direction = ParameterDirection.Input;
                                                p4_CONVE_ALLOW.Value = 0;
                                                cmd.Parameters.Add(p4_CONVE_ALLOW);
                                            }
                                            else
                                            {
                                                OracleParameter p4_CONVE_ALLOW = new OracleParameter("Pv_CONVE_ALLOW", OracleType.Number);
                                                p4_CONVE_ALLOW.Direction = ParameterDirection.Input;
                                                p4_CONVE_ALLOW.Value = int.Parse(MgmDeptTxtConvallow.Text.ToString());
                                                cmd.Parameters.Add(p4_CONVE_ALLOW);
                                            }

                                            if (MgmDeptTxtPhoneallow.Text == "")
                                            {
                                                OracleParameter p4_PHONE_ALLOW = new OracleParameter("Pv_PHONE_ALLOW", OracleType.Number);
                                                p4_PHONE_ALLOW.Direction = ParameterDirection.Input;
                                                p4_PHONE_ALLOW.Value = 0;
                                                cmd.Parameters.Add(p4_PHONE_ALLOW);
                                            }
                                            else
                                            {
                                                OracleParameter p4_PHONE_ALLOW = new OracleParameter("Pv_PHONE_ALLOW", OracleType.Number);
                                                p4_PHONE_ALLOW.Direction = ParameterDirection.Input;
                                                p4_PHONE_ALLOW.Value = int.Parse(MgmDeptTxtPhoneallow.Text.ToString());
                                                cmd.Parameters.Add(p4_PHONE_ALLOW);
                                            }
                                            if (MgmDeptTxtLatesitallow.Text == "")
                                            {
                                                OracleParameter p4_LATESIT_ALLOW = new OracleParameter("Pv_LATESIT_ALLOW", OracleType.Number);
                                                p4_LATESIT_ALLOW.Direction = ParameterDirection.Input;
                                                p4_LATESIT_ALLOW.Value = 0;
                                                cmd.Parameters.Add(p4_LATESIT_ALLOW);
                                            }
                                            else
                                            {
                                                OracleParameter p4_LATESIT_ALLOW = new OracleParameter("Pv_LATESIT_ALLOW", OracleType.Number);
                                                p4_LATESIT_ALLOW.Direction = ParameterDirection.Input;
                                                p4_LATESIT_ALLOW.Value = int.Parse(MgmDeptTxtLatesitallow.Text.ToString());
                                                cmd.Parameters.Add(p4_LATESIT_ALLOW);
                                            }

                                            if (MgmDeptTxtOtherallow.Text == "")
                                            {
                                                OracleParameter p4_OTHER_ALLOW = new OracleParameter("Pv_OTHER_ALLOW", OracleType.Number);
                                                p4_OTHER_ALLOW.Direction = ParameterDirection.Input;
                                                p4_OTHER_ALLOW.Value = 0;
                                                cmd.Parameters.Add(p4_OTHER_ALLOW);
                                            }
                                            else
                                            {
                                                OracleParameter p4_OTHER_ALLOW = new OracleParameter("Pv_OTHER_ALLOW", OracleType.Number);
                                                p4_OTHER_ALLOW.Direction = ParameterDirection.Input;
                                                p4_OTHER_ALLOW.Value = int.Parse(MgmDeptTxtOtherallow.Text.ToString());
                                                cmd.Parameters.Add(p4_OTHER_ALLOW);
                                            }

                                            if (MgmDeptTxtLatededuc.Text == "")
                                            {
                                                OracleParameter p4_LATE_DEDUC = new OracleParameter("Pv_LATE_DEDUC", OracleType.Number);
                                                p4_LATE_DEDUC.Direction = ParameterDirection.Input;
                                                p4_LATE_DEDUC.Value = 0;
                                                cmd.Parameters.Add(p4_LATE_DEDUC);
                                            }
                                            else
                                            {
                                                OracleParameter p4_LATE_DEDUC = new OracleParameter("Pv_LATE_DEDUC", OracleType.Number);
                                                p4_LATE_DEDUC.Direction = ParameterDirection.Input;
                                                p4_LATE_DEDUC.Value = MgmDeptTxtLatededuc.Text;
                                                cmd.Parameters.Add(p4_LATE_DEDUC);
                                            }
                                            if (MgmDeptTxtOtherdeduc.Text == "")
                                            {
                                                OracleParameter p4_OTHER_DEDUC = new OracleParameter("Pv_OTHER_DEDUC", OracleType.Number);
                                                p4_OTHER_DEDUC.Direction = ParameterDirection.Input;
                                                p4_OTHER_DEDUC.Value = 0;
                                                cmd.Parameters.Add(p4_OTHER_DEDUC);
                                            }
                                            else
                                            {
                                                OracleParameter p4_OTHER_DEDUC = new OracleParameter("Pv_OTHER_DEDUC", OracleType.Number);
                                                p4_OTHER_DEDUC.Direction = ParameterDirection.Input;
                                                p4_OTHER_DEDUC.Value = int.Parse(MgmDeptTxtOtherdeduc.Text.ToString());
                                                cmd.Parameters.Add(p4_OTHER_DEDUC);
                                            }


                                            OracleParameter p13_Answer = new OracleParameter("Pv_Answer", OracleType.VarChar, 200);
                                            p13_Answer.Direction = ParameterDirection.Output;
                                            cmd.Parameters.Add(p13_Answer);
                                            cmd.ExecuteNonQuery();
                                            String proc_output = cmd.Parameters["Pv_Answer"].Value.ToString();
                                            switch (proc_output)
                                            {
                                                case "SUCCESS":
                                                    {
                                                        MessageBox.Show("One Row successfully added.", "Record Added Succesfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                        break;
                                                    }
                                                case "LESS":
                                                    {
                                                        MessageBox.Show("Max Value can not be less thn Min Value.\nRow Insertion Failed.", "Un-Successful Addition", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        break;
                                                    }
                                            }//Off switch.
                                        }//Try off.
                                        catch (DataException ex) { ex.Message.ToString(); }
                                        finally
                                        {
                                            gl_dbaseclsobj.cls_con.Close();
                                        }
                                        MgmDeptGroupadddeptscale.Hide();
                                        show_scalefun();//Now refresh it, and show in Grid.

                                    }//Dialog confirmd.
                                }//Default value stored OFF.
                            }//Verification done.
                            else if (MgmDeptTxtPerIncr.Text != "" && MgmDeptTxtHouserent.Text != "" && MgmDeptTxtMedicalallow.Text != "" && MgmDeptTxtConvallow.Text != "" && MgmDeptTxtPhoneallow.Text != "" && MgmDeptTxtLatesitallow.Text != "" && MgmDeptTxtOtherallow.Text != "" && MgmDeptTxtLatededuc.Text != "" && MgmDeptTxtOtherdeduc.Text != "")
                            {
                                string query = "Insert into Scale Values(" + int.Parse(MgmDeptTxtBps.Text) + ",'" + MgmDeptTxtJobid.Text.ToUpper() + "'," + int.Parse(MgmDeptTxtMinbpay.Text) + "," + int.Parse(MgmDeptTxtMaxbpay.Text) + "," + int.Parse(MgmDeptTxtPerIncr.Text) + "," +
                                    +int.Parse(MgmDeptTxtHouserent.Text) + "," + int.Parse(MgmDeptTxtMedicalallow.Text) + "," + int.Parse(MgmDeptTxtConvallow.Text) + "," + int.Parse(MgmDeptTxtPhoneallow.Text) + "," + int.Parse(MgmDeptTxtLatesitallow.Text) + "," + int.Parse(MgmDeptTxtOtherallow.Text) + "," + int.Parse(MgmDeptTxtLatededuc.Text) + "," + int.Parse(MgmDeptTxtOtherdeduc.Text) + ")";
                                //MessageBox.Show(query);
                                if (gl_dbaseclsobj.Dml_Updatefun(query) == 1)
                                {
                                    MessageBox.Show("One Row successfully added.", "Record Added Succesfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    MessageBox.Show("Row Insertion Failed.", "Un-Successful Addition", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                MgmDeptGroupadddeptscale.Hide();
                                show_scalefun();//Now refresh it, and show in Grid.
                            }//Its mean nothing is null, And don't save default value. i'm giving all values.
                            else
                            {
                                MessageBox.Show("Please Verify all Fields data.", "Fill All Fields", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }//Now Run Insert Query for Scale table OFF.                        
                    }//Now Radio button for Scale table OFF.  
            }//If mgmdeptBtnsave=="Save"

            else if (MgmDeptBtnSave.Text == "Update Now")//Check yourself what is your Lable, If it is 'Update Now' then mean you want Just update a record.
            {//UPDATE CODE START HERE.
                DialogResult result = MessageBox.Show("Are you Sure to Permanently save changes.", "Update Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    if (MgmDeptRbJobscales.Checked == true)//Update record for which table, if you checked box for 'Scale' then this one.
                    {
                        
                        if(verify_scalefieldfun())
                        {//Verify Fields.
                            try
                            {
                                string search_query = "Select * from Scale Where Bps=" + Convert.ToInt32(MgmDeptTxtBps.Text) + " And " + " Job_id='" + Convert.ToString(MgmDeptTxtJobid.Text) + "'";
                                gl_dbaseclsobj.Dml_UpdateAdapterfun(search_query);
                                //gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0] = Convert.ToInt32(MgmDeptTxtBps.Text);//You can't edit pk.
                                //gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][1] = MgmDeptTxtJobid.Text.ToUpper();     //You can't edit pk.
                                if (Convert.ToInt32(MgmDeptTxtMaxbpay.Text) < Convert.ToInt32(MgmDeptTxtMinbpay.Text))
                                {
                                    MessageBox.Show("Other than Max basic Pay values are being updated.", "Save Caution", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2] = Convert.ToInt32(MgmDeptTxtMinbpay.Text);
                                }
                                else
                                {
                                    gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2] = Convert.ToInt32(MgmDeptTxtMinbpay.Text);
                                    gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][3] = Convert.ToInt32(MgmDeptTxtMaxbpay.Text);
                                }

                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][4] = Convert.ToInt32(MgmDeptTxtPerIncr.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][5] = Convert.ToInt32(MgmDeptTxtHouserent.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][6] = Convert.ToInt32(MgmDeptTxtMedicalallow.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][7] = Convert.ToInt32(MgmDeptTxtConvallow.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][8] = Convert.ToInt32(MgmDeptTxtPhoneallow.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][9] = Convert.ToInt32(MgmDeptTxtLatesitallow.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][10] = Convert.ToInt32(MgmDeptTxtOtherallow.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][11] = Convert.ToInt32(MgmDeptTxtLatededuc.Text);
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][12] = Convert.ToInt32(MgmDeptTxtOtherdeduc.Text);

                                gl_dbaseclsobj.cls_odadptr.Update(gl_dbaseclsobj.cls_dataset.Tables[0]);
                                MessageBox.Show("Record updated successfully.", "Record Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                MgmDeptGroupadddeptscale.Hide();
                                show_scalefun();//Now refresh what you updated.
                                MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                            }
                            catch (OracleException Oex)
                            {
                                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                MgmDeptGroupadddeptscale.Hide();
                                MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "General Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                                MgmDeptGroupadddeptscale.Hide();
                            }

                        }
                        
                    }
                    else if (MgmDeptRbDept.Checked == true)//Update record for which table, if you checked box for 'Dept' then this one.
                    {
                        try
                        {
                            if (MgmDeptTxtMgrid.Text == "")
                            {
                                MessageBox.Show("Please Enter Mgr_Id", "Null Not allowed", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            else
                            {
                                string search_query = "Select * from Department Where Dept_id=" + Convert.ToInt32(MgmDeptTxtDid.Text);
                                //gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0]=Convert.ToInt32(MgmDeptTxtDid.Text);//You can't edit pk.
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][1] = MgmDeptTxtDname.Text.ToUpper();
                                gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2] = Convert.ToInt32(MgmDeptTxtMgrid.Text);                                

                                gl_dbaseclsobj.cls_odadptr.Update(gl_dbaseclsobj.cls_dataset.Tables[0]);
                                //Below query will re-set mgr_id of all employee. Who's department mgr_id you've changed/Assigned.
                                string assign_mgr_query = "Update Employee Set Mgr_id=(Select Mgr_id From Department Where Dept_id=" + Convert.ToInt32(MgmDeptTxtDid.Text) + ") Where Dept_id=" + Convert.ToInt32(MgmDeptTxtDid.Text);
                                gl_dbaseclsobj.Dml_Updatefun(assign_mgr_query);
                                MessageBox.Show("One record successfully updated.", "Record Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                MgmDeptGroupadddeptscale.Hide();
                                show_deptfun();//Now refresh what you updated.
                                MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.

                            }
                        }
                        catch (OracleException Oex)
                        {
                            MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            MgmDeptGroupadddeptscale.Hide();
                            MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "General Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                            MgmDeptGroupadddeptscale.Hide();
                        }
                    }
                }//End if
            }//End Update option
        }
        private void MgmJobBtnEdit_Click(object sender, EventArgs e)
        {
            MgmJobGroupSelectedemp.Show();
            MgmJobGroupeEmpdetail.Hide();
            MgmJobGroupRestricted.Show();
            MgmJobBtnSaverec.Enabled = false;
            MgmJobTxtFormid1.Text = MgmJobGridvu.CurrentRow.Cells[0].Value.ToString(); 
            if (DateTime.Today.Day <= 3)
            {//check if today>=25 then Set Monthly settings for this employee.
                MgmJobGroupRestricted.Hide();
                MgmJobGroupeEmpdetail.Show();
                //MgmJobTxtFormid1.Text = MgmJobGridvu.CurrentRow.Cells[0].Value.ToString();                    
                search_Empfun(1, int.Parse(MgmJobTxtFormid1.Text));
                MgmJobBtnEdit.Enabled = false;      //Now i've disabled it. Bcz next time you'll not be able to press it.            
                MgmJobBtnCancel.Enabled = true;     //Now i've Enabled it. Bcz next time you'll be able to Cancel and Enable the Edit button.
                MgmJobBtnSaverec.Enabled = true; //Now i've Enabled it. Bcz You can save data.
            }
            else
            {
                MgmJobBtnSaverec.Enabled = false;
            }
        }
        private void MgmJobBtnSearch_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Cmbo contain:" + MgmJobCmb0Dept.SelectedItem);
            switch(gl_searchfor)
            {
                case -1:
                    {
                        show_Empfun();//It gives result, if (gl_searchfor=-1) for 'job_id is null'.                        
                        break;
                    }
                case 0:
                    {
                        if (MgmJobCmb0Dept.Text == "")
                        { gl_str = null; }
                        else
                        {//Here i gave name to search only this "dept's" Employees.
                            gl_str = "'" + MgmJobCmb0Dept.SelectedItem.ToString().ToUpper() + "'";
                        }
                        show_Empfun();//It gives result, if (gl_form_id=1) for department name wise.
                        break;
                    }
                    case 1:
                    {
                        if (gl_searchfor == 1 && MgmJobTxtformid0.Text == "")
                        {
                            MessageBox.Show("Please Enter Valid Form Id.", "Form Id Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            gl_form_id = int.Parse(MgmJobTxtformid0.Text);
                            show_Empfun();//It gives result, if (gl_form_id=1) for department name wise.
                        }
                        break;
                    }
            }
      
            gl_str = null;  //now set all global variable to null;
            MgmJobGroupSelectedemp.Hide();
            MgmJobCkbPermotion.Checked = false;

        }
        private void MgmJobLblJobisnull_CheckStateChanged(object sender, EventArgs e)
        {
            if (MgmJobCkbJobisnull.Checked == true)
            {
                MgmJobCkbJobisnull.Checked = true;
                MgmJobCkbdDeptwise.Enabled = false;
                MgmJobCkbFormid.Enabled = false;
                MgmJobBtnSearch.Enabled = true;     //It enable button for work.               

                gl_searchfor = -1;                   //Search for them where job is null.                                
                if (MgmJobGridvu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                {
                    MgmJobGroupSelectedemp.Hide();
                    int available_rows = MgmJobGridvu.RowCount;
                    while (available_rows > 0)
                    {
                        MgmJobGridvu.Rows.RemoveAt(available_rows - 1);
                        available_rows--;
                    }
                }
            }
            else
            {
                
                MgmJobCkbdDeptwise.Enabled = true;
                MgmJobCkbFormid.Enabled = true;
                MgmJobBtnEdit.Enabled = false;
                gl_chk_editbtn = false;

                gl_searchfor = -2;                  //gl_search has -2 mean nothing.Re-Set it for next time usage.
                MgmJobBtnSearch.Enabled = false;    //It Disable button.

            }

        }
       
        private void MgmJobLblFormid_CheckedChanged(object sender, EventArgs e)
        {


            if (MgmJobCkbFormid.Checked == true)
            {
                MgmJobGroupSelectedemp.Hide();
                int available_rows=MgmJobGridvu.RowCount;
                while(available_rows>0)                
                {
                    MgmJobGridvu.Rows.RemoveAt(available_rows-1);
                    available_rows--;
                }
                MgmJobCkbJobisnull.Enabled = false;
                MgmJobCkbdDeptwise.Enabled = false;
                MgmJobCkbFormid.Checked = true;
                MgmJobTxtformid0.Text="";
                MgmJobTxtformid0.Visible = true;       //It enable Control for work.
                //MgmJobBtnEdit.Enabled = true;

                gl_searchfor = 1;                   //Search According to provide form_id.
                MgmJobBtnSearch.Enabled = true;     //It enable Control for work.
                
           
           
            }
            else
            {
                MgmJobCkbJobisnull.Enabled = true;
                MgmJobCkbdDeptwise.Enabled = true;
                MgmJobCkbFormid.Checked = false;
                MgmJobTxtformid0.Visible = false;    //It Disable Control for work.
                MgmJobBtnEdit.Enabled = false;
                gl_chk_editbtn = false;

                gl_searchfor = -2;                  //Re-Set it for next time usage.
                MgmJobBtnSearch.Enabled = false;    //It Disable Control for work.
                
            }
        }
        private void MgmJobBtnSave_Click(object sender, EventArgs e)
        {

            //Here You insert Record for Following Tables.
            //Permotion Table.      //When Check Permotion option=True.(FOR updating.)            
            //Allo_Deduct Table.    //When Job Assigning+Hiring FIRST TIME.
            //Salary Table.         //When Job Assigning+Hiring FIRST TIME.


            if (MgmJobCkbPermotion.Checked == true)     //Its mean you want to Permote sb.
            {   //Its mean you want to Permote sb.
                //For permotion we'll store old record.
                //If date in (01-05) Thn PERMOTION ALLOWED.
                {

                    string query1 = "Select bps,job_id,dept_id from employee where form_id=" + int.Parse(MgmJobTxtFormid1.Text);
                    gl_dbaseclsobj.Dml_UpdateAdapterfun(query1);
                    string temp_bps = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();       //Fetch and store old record.
                    string temp_jobid = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][1].ToString();     //Fetch and store old record.
                    string temp_deptid = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2].ToString();    //Fetch and store old record.
                    DialogResult result;

                    result = MessageBox.Show("Are you sure to update record.", "Update Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {//HERE YOU PERFORM 02 MAIN TASKS.
                        //(01). You save old record into Permotion table.
                        //(02). You Update Employee Table with new enterd record.
                        int a, b;
                        {//(01). You save old Record into Permotion table.
                            string insert_permotion = "Insert into Permotion" +
                                " Values(" + 
                                int.Parse(MgmJobTxtFormid1.Text) + ",Trunc(sysdate,'MONTH')," + int.Parse(temp_bps) + ",'" + temp_jobid + "'," + temp_deptid+", "+Safeinfo.loginid + ",'" + MgmJobTxtComments.Text.ToUpper() + "')";
                            //MessageBox.Show(insert_permotion);
                            a = gl_dbaseclsobj.Dml_Updatefun(insert_permotion);//Here you save old record into Permotion Table.
                        }
                        { //(02). You Update Employee Table with new enterd record.

                            string update_employee = "Update Employee set Joining_date=(select Trunc(sysdate,'MONTH') from dual), Admin_by=" + Safeinfo.loginid + ", Bps=" + MgmJobCmb1Bps.SelectedItem.ToString() +
                                ", Job_id='" + MgmJobCmb1Jobid.SelectedItem.ToString() + "', Dept_id=(Select Dept_id From Department where Dept_name='" + MgmJobCmb3Dept.SelectedItem.ToString() + "') , Mgr_id=(Select Mgr_id From Department where Dept_name='" + MgmJobCmb3Dept.SelectedItem.ToString() + "') Where form_id=" + int.Parse(MgmJobTxtFormid1.Text);
                            //MessageBox.Show(update_employee);
                            b = gl_dbaseclsobj.Dml_Updatefun(update_employee);//Here you Update Employee Table. 
                        }
                        if ((a == b) && true)
                        {
                            MessageBox.Show("One record updated successfully.", "Record Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            MgmJobGroupSelectedemp.Hide();
                            MgmJobBtnSearch_Click(this, new EventArgs());
                        }
                        else
                        {
                            MessageBox.Show("Sorry! can not updated.","Fail to save",MessageBoxButtons.OK,MessageBoxIcon.Error);
                        }
                    }
                }
            }
            else
            { //Its mean you want to Update sb.
                if (gl_searchfor == 1 || gl_searchfor == 2)
                {   //By gl_searchfor variable we'll check Which type is searching you r using.
                    //Its mean you've searched record by department wise or by Form_id wise.
                    //AND MOST IMPORTANT YOU JUST WANT TO UPDATE record, U don't want TO Permote.
                    //So you've change your query that you are not changing BPS and JOB_ID.
                    //You'll just change RFID_ID,DEPARTMENT, EMP_TYPE, Joining Date. That's it.

                    //If date in (25-31) Then control enable for work.
                    {
                    char ch = '\0';
                    ch = (Convert.ToString(MgmJobCmbEmptype.SelectedItem) == "CONTRACT") ? 'C' : 'P';
                    {
                        string find_stmt = "Select Tag_id From Tag_ids Where Tag_id=" + MgmJobTxtRftag.Text;
                        Deo_Module.Update_Database chk_obj = new Update_Database();
                        chk_obj.Dml_UpdateAdapterfun(find_stmt);

                        if (chk_obj.cls_dataset.Tables[0].Rows.Count == 0)
                        {//its mean rftag id is NOT VALID. RF tagid must be same as in Database Table 'Tag_ids'.
                            MessageBox.Show("RF Tag id is not valid.\nPlease Enter Valid RF Tag id.", "Valid Tagid required", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            MgmJobTxtRftag.Text = "";
                        }
                        else
                        {//its mean rftag id is VALID. RF tagid must be same as in Database Table 'Tag_ids'.
                            string update_employee = null;
                            {
                                update_employee = "Update Employee set Formstatus_Id=6, Admin_by=" + Safeinfo.loginid + ", Rfid_id=" +
                                    MgmJobTxtRftag.Text + ", Dept_id=(Select Dept_id From Department where Dept_name='" +
                                    Convert.ToString(MgmJobCmb1Dept.SelectedItem) + "')" + ",Mgr_id=(Select Mgr_id From Department where Dept_name='" +
                                    Convert.ToString(MgmJobCmb1Dept.SelectedItem) + "'),Emp_type='" + ch + "', Joining_date=To_Date('" + MgmJobCtlHiredate.Value.Day + "-" + MgmJobCtlHiredate.Value.Month.ToString() + "-" + MgmJobCtlHiredate.Value.Year + "','dd-mm-rrrr')" +
                                    " Where Form_id=" + int.Parse(MgmJobTxtFormid1.Text);
                            }
                            //    MessageBox.Show(update_employee);
                            int ck_updated = gl_dbaseclsobj.Dml_Updatefun(update_employee);
                            if (ck_updated == 1)
                            {
                                MessageBox.Show("One record updated successfully.", "Record Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                MgmJobBtnSearch_Click(this, new EventArgs());
                                MgmJobGroupSelectedemp.Hide();
                            }
                        }//check RF tag is valid or not OFF.
                    }
                    }
                }
                else
                {   //Its mean gl_searchfor variable have '0' value.
                    //And its mean you are Assigning Bps/Job to this person first time, YOU WILL ALSO SET HIRE_DATE. YOU'LL NOT SET JOINING DATE.
                    //You'll set Joining date After giving hiredate. So here you just setting Hiredate.
                    if (Convert.ToString(MgmJobCmbEmptype.SelectedItem) == "" || Convert.ToString(MgmJobCmb0Jobid.SelectedItem) == "" || MgmJobTxtRftag.Text == "" || Convert.ToString(MgmJobCmb1Dept.SelectedItem) == "")
                    {
                        MessageBox.Show("Please fill required fields.", "Required Fields Missing", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        char temp_emptype;
                        temp_emptype = (Convert.ToString(MgmJobCmbEmptype.SelectedItem) == "CONTRACT") ? 'C' : 'P';
                        string update_employee =
                            "Update Employee set Admin_by=" + Safeinfo.loginid + ",Formstatus_id=6, Rfid_id=" + MgmJobTxtRftag.Text +
                            ", Bps=(Select Bps From Scale Where Job_id='" + MgmJobCmb0Jobid.SelectedItem.ToString() + "'),Job_id='" + MgmJobCmb0Jobid.SelectedItem.ToString() + "', Mgr_id=(Select Mgr_id From Department where Dept_name='" +
                            Convert.ToString(MgmJobCmb1Dept.SelectedItem) + "')" + ", Dept_id=(Select Dept_id From Department where Dept_name='" +
                            Convert.ToString(MgmJobCmb1Dept.SelectedItem) + "'), Emp_type='" + temp_emptype + "'" + ",Joining_date=To_Date('" + MgmJobCtlHiredate.Value.Day + "-" + MgmJobCtlHiredate.Value.Month.ToString() + "-" + MgmJobCtlHiredate.Value.Year + "','dd-mm-rrrr')" +
                            " Where Form_id=" + int.Parse(MgmJobTxtFormid1.Text);                    //MessageBox.Show(update_employee);
                        int ck_updated = gl_dbaseclsobj.Dml_Updatefun(update_employee);//Here you Update Employee Table. 
                        if (ck_updated == 1)
                        {
                            MessageBox.Show("One record updated successfully.", "Record Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            MgmJobBtnSearch_Click(this,new EventArgs());
                            MgmJobGroupSelectedemp.Hide();
                        }
                    }

                    //string check_formidquery = "Select * From Allow_Deduct Where Form_id=" + int.Parse(MgmJobTxtFormid1.Text);
                    //gl_dbaseclsobj.Dml_UpdateAdapterfun(check_formidquery);
                    ////Check is there any Row already in the table. OFFCOURSE THERE WILL BE NO ROWS (INSHALLAH).
                    //if (gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count >= 1)
                    //{//Do nothing.                    

                    //}
                    //else
                    //{//Its mean first time This person IS GOING TO BE HIRED BY this Organization, So insert his record into Allow_Deduct And Salary Table.
                    //    //string insert_allowdeduct = "Insert Into Allow_Deduct (Form_id,Month_Year) Select Form_id,Hire_date From Employee Where Form_id=" + int.Parse(MgmJobTxtFormid1.Text);
                    //    //gl_dbaseclsobj.Dml_Updatefun(insert_allowdeduct);
                    //    //insert_allowdeduct = "Insert Into Salary (Form_id,Sal_Mon) Select Form_id,Hire_date From Employee Where Form_id=" + int.Parse(MgmJobTxtFormid1.Text);
                    //    //gl_dbaseclsobj.Dml_Updatefun(insert_allowdeduct);
                    //}
                }
            }//We've updated employee record.

            /*Pls implement below exception. with save button.
             * catch (OracleException Oex)
             * {
             * MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
             * }
             * catch (Exception ex)
             * {
             * MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
             * }
             * finally
             * {
             * // MessageBox.Show("I'm going to close connection.");
             * con.Close();
             * }*/
        }//End func.                        
        private void MgmAttBtnRefresh_Click(object sender, EventArgs e)
        {
            MgmAttGroupDetailview.Hide();//Hide detail view.
            if (MgmAttGridAttendVu.RowCount >= 1)
            {//Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                for (int i = 0; i < MgmAttGridAttendVu.RowCount; i++)
                {
                    MgmAttGridAttendVu.Rows.RemoveAt(i);
                }
            }
            MgmAttCkbDept.Enabled = true;       //Enable ckbox Dept view.
            MgmAttCkbDept.Checked = false;      //checked CHKBOX of dept. 
            MgmAttCmbDept.Enabled = true;       //Enable Combo Dept view.
            MgmAttCmbDept.Text="";              //EMPTY Txtarea of Dept combo.


            MgmAttCkbRftag.Enabled = true;      //Rfid ck option Enable.
            MgmAttCkbRftag.Checked = false;     //Rfid ck option un-checkd.                
            MgmAttTxtRftag.Enabled = true;      //Rfid Txt box Enable..
            MgmAttTxtRftag.Text = "";           //EMPTY Txtbox Rfid.   

            MgmAttCkbFormid.Enabled = true;     //Enable chkbox formid.                
            MgmAttCkbFormid.Checked = false;    //Un-chkd chkbox formid.                
            MgmAttTxtFormid.Enabled = true;     //Enable Txtbox formid.  
            MgmAttTxtFormid.Text = "";          //EMPTY Txtbox formid.   
        }
        private void MgmAttGridAttendVu_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i <= gl_grid_Rows; i++)
            {
                if (MgmAttGridAttendVu.Rows[i].Cells[0].Selected == true)
                {
                    gl_form_id = int.Parse(MgmAttGridAttendVu.Rows[i].Cells[1].Value.ToString());
                    MgmAttGroupDetailview.Show();                    
                    show_Attendfun(99, "");
                }
            }


        }
        private void AdmBtnTabScanner_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Welcome Administrator for scanning Process.");            
            Scanner newobj = new Scanner();
            newobj.Show();
            this.Hide();
        }
        private void MgmDeptRbJobscales_CheckedChanged(object sender, EventArgs e)
        {
            if (MgmDeptRbDept.Checked == true)
            {
                richTextBox7.Text = "                                                   MIA DEPARTMENTS VIEW";
                MgmDeptLblDeptscaleview.Text = "Departments detailed view.";
                //Actually u call here Dept table view.
                show_deptfun();
                MgmDeptBtnEdit.Enabled = true;
            }
            else if (MgmDeptRbJobscales.Checked == true)
            {
                richTextBox7.Text = "                                                   MIA JOBS SCALE VIEW";
                MgmDeptLblDeptscaleview.Text = "Job Scale detailed view.";
                //Actually u call here scale table view.
                show_scalefun();
                MgmDeptBtnEdit.Enabled = true;
            }

        }
        private void EmpGenBtnBrowse_Click(object sender, EventArgs e)
        {
            FileDialog fileDlg = new OpenFileDialog();
            fileDlg.InitialDirectory = "C:\\Documents and Settings\\Pk_ocp\\My Documents";
            fileDlg.Filter = "Image File (*.jpg;*.bmp;*.gif)|*.jpg;*.bmp;*.gif";
            if (fileDlg.ShowDialog() == DialogResult.OK)
            {
                gl_str = fileDlg.FileName;     //Store picture's Path for when Save Button press.
                //MessageBox.Show(gl_str);
                {
                    EmpGenPicbox.Image = Image.FromFile(fileDlg.FileName);
                    EmpGenPicbox.SizeMode = PictureBoxSizeMode.StretchImage;
                }
            }
        }
        
        private void MgmJobCkbPermotion_CheckedChanged(object sender, EventArgs e)
        {
            if (MgmJobCkbPermotion.Checked == true)
            {
                MgmJobGroupRestrictedpermote.Show();
                if (DateTime.Today.Day >= 25)
                {
                    MgmJobGroupRestrictedpermote.Hide();
                    search_Empfun(1, int.Parse(MgmJobGridvu.CurrentRow.Cells[0].Value.ToString()));
                    string query0 = "Select bps from employee where form_id=" + int.Parse(MgmJobGridvu.CurrentRow.Cells[0].Value.ToString());
                    gl_dbaseclsobj.Dml_UpdateAdapterfun(query0);
                    if (gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0] == DBNull.Value)
                    {
                        MessageBox.Show("First you must assign Job and then try to Permote him.", "Permotion Denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //Hey! you can't hide it. Its already in Hiden mode. hahaha hhaha haha 
                        //MgmJobPanelPermotion.Hide();
                        MgmJobCkbPermotion.Checked = false;
                    }
                    else
                    {
                        MgmJobPanelPermotion.Show();
                        string load_Scalequery = "Select * From Scale";
                        gl_dbaseclsobj.Dml_UpdateAdapterfun(load_Scalequery);
                        for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                        {//Now fill/Populate comboboxes with scale table data.
                            MgmJobCmb1Bps.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                            MgmJobCmb1Jobid.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][1].ToString());
                        }
                        string load_deptquery = "Select Dept_Name From Department";
                        gl_dbaseclsobj.Dml_UpdateAdapterfun(load_deptquery);

                        for (int j = 0; j < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; j++)
                        {
                            MgmJobCmb3Dept.Items.Insert(j, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[j][0].ToString());
                        }
                        MgmJobBtnCancel.Enabled = true;
                        MgmJobBtnSaverec.Enabled = true;

                        //MgmJobCmbo1Dept.DataSource = gl_dbaseclsobj.cls_dataset.Tables[0];
                        //MgmJobCmbo1Dept.ValueMember = "Dept_id";
                    }
                }
            }
            else //(MgmJobCkbPermotion.Checked == false)
            {
                MgmJobPanelPermotion.Hide();
                MgmJobGroupRestrictedpermote.Hide();
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //char charemptype = '\0';
            //if (CmbEmployeetype.Text.ToUpper() == "CONTRACT") { charemptype = 'C'; } else { charemptype = 'P'; }
            //This button is available at two TAB Control Pages. (Dept,Emp)
            this.Close();
            LoginForm newobj_form = new LoginForm();
            newobj_form.ShowDialog();
        }
        private void Admin_Load(object sender, EventArgs e)
        {
            //if (TabAdmin.SelectedIndex == 0)
                if (MgmDeptRbDept.Checked == true)
                {
                    MgmDeptLblDeptscaleview.Text = "Departments detailed view.";
                    show_deptfun();
                    MgmDeptBtnEdit.Enabled = true;

                }
                else if (MgmDeptRbJobscales.Checked == true)
                {
                    MgmDeptLblDeptscaleview.Text = "Job Scale detailed view.";
                    //Actually u call here scale table view.
                    show_scalefun();
                    MgmDeptBtnEdit.Enabled = true;
                }
                if (DateTime.Today.Day >= 1 && DateTime.Today.Day <= 2)
                {
                    //MessageBox.Show("Date b/w 1..2, So I'm Insertint Sunday as Holiday for this month against all Employee.");
                    //This query will put Sunday for This monty(Sysdate) For all Employess of organization.
                    //Who have given at least Joining not just Hired.
                    string sunday_query =
                    "Declare "+
                    "V_Sysdate       Date :=Sysdate; " +
                    "V_Formid        Number; "+
                    "v_Temp          Number:=1; "+
                    "v_Nextsunday    Date; "+
                    "Cursor c1 is Select   Form_id From Employee Where Formstatus_Id=6; "+
                    "Begin "+
                    "Open c1; "+
                    "Fetch c1 Into V_FORMID; "+
                    "While(c1%FOUND) Loop "+
                    "WHILE (v_Temp=1) Loop "+ 
                    "Begin "+
                    "Select Next_day(V_sysdate,'SUNDAY') Into v_Nextsunday From Dual; "+
                    "   If ( To_char(v_Nextsunday,'MON')= To_Char(Sysdate,'MON') ) Then  "+
                    "Insert Into Attendance (Form_id,Date_Time,Status) "+
                    "Values (V_Formid,TRUNC(V_Nextsunday),'OFFICIAL HOLIDAY'); "+
                    "Commit; "+
                    //"Dbms_output.put_line('Sunday in Month of '||V_Nextsunday);"
                    "V_Sysdate:=V_nextsunday; "+
                    "Else  /*Dbms_Output.Put_Line('Sunday in NEXT Month on '||V_Nextsunday);*/ "+
                    "V_Temp:=0; "+
                    "End if; "+
                    "Exception "+
                    "When Dup_Val_On_Index Then "+
                    " /*Dbms_Output.Put_Line('Exceptin Is Running for Date '||V_Nextsunday);*/ "+
                    "Update Attendance "+
                    "Set Status='OFFICIAL HOLIDAY' "+
                    "Where Form_id=V_Formid And To_Date(To_char(DATE_TIME,'DD-MON-YY'))=Trunc(V_Nextsunday); "+
                    "Commit; "+
                    "V_Sysdate:=V_nextsunday; "+
                    "End; "+
                    "End loop; "+
                    "V_Sysdate:=Sysdate; "+
                    "Fetch c1 Into V_FORMID; "+
                    "End loop; "+
                    "Close c1; "+
                    "End;" ;

                    gl_dbaseclsobj.Dml_Updatefun(sunday_query);
                }
        }
        private void MgmDeptBtnCancel_Click(object sender, EventArgs e)
        {
            MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
            MgmDeptGroupadddeptscale.Hide();

        }
        private void MgmDeptBtnDelete_Click(object sender, EventArgs e)
        {
            if (MgmDeptLblDeptScaleadd.Text.ToUpper().StartsWith("DEPARTMENT") == true)//Check yourself what is your Lable, If it is 'Save' then mean you want a New row to insert.
            {
                DialogResult rs;
                rs = MessageBox.Show("Record will be deleted permanently.\nAre you sure?", "Deletion Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (rs == DialogResult.Yes)
                {
                    string delete_query = "Delete from Department Where Dept_id=" + Convert.ToInt32(MgmDeptGridvu.CurrentRow.Cells[0].Value);
                    if (gl_dbaseclsobj.Dml_Updatefun(delete_query) == 1)
                    {
                        MessageBox.Show("Record deleted.", "Row Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                    }
                    else 
                    {
                        MgmDeptBtnEdit.Enabled = true;                        
                    }
                    MgmDeptGroupadddeptscale.Hide();
                    show_deptfun();
                }
            }
            else if (MgmDeptLblDeptScaleadd.Text.ToUpper().StartsWith("BPS") == true)
            {
                DialogResult rs;
                rs = MessageBox.Show("Record will be deleted permanently.\nAre you sure?", "Deletion Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (rs == DialogResult.Yes)
                {
                    string delete_query = "Delete from Scale Where Bps=" + Convert.ToInt32(MgmDeptGridvu.CurrentRow.Cells[0].Value) + " And " + " Job_id='" + Convert.ToString(MgmDeptGridvu.CurrentRow.Cells[1].Value) + "'";
                    if (gl_dbaseclsobj.Dml_Updatefun(delete_query) == 1)
                    {
                        MessageBox.Show("Record deleted.", "Row Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                    }
                    else
                    {
                        MgmDeptBtnEdit.Enabled = true;  //Now you can Edit.
                    }
                    MgmDeptGroupadddeptscale.Hide();
                    show_scalefun();
                }
            }

        }
        private void MgmJobBtnCancel_Click(object sender, EventArgs e)
        {
            MgmJobGroupSelectedemp.Hide();

            MgmJobBtnEdit.Enabled = true;
            MgmJobLblPic.Enabled = true;
            //MgmJobLblPic.FindForm();
            MgmJobCkbPermotion.Checked = false;
        }
        private void EmpGenBtnEdit_Click(object sender, EventArgs e)
        {
            EmpGenLblUpdateby.Show();
            EmpGenTxtUpdateby.Show();
            EmpGenBtnEdit.Enabled = false;
            EmpGenBtnSave.Enabled = true;
            EmpGenBtnCancel.Text = "Cancel";
            EmpGenBtnCancel.Enabled = true;
            make_readonlyfun(0, false);//Now Here 0 Means that Set Readonly=true for few controls and also change color only for Emp/Gen Tab.
        }
        private void EmpGenBtnCancel_Click(object sender, EventArgs e)
        {
            if (EmpGenBtnCancel.Text == "Clear")
            {
                //clear all things.
                EmpGenBtnEdit.Enabled = false;
                EmpGenBtnSave.Enabled = false;
                clear_controlfun(0);//I'm passing 0 as Arguments. Just to tell function that i'm calling from Emp/Gen Tab page.
            }
            else
            {   //Its mean here Button is working for "Cancel" purpose.
                EmpGenLblUpdateby.Hide();
                EmpGenTxtUpdateby.Hide();
                EmpGenBtnEdit.Enabled = true;
                EmpGenBtnSave.Enabled = false;
                search_Empfun(0, int.Parse(EmpGenTxtRfid.Text)); //Now again Refresh its result by displaying same forms particulars.                

                EmpGentxtEname.ReadOnly = false;
                EmpGenTxtFname.ReadOnly = false;
                EmpGenTxtCnic.ReadOnly = false;
                EmpGenTxtDob.ReadOnly = false;
                EmpGenTxtDomicile.ReadOnly = false;
                EmpGenTxtHiredate.ReadOnly = false;
                EmpGenTxtRfid.ReadOnly = false;
                EmpGenTxtReligion.ReadOnly = false;
                EmpGenTxtLogname.ReadOnly = false;
            }
        }
        private void MgmAttLblRftag_CheckedChanged(object sender, EventArgs e)
        {
            if (MgmAttCkbRftag.Checked == true)
            {                
                MgmAttCkbDept.Checked = false;  //Un-checked of dept.
                MgmAttCkbDept.Enabled = false;  //Disable ckbox Dept view.
                MgmAttCmbDept.Enabled = false;  //Disable Combo dept view.

                MgmAttCkbFormid.Enabled = false;    //Disable chkbox formid.                
                MgmAttCkbFormid.Checked = false;    //Un-chkd chkbox formid.                
                MgmAttTxtFormid.Enabled = false;    //Disable Txtbox formid.                

                MgmAttGroupDetailview.Hide();//Hide detail view.
                if (MgmAttGridAttendVu.RowCount >= 1)
                {//Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                    for (int i = 0; i < MgmAttGridAttendVu.RowCount; i++)
                    {
                        MgmAttGridAttendVu.Rows.RemoveAt(i);
                    }
                }                
            }
            else if (MgmAttCkbRftag.Checked == false)
            {
                MgmAttCkbFormid.Enabled = true;    //Enable chkbox formid.                
                MgmAttCkbFormid.Checked = false;   //Un-chkd chkbox formid.                
                MgmAttTxtFormid.Enabled = true;    //Enable Txtbox formid.  

                MgmAttCkbRftag.Enabled = true;   //Rfid ck option Enable.
                MgmAttCkbRftag.Checked = false;  //Rfid ck option un-checkd.                
                MgmAttTxtRftag.Enabled = true;   //Rfid Txt box Enable..
                MgmAttTxtRftag.Text = "";        //EMPTY Txtbox Rfid.      

                MgmAttCkbDept.Checked = false;   //checked CHKBOX of dept. 
                MgmAttCkbDept.Enabled = true;   //Enable ckbox Dept view.
                MgmAttCmbDept.Enabled = true;   //Enable Combo Dept view.
                        
            }
        }
        private void MgmAttCkbFormid_CheckedChanged(object sender, EventArgs e)
        {
            if (MgmAttCkbFormid.Checked == true)
            {

                MgmAttCkbDept.Checked = false;  //Un-checked of dept.
                MgmAttCkbDept.Enabled = false;  //Disable ckbox Dept view.
                MgmAttCmbDept.Enabled = false;  //Disable Combo dept view.

                MgmAttCkbRftag.Enabled = false;    //Disable chkbox Rfid.                
                MgmAttCkbRftag.Checked = false;    //Un-chkd chkbox Rfid.                
                MgmAttTxtRftag.Enabled = false;    //Disable Txtbox Rfid. 

                MgmAttGroupDetailview.Hide();//Hide detail view.
                if (MgmAttGridAttendVu.RowCount >= 1)
                {//Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                    for (int i = 0; i < MgmAttGridAttendVu.RowCount; i++)
                    {
                        MgmAttGridAttendVu.Rows.RemoveAt(i);
                    }
                }
            }
            else if (MgmAttCkbFormid.Checked == false)
            {                
                MgmAttCkbDept.Enabled = true;   //Enable ckbox Dept view.
                MgmAttCkbDept.Checked = false;   //checked CHKBOX of dept. 
                MgmAttCmbDept.Enabled = true;   //Enable Combo Dept view.


                MgmAttCkbRftag.Enabled = true;   //Rfid ck option Enable.
                MgmAttCkbRftag.Checked = false;  //Rfid ck option un-checkd.                
                MgmAttTxtRftag.Enabled = true;   //Rfid Txt box Enable..

                MgmAttCkbFormid.Enabled = true;    //Enable chkbox formid.                
                MgmAttCkbFormid.Checked = false;    //Un-chkd chkbox formid.                
                MgmAttTxtFormid.Enabled = true;    //Enable Txtbox formid.  
                MgmAttTxtFormid.Text = "";         //EMPTY Txtbox formid.                 
            }
        }
        private void MgmAttBtnSearch_Click(object sender, EventArgs e)
        {

            MgmAttGroupDetailview.Hide();
            if (((MgmAttCkbDept.Checked == false) && (MgmAttCkbRftag.Checked == false)) && (MgmAttCkbFormid.Checked == false))
            {
                MessageBox.Show("Please choose an option.", "Search Option Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            if ((MgmAttCkbDept.Checked == true && Convert.ToString(MgmAttCmbDept.SelectedItem) == "") && ((MgmAttCkbRftag.Checked == false) && (MgmAttCkbFormid.Checked == false)))
            {
                //MessageBox.Show("Searching for all departments.","Search Option",MessageBoxButtons.OK,MessageBoxIcon.Information);
                MgmAttGroupDetailview.Hide();
                show_Attendfun(0, "");
            }
            else if (((MgmAttCkbDept.Checked == true) && (Convert.ToString(MgmAttCmbDept.SelectedItem) != "")) && ((MgmAttCkbRftag.Checked == false) && (MgmAttCkbFormid.Checked == false)))
            {
                show_Attendfun(1, Convert.ToString(MgmAttCmbDept.SelectedItem));
            }
            else if ((MgmAttCkbRftag.Checked == true) && ((MgmAttCkbDept.Checked == false) && (MgmAttCkbFormid.Checked == false)))
            {//If RFID presents.
                show_Attendfun(2, MgmAttTxtRftag.Text);
            }
            else if (((MgmAttCkbDept.Checked == false) && (MgmAttCkbRftag.Checked == false)) && (MgmAttCkbFormid.Checked == true))
            {//If formid presents.
                show_Attendfun(2, MgmAttTxtFormid.Text);
            }
        }
        private void MgmAttBtnClear_Click(object sender, EventArgs e)
        {
            MgmAttCkbFormid.Enabled = true;    //Enable chkbox formid.                
            MgmAttCkbFormid.Checked = false;   //Un-chkd chkbox formid.                
            MgmAttTxtFormid.Enabled = true;    //Enable Txtbox formid.  

            MgmAttCkbRftag.Enabled = true;   //Rfid ck option Enable.
            MgmAttCkbRftag.Checked = false;  //Rfid ck option un-checkd.                
            MgmAttTxtRftag.Enabled = true;   //Rfid Txt box Enable..
            MgmAttTxtRftag.Text = "";        //EMPTY Txtbox Rfid.      

            MgmAttCkbDept.Checked = false;   //checked CHKBOX of dept. 
            MgmAttCkbDept.Enabled = true;   //Enable ckbox Dept view.
            MgmAttCmbDept.Enabled = true;   //Enable Combo Dept view.

            MgmAttGroupDetailview.Hide();

            MgmAttGroupDetailview.Hide();//Hide detail view.


            if (MgmAttGridAttendVu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
            {
                int available_gridrows = MgmAttGridAttendVu.Rows.Count;
                while (available_gridrows != 0)
                {
                    MgmAttGridAttendVu.Rows.RemoveAt(--available_gridrows);

                }
            }           
        }
        private void MgmAttCmbDept_MouseDown(object sender, MouseEventArgs e)
        {
            string dept_load = "Select Dept_name from Department";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_load);

            if (MgmAttCmbDept.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmAttCmbDept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                MgmAttCmbDept.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmAttCmbDept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }

        }
        
        private void Management_Selected(object sender, TabControlEventArgs e)
        {
            //MessageBox.Show("This is slected" + Management.SelectedTab.ToString());
            if (Management.SelectedTab.Name.ToString().ToUpper() == "DEPARTMENT")
            {
                //MessageBox.Show("Department Selected.");
            }            
            if (Management.SelectedTab.Text.ToString().ToUpper() == "JOB UPDATE")
            {
                //MessageBox.Show("Job update Selected.");
                if (MgmJobCkbFormid.Checked == true || MgmJobCkbdDeptwise.Checked == true)
                {
                }
                else
                {
                    MgmJobCkbJobisnull.Checked = true;//Now set defaul is 'Job_id' is checked.
                    gl_searchfor = -1;//It gives result, if (gl_form_id=-1) for 'job_id is null'.
                    show_Empfun();
                }
            }
            if (Management.SelectedTab.Text.ToString().ToUpper() == "ATTENDANCE")
            {
                //MessageBox.Show("Attendacne Selected.");
                //MessageBox.Show("Searching for all departments");
                MgmAttCmbDept.Text = "";
                MgmAttGroupDetailview.Hide();
                show_Attendfun(0, "");
            }

            if (Management.SelectedTab.Name.ToUpper() == "TABMGMGENERAL")
            {
                MgmGenGroupLeave.Hide();
                /*string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                MessageBox.Show("ur windows login is:" + username);
                Safeinfo.purpose_deo_hr = 1;        //Setting this variable==1, Its mean my purpose is TO GIVE/SET LEAVE.
                Recep newobj = new Recep();         //Its mean Load Reception Form.
                newobj.Show();
                this.Hide();
                 * */
            }
        }        
        private void MgmBtnTabLeave_Click(object sender, EventArgs e)
        {
            string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            MessageBox.Show("ur windows login is:" + username);
            Safeinfo.purpose_deo_hr_receptionst = 1;        //Setting this variable==1, Its mean my purpose is TO GIVE/SET LEAVE.
            Recep newobj = new Recep();         //Its mean Load Reception Form.
            newobj.Show();
            this.Hide();
        }
        private void TabAdmin_Selected(object sender, TabControlEventArgs e)
        {

            if (TabAdmin.SelectedTab.Name.ToUpper() == "TABDEO")
            {
                Safeinfo.purpose_deo_hr_receptionst = 0;        //Its mean my purpose is for DEO Only.And when you load deo page then check this var and load for this purpose not for HR purpose.            
                Dataentry newobj = new Dataentry(); //Its mean Load DEO Form.
                newobj.Show();
                this.Hide();
            }
            if (TabAdmin.SelectedTab.Name.ToUpper() == "TABHR")
            {
                Safeinfo.purpose_deo_hr_receptionst = 1;        //Its mean my purpose is for HR Only.
                OA_PS_RFID.RFTag_Loader newobj = new  OA_PS_RFID.RFTag_Loader(); //Its mean Load DEO Form for HR purpose.
                newobj.Show();
                this.Hide();

                //Safeinfo.purpose_deo_hr_receptionst = 1;        //Its mean my purpose is for HR Only.
                //Dataentry newobj = new Dataentry(); //Its mean Load DEO Form for HR purpose.
                //newobj.Show();
                //this.Hide();
            }
            if (TabAdmin.SelectedTab.Name.ToUpper() == "TABSCANNER")
            {
                Safeinfo.purpose_deo_hr_receptionst = 2;        //Its mean my purpose is sCANNING BUT FROM Admin accout.
                Scanner newobj = new Scanner();
                newobj.Show();
                this.Hide();
            } 

            if (TabAdmin.SelectedTab.Name.ToUpper() == "TABRECEPTION")
            {
                Safeinfo.purpose_deo_hr_receptionst = 0;        //Setting this variable==0, Its mean my purpose is JUST VIEW RECEPTION APP.            
                OA_PS_RFID.RFReader newobj = new OA_PS_RFID.RFReader();        //Its mean Load Rf Form.
                newobj.Show();
                this.Hide();
            }           
        }        
        private void MgmJobCmb0Dept_MouseDown(object sender, MouseEventArgs e)
        {
            string dept_load = "Select Dept_name from Department";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_load);

            if (MgmJobCmb0Dept.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb0Dept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                MgmJobCmb0Dept.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb0Dept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }
        private void MgmAttCkbDept_CheckStateChanged(object sender, EventArgs e)
        {
            if (MgmAttCkbDept.Checked == true)
            {

                MgmAttCkbRftag.Checked = false;     //Rfid ck option un-checkd.
                MgmAttCkbRftag.Enabled = false;     //Rfid ck option Disable.
                MgmAttTxtRftag.Enabled = false;     //Rfid Txt box disable.

                MgmAttCkbFormid.Checked = false;    //Formid ck option un-checkd.
                MgmAttCkbFormid.Enabled = false;    //Formid ck option Disable.
                MgmAttTxtFormid.Enabled = false;    //Formid Txt box disable.


                MgmAttGroupDetailview.Hide();//Hide detail view.
                if (MgmAttGridAttendVu.RowCount >= 1)
                {//Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                    for (int i = 0; i < MgmAttGridAttendVu.RowCount; i++)
                    {
                        MgmAttGridAttendVu.Rows.RemoveAt(i);
                    }
                }
            }
            else if (MgmAttCkbDept.Checked == false)
            {
                MgmAttCkbRftag.Checked = false;  //Rfid ck option un-checkd.
                MgmAttCkbRftag.Enabled = true;   //Rfid ck option Enable.
                MgmAttTxtRftag.Enabled = true;   //Rfid Txt box Enable..

                MgmAttCkbFormid.Checked = false; //Formid ck option un-checkd.
                MgmAttCkbFormid.Enabled = true;  //Formid ck option Enable..
                MgmAttTxtFormid.Enabled = true;  //Formid Txt box Enable..
                
                MgmAttGroupDetailview.Hide();
            
            }


        }
        private void MgmJobCmb1Dept_MouseDown(object sender, MouseEventArgs e)
        {
            string dept_load = "Select Dept_Name from Department";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_load);

            if (MgmJobCmb1Dept.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb1Dept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                MgmJobCmb1Dept.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb1Dept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }
        private void MgmJobCmb0Jobid_MouseDown(object sender, MouseEventArgs e)
        {
            string Job_load = "Select Job_id from Scale";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(Job_load);

            if (MgmJobCmb0Jobid.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb0Jobid.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                MgmJobCmb0Jobid.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb0Jobid.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }
        private void MgmJobCmb0Bps_MouseDown(object sender, MouseEventArgs e)
        {
            //string Bps_load = "Select Bps from Scale";
            //gl_dbaseclsobj.Dml_UpdateAdapterfun(Bps_load);

            //if (MgmJobCmb0Bps.Items.Count <= 0)
            //{//If It is loading 1st time, or there is nothing in the combobox. then Load data.
            //    for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
            //    {
            //        MgmJobCmb0Bps.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
            //    }
            //}
            //else
            //{//If There is already data in the combobox. then release all control.
            //    //MgmJobCmb0Bps.Items.Clear();
            //    for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
            //    {
            //        MgmJobCmb0Bps.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
            //    }
            //}
        }
        private void MgmJobCmb1Bps_MouseDown(object sender, MouseEventArgs e)
        {
            string Bps_load = "Select Bps from Scale";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(Bps_load);

            if (MgmJobCmb1Bps.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb1Bps.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                MgmJobCmb1Bps.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb1Bps.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }
        private void MgmJobCmb1Jobid_MouseDown(object sender, MouseEventArgs e)
        {
            string Job_load = "Select Job_id from Scale";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(Job_load);

            if (MgmJobCmb1Jobid.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb1Jobid.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                MgmJobCmb1Jobid.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb1Jobid.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }
        private void MgmJobCmb3Dept_MouseDown(object sender, MouseEventArgs e)
        {
            string dept_load = "Select Dept_Name from Department";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_load);

            if (MgmJobCmb3Dept.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb3Dept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                MgmJobCmb3Dept.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    MgmJobCmb3Dept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }        
        private void EmpMSetCbmDepartment_MouseDown(object sender, MouseEventArgs e)
        {
            string dept_load = "Select Dept_name from Department";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_load);

            if (EmpMonCbmDepartment.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    EmpMonCbmDepartment.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                EmpMonCbmDepartment.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    EmpMonCbmDepartment.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }
        private void tabControl2_Selected(object sender, TabControlEventArgs e)
        {
            if (Management.SelectedTab.Text.ToString().ToUpper() == "ATTENDANCE")
            {
                show_Attendfun(0, "");
            }
           
        }
        private void EmpMSetBtnSearch_Click(object sender, EventArgs e)
        {
            EmpMonGroupSet0.Hide();
            if (((EmpMonCkbDepartment.Checked == false) && (EmpMonCkbByrfid.Checked == false)) && (EmpMonCkbByformid.Checked == false))
            {
                MessageBox.Show("Please choose an option.", "Search Option Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            if ((EmpMonCkbDepartment.Checked == true && Convert.ToString(EmpMonCbmDepartment.SelectedItem) == "") && ((EmpMonCkbByrfid.Checked == false) && (EmpMonCkbByformid.Checked == false)))
            {
                show_monthlysettingfun(0, "");
            }
            else if (((EmpMonCkbDepartment.Checked == true) && (Convert.ToString(EmpMonCbmDepartment.SelectedItem) != "")) && ((EmpMonCkbByrfid.Checked == false) && (EmpMonCkbByformid.Checked == false)))
            {
                show_monthlysettingfun(1, Convert.ToString(EmpMonCbmDepartment.SelectedItem));
            }
            else if ((EmpMonCkbByrfid.Checked == true) && ((EmpMonCkbDepartment.Checked == false) && (EmpMonCkbByformid.Checked == false)))
            {
                show_monthlysettingfun(2, EmpMonTxtRfid.Text);
            }
            else if (((EmpMonCkbDepartment.Checked == false) && (EmpMonCkbByrfid.Checked == false)) && (EmpMonCkbByformid.Checked == true))
            {
                show_monthlysettingfun(3, EmpMonTxtFormid.Text);
            }
        }
        private void EmpMSetCkbDepartment_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpMonCkbDepartment.Checked == true)
            {
                EmpMonGroupSet0.Hide();
                if (EmpMonGridvu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                {
                    for (int i = 0; i < EmpMonGridvu.RowCount; i++)
                    {
                        EmpMonGridvu.Rows.RemoveAt(i);
                    }
                }
            }
            else
            {
                EmpMonCbmDepartment.Text = "";
                EmpMonGroupSet0.Hide();
            }

        }
        private void EmpMSetCkbByformid_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpMonCkbByformid.Checked == true)
            {
                EmpMonGroupSet0.Hide();
                EmpMonTxtFormid.Show();               //Show txt for formid.
                EmpMonTxtFormid.Text="";
                EmpMonCkbDepartment.Checked = false;  //Un-chked for Dept view.
                EmpMonCkbDepartment.Enabled = false;  //Disable for chked dept view.
                EmpMonCbmDepartment.Enabled = false;  //Disable Combo dept view.
                EmpMonCkbByrfid.Enabled = false; //Disable ckbox for rftag.
                if (EmpMonGridvu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
                {
                    int available_gridrows = EmpMonGridvu.Rows.Count;
                    while (available_gridrows != 0)
                    {
                        EmpMonGridvu.Rows.RemoveAt(--available_gridrows);

                    }
                }
            }
            else
            {
                EmpMonTxtRfid.Hide();          //Hide txt for Rftag.
                EmpMonCkbDepartment.Checked = true;   //Chked for Dept view.
                EmpMonCkbDepartment.Enabled = true;   //Enable for chked dept view.
                EmpMonCbmDepartment.Enabled = true;   //Enable Combo dept view.

                EmpMonCkbByformid.Show();         //Hide txt for formid.
                EmpMonTxtFormid.Hide();               //Show txt for formid.
                EmpMonCkbByrfid.Enabled = true;  //Disable ckbox for rftag.
                EmpMonGroupSet0.Hide();
            }
        }
        private void EmpMSetCkbByrfid_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpMonCkbByrfid.Checked == true)
            {
                EmpMonGroupSet0.Hide();
                EmpMonTxtRfid.Show();          //Show rf Txt.
                EmpMonTxtRfid.Text = "";
                EmpMonCkbDepartment.Checked = false;  //Un-checked of dept.
                EmpMonCkbDepartment.Enabled = false;  //Disable ckbox Dept view.
                EmpMonCbmDepartment.Enabled = false;  //Disable Combo dept view.
                EmpMonCkbByformid.Enabled = false;    //Disable chkbox formid.   
                

                if (EmpMonGridvu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
                {
                    int available_gridrows = EmpMonGridvu.Rows.Count;
                    while (available_gridrows != 0)
                    {
                        EmpMonGridvu.Rows.RemoveAt(--available_gridrows);

                    }
                }                
            }
            else
            {
                EmpMonTxtRfid.Hide();          //Hide rf Txt.
                EmpMonCkbDepartment.Checked = true;   //checked CHKBOX of dept. 
                EmpMonCkbDepartment.Enabled = true;   //Enable ckbox Dept view.
                EmpMonCbmDepartment.Enabled = true;   //Enable Combo Dept view.
                EmpMonCkbByformid.Enabled = true;     //Enable chkbox formid.            
                EmpMonGroupSet0.Hide();
            }
        }
        private void EmpMSetBtnCancel_Click(object sender, EventArgs e)
        {
            if (EmpMonCkbDepartment.Checked == true)
            {
                EmpMonCbmDepartment.Text = ""; 
            }
            if (EmpMonCkbByrfid.Checked == true)
            {
                EmpMonTxtRfid.Text = "";            
            }
            if (EmpMonCkbByformid.Checked == true)
            {
                EmpMonTxtFormid.Text = "";            
            }
            EmpMonGroup2.Hide();
            EmpMonGroupSet0.Hide();


            if (EmpMonGridvu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
            {
                int available_gridrows=EmpMonGridvu.Rows.Count;
                while (available_gridrows != 0)
                {
                    EmpMonGridvu.Rows.RemoveAt(--available_gridrows);
                    
                }                
            }
        }
        private void EmpMSetGridvu_CellClick(object sender, DataGridViewCellEventArgs e)
        {            
            if (DateTime.Today.Day >= 23)
            {//check if today>=25 then Set Monthly settings for this employee.
                for (int i = 0; i <= gl_grid_Rows; i++)
                {
                    if (EmpMonGridvu.Rows[i].Cells[0].Selected == true)
                    {//If any Row selected then enable other controls.
                        EmpMonGroupSet0.Show();
                        gl_form_id = int.Parse(EmpMonGridvu.Rows[i].Cells[1].Value.ToString());
                        gl_dbaseclsobj.Dml_UpdateAdapterfun("Select Ename,Bps,Job_id From Employee Where Form_id=" + gl_form_id);
                        EmpMonTxtName.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();
                        EmpMonTxtBps.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][1].ToString();
                        EmpMonTxtJobid.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][2].ToString();
                        EmpMonTxtName.BackColor = Color.Linen;
                        EmpMonTxtBps.BackColor = Color.Linen;
                        EmpMonTxtJobid.BackColor = Color.Linen;
                        EmpMonTxtAllowtotal.BackColor = Color.Linen;
                        EmpMonTxtDedtotal.BackColor = Color.Linen;
                        EmpMonTxtTotalamount.BackColor = Color.Linen;
                        break;

                        //MgmAttGroupDetailview.Show();
                        //show_Attendfun(99, "");
                        //MgmAttLblDate.Text = MgmAttLblDate.Text + MgmAttDTP.Value.Day + "-" + MgmAttDTP.Value.Month + "-" + MgmAttDTP.Value.Year;
                    }
                }
            }
            else
            {//If today<25 then DON'T DISPLYAY Monthly settings OPTIONS.
                EmpMonGroup2.Show();
            }
        }
        private void EmpMsetCkbConvAllowance_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void EmpMSetBtnCancel1_Click(object sender, EventArgs e)
        {
            EmpMonGroupSet0.Hide();
        }
        private void EmpMsetCkbHouserent_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpMonCkbHouserent.Checked == true)
            {
                //EmpMsetCkbHouserent.Font.Bold = true;
            }
            else
            {
                //EmpMsetCkbHouserent.Font.Bold = true;
            }
        }
        private void MgmDeptTxtDname_Leave(object sender, EventArgs e)
        {
            Safeinfo.charonly_fun(MgmDeptTxtDname.Text);
            if (MgmDeptTxtDname.Text.Length < 2)
                MgmDeptTxtDname.Focus();
        }                
        private void EmpAttGridvu_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < EmpAttGridvu.Rows.Count; i++)
            {
                if (EmpAttGridvu.Rows[i].Cells[0].Selected == true)
                {
                    EmpAttGroup1.Show();
                    string disdate = EmpAttGridvu.Rows[i].Cells[3].Value.ToString();
                    int formid = int.Parse(EmpAttGridvu.Rows[i].Cells[1].Value.ToString());
                    EmpAttGridvudetail.Show();
                    gl_dbaseclsobj.Grid_Loaderfun(EmpAttGridvudetail, "Select Form_id,To_char(Date_Time)Date_Time, Status, NVL(HH,0)||':'||NVL(MM,0)||':'||NVL(SS,0) Spent_Time  From Attendance Where Form_Id=" + formid + " And Date_Time like '" + disdate + "%'");                   
                }
            }
        }
        private void EmpAttCkbSearchbyrfid_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpAttCkbSearchbyrfid.Checked == true)
            {
                EmpAttCkbSearchbyformid.Checked = false;    //Un-checked of formid option.
                EmpAttCkbSearchbyformid.Enabled = false;    //Disable ckbox  of formid option.                
                EmpAttTxtSearchbyformid.Enabled = false;    //Disable txtbox  of formid.                
                EmpAttGroup1.Hide();

                if (EmpAttGridvu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                {
                    for (int i = 0; i < EmpAttGridvu.RowCount; i++)
                    {
                        EmpAttGridvu.Rows.RemoveAt(i);
                    }
                }
                EmpAttPicBox.Image = null;
            }
//            else if (EmpAttCkbSearchbyformid.Checked == true)
//            {
//                EmpAttTxtSearchbyrfid.Enabled = false;      //Disable ckbox  of rfid option.
//                EmpAttCkbSearchbyrfid.Checked = false;      //Un-Checked rfid option.

//                EmpAttCkbSearchbyformid.Enabled = true;     //Enable ckbox  of formid option.

////                EmpAttCkbSearchbyformid.Checked = false;    //Un-checked of formid option.                
//                EmpAttGroup1.Hide();
//                if (EmpAttGridvu.RowCount >= 1)
//                {//if grid contain any data then remove it.
//                    for (int i = 0; i <= EmpAttGridvu.RowCount; i++)
//                    {
//                        EmpAttGridvu.Rows.RemoveAt(0);
//                    }
//                }
//            }
            else if (EmpAttCkbSearchbyformid.Checked == false)
            {
                EmpAttCkbSearchbyrfid.Checked = false;      //Un-Checked of Rfid Option.
                EmpAttCkbSearchbyrfid.Enabled = true;       //Enable Checked of Rfid Option.
                EmpAttTxtSearchbyrfid.Enabled = true;       //Enable txtbox  of Rfid.            
                EmpAttTxtSearchbyrfid.Text = "";            //Empty rfid txt;

                EmpAttCkbSearchbyformid.Checked = false;    //Un-Checked of Formid option.
                EmpAttCkbSearchbyformid.Enabled = true;     //Enable Checked of Formid Option.
                EmpAttTxtSearchbyformid.Enabled = true;     //Enable txtbox  of formid.
                EmpAttTxtSearchbyformid.Text = "";          //Empty Formid txt.

                EmpAttGroup1.Hide();
                if (EmpAttGridvu.RowCount >= 1)
                {//if grid contain any data then remove it.
                    for (int i = 0; i <= EmpAttGridvu.RowCount; i++)
                    {
                        EmpAttGridvu.Rows.RemoveAt(0);
                    }
                }
            
            }
        }                   
        private void EmpAttLblSearchbyformid_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpAttCkbSearchbyformid.Checked == true)
            {
                EmpAttCkbSearchbyrfid.Checked = false;       //Un-Checked Rfid option.
                EmpAttCkbSearchbyrfid.Enabled = false;       //Disable Rfid option.
                EmpAttTxtSearchbyrfid.Enabled = false;       //Disable Rfid Txt.
                EmpAttGroup1.Hide();
                if (EmpAttGridvu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                {
                    for (int i = 0; i < EmpAttGridvu.RowCount; i++)
                    {
                        EmpAttGridvu.Rows.RemoveAt(i);
                    }
                }
                EmpAttPicBox.Image = null;
            }           
             else if (EmpAttCkbSearchbyformid.Checked == false)
            {
                EmpAttCkbSearchbyrfid.Checked = false;      //Un-Checked of Rfid Option.
                EmpAttCkbSearchbyrfid.Enabled = true;       //Enable Checked of Rfid Option.
                EmpAttTxtSearchbyrfid.Enabled = true;       //Enable txtbox  of Rfid.            
                EmpAttTxtSearchbyrfid.Text = "";            //Empty rfid txt;

                EmpAttCkbSearchbyformid.Checked = false;    //Un-Checked of Formid option.
                EmpAttCkbSearchbyformid.Enabled = true;     //Enable Checked of Formid Option.
                EmpAttTxtSearchbyformid.Enabled = true;     //Enable txtbox  of formid.
                EmpAttTxtSearchbyformid.Text = "";          //Empty Formid txt.

                EmpAttGroup1.Hide();
                if (EmpAttGridvu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                {
                    for (int i = 0; i < EmpAttGridvu.RowCount; i++)
                    {
                        EmpAttGridvu.Rows.RemoveAt(i);
                    }
                }
                EmpAttPicBox.Image = null;
            
            }

        }        
        private void EmAttBtnSearch_Click_1(object sender, EventArgs e)
        {
            if ((EmpAttCkbSearchbyformid.Checked == false) && (EmpAttCkbSearchbyrfid.Checked == false))
            {
                MessageBox.Show("Please choose an option.", "Search Option Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {                
                if (EmpAttCkbSearchbyrfid.Checked == true)
                {//Search by rfid.
                    if (EmpAttTxtSearchbyrfid.Text == "") { MessageBox.Show("Please Enter RFID.", "Valid RFID Require", MessageBoxButtons.OK, MessageBoxIcon.Stop); }
                    else
                    {
                        show_EmpAttfun(EmpAttTxtSearchbyrfid.Text);
                        EmpAttGroup1.Hide();
                    }
                }
                else if (EmpAttCkbSearchbyformid.Checked == true)
                {//Search by formid.
                    if (EmpAttTxtSearchbyformid.Text == "") { MessageBox.Show("Please Enter From ID.", "Valid FORMID Require", MessageBoxButtons.OK, MessageBoxIcon.Stop); }
                    else
                    {
                        show_EmpAttfun(EmpAttTxtSearchbyformid.Text);
                        EmpAttGroup1.Hide();
                    }
                }
            }
        }
        private void EmpAttBtnClear_Click(object sender, EventArgs e)
        {
            if (EmpAttGridvu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
            {
                int available_gridrows = EmpAttGridvu.Rows.Count;
                while (available_gridrows != 0)
                {
                    EmpAttGridvu.Rows.RemoveAt(--available_gridrows);

                }
            }


            
            EmpAttPicBox.Image = null;
            EmpAttCkbSearchbyrfid.Checked = false;      //Un-Checked of Rfid Option.
            EmpAttCkbSearchbyrfid.Enabled = true;       //Enable Checked of Rfid Option.
            EmpAttTxtSearchbyrfid.Enabled = true;       //Enable txtbox  of Rfid.            
            EmpAttTxtSearchbyrfid.Text = "";            //Empty rfid txt;

            EmpAttCkbSearchbyformid.Checked = false;    //Un-Checked of Formid option.
            EmpAttCkbSearchbyformid.Enabled = true;     //Enable Checked of Formid Option.
            EmpAttTxtSearchbyformid.Enabled = true;     //Enable txtbox  of formid.
            EmpAttTxtSearchbyformid.Text = "";          //Empty Formid txt.

            EmpAttGroup1.Hide();
        }
        private void EmpMonBtnCancel1_Click(object sender, EventArgs e)
        {
            EmpMonGroup2.Hide();
        }        
        private void EmpMonBtnApply_Click(object sender, EventArgs e)
        {
            if (gl_form_id <= 0)
            {
                MessageBox.Show("Please choose a Row from upper Grid.", "Select Rows", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                monthly_setfun(gl_form_id);
                //EmpMonGroupSet0.Hide();
                gl_form_id = 0;
            }
        }        
        private void EmpSalCkbDept_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpSalCkbDept.Checked == true)
            {
                if (EmpSalGridvu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                {
                    for (int i = 0; i < EmpSalGridvu.RowCount; i++)
                    {
                        EmpSalGridvu.Rows.RemoveAt(i);
                    }
                }
                EmpSalTxtRftag.Text = "";
                EmpSalTxtRftag.Enabled = false;
                EmpSalTxtFormid.Text = "";
                EmpSalTxtFormid.Enabled = false;
                EmpSalPicbox.Image = null;
            }
            else
            {
                EmpSalCombDept.Text = "";
                EmpSalTxtRftag.Enabled = true;
                EmpSalTxtFormid.Enabled = true;

                EmpSalGroupSalarydetail.Hide();
                EmpSalGroupdaysdetail.Hide();
                EmpSalGroupAlloDeductdetail.Hide();  
            }
        }
        private void EmpSalCkbRftag_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpSalCkbRftag.Checked == true)
            {

                EmpSalCombDept.Text = "";
                EmpSalCkbDept.Enabled = false;
                EmpSalCombDept.Enabled = false;
                EmpSalCkbDept.Checked = false;  //Un-checked of dept.

                EmpSalTxtFormid.Text = "";
                EmpSalCkbFormid.Enabled = false;
                EmpSalTxtFormid.Enabled = false;
                EmpSalCkbFormid.Checked = false;  //Un-checked of Formid.
                                
                EmpSalTxtRftag.Text="";
                EmpSalTxtRftag.Enabled = true;

                EmpSalPicbox.Image = null;
                if (EmpSalGridvu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
                {
                    int available_gridrows = EmpSalGridvu.Rows.Count;
                    while (available_gridrows != 0)
                    {
                        EmpSalGridvu.Rows.RemoveAt(--available_gridrows);

                    }
                }
                
               
                
            }
            else
            {
                EmpSalTxtRftag.Text = "";

                EmpSalCombDept.Text = "";
                EmpSalCkbDept.Enabled = true;
                EmpSalCombDept.Enabled = true;                

                EmpSalTxtFormid.Text = "";
                EmpSalCkbFormid.Enabled = true;
                EmpSalTxtFormid.Enabled = true;

                EmpSalGroupSalarydetail.Hide();
                EmpSalGroupdaysdetail.Hide();
                EmpSalGroupAlloDeductdetail.Hide();  
            }
        }
        private void EmpSalCkbFormid_CheckStateChanged(object sender, EventArgs e)
        {
            if (EmpSalCkbFormid.Checked == true)
            {
                EmpSalCombDept.Text = "";
                EmpSalCkbDept.Enabled = false;
                EmpSalCombDept.Enabled = false;
                EmpSalCkbDept.Checked = false;  //Un-checked of dept.

                EmpSalTxtRftag.Text = "";
                EmpSalCkbRftag.Enabled = false;
                EmpSalTxtRftag.Enabled = false;
                EmpSalCkbRftag.Checked = false;  //Un-checked of Formid.

                EmpSalTxtFormid.Text = "";
                EmpSalTxtFormid.Enabled = true;
                if (EmpSalGridvu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
                {
                    int available_gridrows = EmpSalGridvu.Rows.Count;
                    while (available_gridrows != 0)
                    {
                        EmpSalGridvu.Rows.RemoveAt(--available_gridrows);

                    }
                }
                EmpSalPicbox.Image = null;
            }
            else
            {
                EmpSalTxtFormid.Text = "";

                EmpSalCombDept.Text = "";
                EmpSalCkbDept.Enabled = true;
                EmpSalCombDept.Enabled = true;

                EmpSalTxtRftag.Text = "";
                EmpSalCkbRftag.Enabled = true;
                EmpSalTxtRftag.Enabled = true;

                EmpSalGroupSalarydetail.Hide();
                EmpSalGroupdaysdetail.Hide();
                EmpSalGroupAlloDeductdetail.Hide();  
            }
        }
        private void EmpSalBtnCancel0_Click(object sender, EventArgs e)
        {
            EmpSalTxtName.Text = "";
            EmpSalCombDept.Text = ""; EmpSalCombDept.Enabled = true; EmpSalCkbDept.Checked = true;
            EmpSalTxtRftag.Text = ""; EmpSalTxtRftag.Enabled = false; EmpSalCkbRftag.Checked = false;
            EmpSalTxtFormid.Text = ""; EmpSalTxtFormid.Enabled = false; EmpSalCkbFormid.Checked = false;            
            EmpSalPicbox.Image = null;


            if (EmpSalGridvu.RowCount >= 1) //REMOVE PREVIOUS/Already Displayed DATA.
            {
                int available_gridrows = EmpSalGridvu.Rows.Count;
                while (available_gridrows != 0)
                {
                    EmpSalGridvu.Rows.RemoveAt(--available_gridrows);

                }
            }
                        
            EmpSalGroupSalarydetail.Hide();
            EmpSalGroupdaysdetail.Hide();
            EmpSalGroupAlloDeductdetail.Hide();  
        }
        private void EmpSalCombDept_MouseDown(object sender, MouseEventArgs e)
        {
            string dept_load = "Select Dept_name from Department";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_load);

            if (EmpSalCombDept.Items.Count <= 0)
            {//If It is loading 1st time, or there is nothing in the combobox. then Load data.
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    EmpSalCombDept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
            else
            {//If There is already data in the combobox. then release all control.
                EmpSalCombDept.Items.Clear();
                for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                {
                    EmpSalCombDept.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                }
            }
        }
        private void EmpSalBtnSearch_Click(object sender, EventArgs e)
        {
            EmpSalGroupSalarydetail.Hide();
            EmpSalGroupdaysdetail.Hide();
            EmpSalGroupAlloDeductdetail.Hide();
            EmpSalPicbox.Image = null;
            EmpSalTxtName.Text = "";


             if (((EmpSalCkbDept.Checked == false) && (EmpSalCkbRftag.Checked == false)) && (EmpSalCkbFormid.Checked == false))
            {
                MessageBox.Show("Please choose an option.", "Search Option Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            if ((EmpSalCkbDept.Checked == true && Convert.ToString(EmpSalCombDept.SelectedItem) == "") && ((EmpSalCkbRftag.Checked == false) && (EmpSalCkbFormid.Checked == false)))
            {                
                show_salaryfun(0, "",true);//Load Data into Grid For All Department.
            }
            else if (((EmpSalCkbDept.Checked == true) && (Convert.ToString(EmpSalCombDept.SelectedItem) != "")) && ((EmpSalCkbRftag.Checked == false) && (EmpSalCkbFormid.Checked == false)))
            {
                show_salaryfun(1, Convert.ToString(EmpSalCombDept.SelectedItem), true);//Load Data into Grid For Selected Department.
            }
            else if ((EmpSalCkbRftag.Checked == true) && ((EmpSalCkbDept.Checked == false) && (EmpSalCkbFormid.Checked == false)))
            {
                show_salaryfun(2, EmpSalTxtRftag.Text, true);//Load Data into Grid For Provided RfTage id.
            }
            else if (((EmpSalCkbDept.Checked == false) && (EmpSalCkbRftag.Checked == false)) && (EmpSalCkbFormid.Checked == true))
            {
                show_salaryfun(3, EmpSalTxtFormid.Text, true);//Load Data into Grid For Provided Formid.
            }
        }
        private void EmpSalGridvu_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            
            //for (int i = 0; i <= gl_grid_Rows; i++)
            //{
            //    if (EmpSalGridvu.Rows[i].Cells[0].Selected == true)
            //    {
            //        EmpSalGroupSalarydetail.Show();
            //        EmpSalGroupdaysdetail.Show();

            //        string strformid=EmpSalGridvu.Rows[i].Cells[1].Value.ToString();
            //        show_salaryfun(0, strformid, false);//False mean, I don't want to Load Data into Grid, Rather show Attendance.
            //        run_procefun(strformid);

            //        gl_dbaseclsobj.Show_Datafun("Select Ename,Emp_Image From Employee Where Form_Id="+strformid, 0);
            //        if (gl_dbaseclsobj.cls_dr.Read())
            //        {
            //            EmpSalTxtName.Text = gl_dbaseclsobj.cls_dr.GetString(0).ToUpper();
            //            //Picture loading code.
            //            OracleLob blob = gl_dbaseclsobj.cls_dr.GetOracleLob(1);      //Column # .      Emp_image
            //            Byte[] BLOBData = new Byte[blob.Length];
            //            //Read blob data into byte array
            //            int j = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));
            //            //Get the primitive byte data into in-memory data stream
            //            MemoryStream stmBLOBData = new MemoryStream(BLOBData);
            //            //LOADING INTO PICBOX1                            
            //            EmpSalPicbox.Image = Image.FromStream(stmBLOBData);
            //            //MessageBox.Show("Now i'm reseting it.");
            //            EmpSalPicbox.SizeMode = PictureBoxSizeMode.StretchImage;
            //        }
            //        gl_dbaseclsobj.Show_Datafun("Select Ename,Emp_Image From Employee Where Form_Id=" + strformid, 1);
            //        //gl_dbaseclsobj.Show_Datafun("Select Ename, Emp_Image From Employee Where Form_Id=" + 2840, 1);
                                        
            //        EmpSalCkbAllowamount.Checked = false;
            //        EmpSalCkbDeducamount.Checked = false;

            //    }
            //}
            

        }
        private void EmpSalCkbAllowamount_CheckStateChanged(object sender, EventArgs e)
        {

            if (EmpSalCkbAllowamount.Checked == true)
            {
                EmpSalCkbDeducamount.Checked = false;
                EmpSalGroupAlloDeductdetail.Show();

                EmpSalLblHouseallow.Show();     EmpSalTxtHouseallow.Show();
                EmpSalLblMedicalallow.Show();   EmpSalTxtMedicalallow.Show();
                EmpSalLblConvallow.Show();      EmpSalTxtConveyanceallow.Show();
                EmpSalLblPhoneallow.Show();         EmpSalTxtPhoneallow.Show();
                EmpSalLblLatesittingallow.Show();   EmpSalTxtLatesittingallow.Show();
                EmpSalLblOtherallow.Show();         EmpSalTxtOtherallow.Show();
                
                EmpSalLblLatededuc.Hide(); EmpSalTxtLatededuc.Hide();
                EmpSalLblAbsentdeduc.Hide(); EmpSalTxtAdvadeduc.Hide();
                EmpSalLblOtherdeduc.Hide(); EmpSalTxtOtherdeduc.Hide();

            }
            else
            {
                EmpSalGroupAlloDeductdetail.Hide();
            }
        }
        private void EmpSalCkbDeducamount_CheckStateChanged(object sender, EventArgs e)
        {
            

            if (EmpSalCkbDeducamount.Checked == true)
            {
                EmpSalCkbAllowamount.Checked = false;
                EmpSalGroupAlloDeductdetail.Show();

                EmpSalLblLatededuc.Show(); EmpSalTxtLatededuc.Show();
                EmpSalLblAbsentdeduc.Show(); EmpSalTxtAdvadeduc.Show();
                EmpSalLblOtherdeduc.Show(); EmpSalTxtOtherdeduc.Show();

                EmpSalLblHouseallow.Hide(); EmpSalTxtHouseallow.Hide();
                EmpSalLblMedicalallow.Hide(); EmpSalTxtMedicalallow.Hide();
                EmpSalLblConvallow.Hide(); EmpSalTxtConveyanceallow.Hide();
                EmpSalLblPhoneallow.Hide(); EmpSalTxtPhoneallow.Hide(); 
                EmpSalLblLatesittingallow.Hide(); EmpSalTxtLatesittingallow.Hide(); 
                EmpSalLblOtherallow.Hide();     EmpSalTxtOtherallow.Hide(); 

                
            }
            else
            {
                EmpSalGroupAlloDeductdetail.Hide();
            }
        }        
        private void EmpGentxtEname_KeyPress(object sender, KeyPressEventArgs e)
        {

            
            Deo_Module.Charverified.check_charctersfun(EmpGentxtEname, 0, e, 25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
        private void EmpGenTxtFname_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtFname, 0, e, 25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
            
        }
        private void EmpGenTxtCnic_KeyPress(object sender, KeyPressEventArgs e)
        {            
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtCnic, 1, e,13);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
        private void EmpGenTxtReligion_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtReligion, 0, e, 12);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
        private void EmpGenTxtDomicile_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtDomicile, 0, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
            
        }
        private void EmpGenTxtBnkacc_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtBnkacc, 1, e, 10);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
            
        }
        private void EmpGenTxtBnkbranch_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtBnkbranch, 2, e, 40);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
        private void EmpGenTxtEmail_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtEmail, 2, e, 25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
        private void EmpGenTxtPhno_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtPhno, 1, e, 10);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
            
        }
        private void EmpGenTxtTempadd_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtTempadd, 2, e, 40);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
            
        }
        private void EmpGenTxtPermadd_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtPermadd, 2, e, 40);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
        private void EmpGenTxtPwd_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtPwd, 2, e, 25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
            
        }
        private void EmpGenTxtSearchrec_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpGenTxtSearchrec, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
       
        
        //--Emp Monthly settings.
        private void EmpMonTxtName_KeyPress(object sender, KeyPressEventArgs e)
        {            
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtName, 0, e, 25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }
        private void EmpMonTxtBps_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtBps, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpMonTxtJobid_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtJobid, 0, e, 25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
            
        }
        private void EmpMonTxtAllowtotal_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtAllowtotal, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpMonTxtDedtotal_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtDedtotal, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
           
        }
        private void EmpMonTxtTotalamount_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtTotalamount, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpMonTxtFormid_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtFormid, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpMonTxtRfid_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpMonTxtRfid, 1, e, 8);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
            
        }
        private void EmpMonCbmDepartment_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Charverified.check_charctersfun(EmpMonCbmDepartment, 0, e);
        }
        //--Emp Salary settings.
        private void EmpSalTxtAllowamount_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtAllowamount, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
           
        }
        private void EmpSalTxtDeducamount_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtDeducamount, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
          
        }
        private void EmpSalTxtBasicpay_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtBasicpay, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
           
        }
        private void EmpSalTxtNetsalary_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtNetsalary, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
         
        }
        private void EmpSalTxtPresentdays_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtPresentdays, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
           
        }
        private void EmpSalTxtLeavedays_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtLeavedays, 1, e, 3);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
           
        }
        private void EmpSalTxtAbsentdays_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtAbsentdays, 1, e, 3);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
           
        }
        private void EmpSalTxtHolidays_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtHolidays, 1, e, 3);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtTotaldays_KeyPress(object sender, KeyPressEventArgs e)
        {
      
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtTotaldays, 1, e, 3);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtOtherdeduc_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtOtherdeduc, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtOtherallow_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtOtherallow, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtConveyanceallow_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtConveyanceallow, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtLatesittingallow_KeyPress(object sender, KeyPressEventArgs e)
        {
         
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtLatesittingallow, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtAdvadeduc_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtAdvadeduc, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtMedicalallow_KeyPress(object sender, KeyPressEventArgs e)
        {
          
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtMedicalallow, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtPhoneallow_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtPhoneallow, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtLatededuc_KeyPress(object sender, KeyPressEventArgs e)
        {
          
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtLatededuc, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtHouseallow_KeyPress(object sender, KeyPressEventArgs e)
        {
         
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtHouseallow, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtFormid_KeyPress(object sender, KeyPressEventArgs e)
        {
         
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtFormid, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpSalTxtRftag_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            Deo_Module.Charverified.check_charctersfun(EmpSalTxtRftag, 1, e, 8);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        //--Employee Attendance Settings.
        private void EmpAttTxtSearchbyrfid_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            Deo_Module.Charverified.check_charctersfun(EmpAttTxtSearchbyrfid, 1, e, 15);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void EmpAttTxtSearchbyformid_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(EmpAttTxtSearchbyformid, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmAttTxtRftag_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmAttTxtRftag, 1, e, 8);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmAttTxtFormid_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            Deo_Module.Charverified.check_charctersfun(MgmAttTxtFormid, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtDid_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtDid, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
           
        }
        private void MgmDeptTxtDname_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtDname, 0, e,25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                            
           
        }
        private void MgmDeptTxtMgrid_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtMgrid, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
            
        }
        private void MgmDeptTxtBps_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtBps, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
            
        }
        private void MgmDeptTxtJobid_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtJobid, 0, e, 25);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
            
        }
        private void MgmDeptTxtMinbpay_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtMinbpay, 1, e, 8);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
            
        }
        private void MgmDeptTxtMaxbpay_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtMaxbpay, 1, e, 8);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtPerIncr_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtPerIncr, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtLatesitallow_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtLatesitallow, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtPhoneallow_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtPhoneallow, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtConvallow_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtConvallow, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtMedicalallow_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtMedicalallow, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtHouserent_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtHouserent, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtOtherallow_KeyPress(object sender, KeyPressEventArgs e)
        {

            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtOtherallow, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }
        private void MgmDeptTxtLatededuc_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtLatededuc, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
        }

        private void MgmDeptTxtOtherdeduc_KeyPress(object sender, KeyPressEventArgs e)
        {
            Deo_Module.Charverified.check_charctersfun(MgmDeptTxtOtherdeduc, 1, e, 2);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                                
            
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            Application.Exit();
        }
        private void scanningToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Scanner newobj = new Scanner();
            newobj.Show();
            this.Hide();
        }
        private void dataEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Safeinfo.purpose_deo_hr_receptionst = 0;        //Its mean my purpose is for DEO Only.And when you load deo page then check this var and load for this purpose not for HR purpose.            
            Dataentry newobj = new Dataentry(); //Its mean Load DEO Form.
            newobj.Show();
            this.Hide();
        }
        private void hRVerifyingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Safeinfo.purpose_deo_hr_receptionst = 1;        //Its mean my purpose is for HR Only.
            OA_PS_RFID.RFTag_Loader newobj = new OA_PS_RFID.RFTag_Loader(); //Its mean Load DEO Form for HR purpose.
            newobj.Show();
            this.Hide();
            //Safeinfo.purpose_deo_hr_receptionst = 1;        //Its mean my purpose is for HR Only.
            //Dataentry newobj = new Dataentry(); //Its mean Load DEO Form for HR purpose.
            //newobj.Show();
            //this.Hide();
        }
        private void receptonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Safeinfo.purpose_deo_hr_receptionst = 0;        //Setting this variable==0, Its mean my purpose is JUST VIEW RECEPTION APP.            
            Recep newobj = new Recep();         //Its mean Load Reception Form.
            newobj.Show();
            this.Hide();
        }
        private void levToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Safeinfo.purpose_deo_hr_receptionst = 1;        //Setting this variable==1, Its mean my purpose is TO GIVE/SET LEAVE.
            Recep newobj = new Recep();         //Its mean Load Reception Form.
            newobj.Show();
            this.Hide();
        }

        private void MgmDeptTxtMinbpay_Leave(object sender, EventArgs e)
        {
            if (MgmDeptTxtMinbpay.Text == "")
            {
                MgmDeptTxtMinbpay.Focus();
                MgmDeptLblMinbpay.ForeColor = Color.Red;
            }
            else if (int.Parse(MgmDeptTxtMinbpay.Text.ToString())<3000)
            {
                MessageBox.Show("Min Basic Pay can not be less than Rs 3000.", "Field validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MgmDeptLblMinbpay.ForeColor = Color.Red;
                MgmDeptTxtMinbpay.Focus();
            }
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Opacity = 0.5;
            OA_PS_RFID.AboutForm newobj = new OA_PS_RFID.AboutForm();
            newobj.ShowDialog();

        }
        
        private void MgmGenCkbSetholidays_CheckStateChanged(object sender, EventArgs e)
        {
            if (MgmGenCkbSetholidays.Checked == true)
            {
                MgmGenCkbSetleave.Checked = false;
                MgmGenGroupBoxLeave.Hide();

                MgmGenCkbSetstatus.Checked = false;
                MgmGenGroupBoxStatus.Hide();

                MgmGenGroupboxSetholiday.Show();
            }
            else
            {
                MgmGenGroupboxSetholiday.Hide();                
            }
        }
        private void MgmGenCkbSetstatus_CheckStateChanged(object sender, EventArgs e)
        {
            if (MgmGenCkbSetstatus.Checked == true)
            {
                MgmGenCkbSetleave.Checked = false;
                MgmGenGroupBoxLeave.Hide();

                MgmGenCkbSetholidays.Checked = false;
                MgmGenGroupboxSetholiday.Hide();

                MgmGenGroupBoxStatus.Show();
            }
            else
            {
                MgmGenGroupBoxStatus.Hide();
            }
        }
        private void MgmGenCkbSetleave_CheckStateChanged(object sender, EventArgs e)
        {
            if (MgmGenCkbSetleave.Checked == true)
            {
                MgmGenCkbSetstatus.Checked = false;
                MgmGenCkbSetholidays.Checked = false;
                MgmGenGroupBoxLeave.Show();
                MgmGenGroupBoxStatus.Hide();
                MgmGenGroupboxSetholiday.Hide();
            }
            else
            {
                MgmGenGroupBoxLeave.Hide();
            }
        }
        private void MgmGenCkbSearchbyrf_CheckStateChanged(object sender, EventArgs e)
        {
            if (MgmGenCkbSearchbyrf.Checked == true)
            {
                MgmGenTxtSearchbyrf.Text = "";
                MgmGenTxtSearchbyrf.Enabled = true;
                MgmGenPicture.Image = null;
                MgmGenTxtName.Text = "";
                MgmGenTxtBps.Text = "";
                MgmGenTxtDepartment.Text = "";
                MgmGenGroupLeave.Hide();

                                
                MgmGenTxtSearchbyformid.Text="";
                MgmGenTxtSearchbyformid.Enabled=false;
                MgmGenCkbSearchbyformid.Checked = false;
                MgmGenCkbSearchbyformid.Enabled = false;
            }
            
            else 
            {
                MgmGenTxtSearchbyrf.Text = "";
                MgmGenTxtSearchbyrf.Enabled = true;
                MgmGenPicture.Image = null;
                MgmGenTxtName.Text = "";
                MgmGenTxtBps.Text = "";
                MgmGenTxtDepartment.Text = "";
                MgmGenGroupLeave.Hide();


                MgmGenTxtSearchbyformid.Text = "";
                MgmGenTxtSearchbyformid.Enabled = true;
                MgmGenCkbSearchbyformid.Checked = false;
                MgmGenCkbSearchbyformid.Enabled = true;
            }
        }
        private void MgmGenCkbSearchbyformid_CheckStateChanged(object sender, EventArgs e)
        {
            if (MgmGenCkbSearchbyformid.Checked == true)
            {
                MgmGenTxtSearchbyrf.Text = "";
                MgmGenTxtSearchbyrf.Enabled = true;
                MgmGenPicture.Image = null;
                MgmGenTxtName.Text = "";
                MgmGenTxtBps.Text = "";
                MgmGenTxtDepartment.Text = "";
                MgmGenGroupLeave.Hide();


                MgmGenTxtSearchbyrf.Enabled = false;                
                MgmGenCkbSearchbyrf.Checked = false;
                MgmGenCkbSearchbyrf.Enabled = false;
            }

            else
            {
                MgmGenTxtSearchbyrf.Text = "";
                MgmGenTxtSearchbyrf.Enabled = true;
                MgmGenPicture.Image = null;
                MgmGenTxtName.Text = "";
                MgmGenTxtBps.Text = "";
                MgmGenTxtDepartment.Text = "";
                MgmGenGroupLeave.Hide();


                MgmGenTxtSearchbyformid.Text = "";
                MgmGenCkbSearchbyrf.Enabled = true;
                MgmGenCkbSearchbyrf.Checked = false;
            }
        }
        private void MgmGenBtnClear_Click(object sender, EventArgs e)
        {
            MgmGenTxtSearchbyrf.Text = "";
            MgmGenTxtSearchbyrf.Enabled = true;
            MgmGenCkbSearchbyrf.Enabled = true;
            MgmGenCkbSearchbyrf.Checked = false;
                        
            MgmGenTxtSearchbyformid.Text = "";
            MgmGenTxtSearchbyformid.Enabled = true;
            MgmGenCkbSearchbyformid.Enabled = true;
            MgmGenCkbSearchbyformid.Checked = false;

            MgmGenTxtFormid.Text = "";
            MgmGenPicture.Image = null;
            MgmGenTxtName.Text = "";
            MgmGenTxtBps.Text = "";
            MgmGenTxtDepartment.Text = "";
            MgmGenGroupLeave.Hide();
        }
        private void MgmGenBtnSearch_Click(object sender, EventArgs e)
        {
            if (MgmGenCkbSearchbyrf.Checked == true)
            {
                if (MgmGenTxtSearchbyrf.Text == "")
                {
                    MessageBox.Show("Please Enter RFID #.", "Valid Rftag Required", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (rfdisp_fun(int.Parse(MgmGenTxtSearchbyrf.Text),true))//if record found or Picture displayed then execute below code.
                    {
                        try
                        {
                            
                        }
                        catch (OracleException ex) { MessageBox.Show(ex.Message, "Oracle Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        catch (Exception ex) { MessageBox.Show(ex.Message, "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        finally
                        {
                            gl_dbaseclsobj.cls_con.Close();
                        }
                    }
                }
            }
            else if (MgmGenCkbSearchbyformid.Checked == true)
            {
                if (MgmGenTxtSearchbyformid.Text== "")
                {
                    MessageBox.Show("Please Enter Form Id.", "Valid Form Id Required", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (rfdisp_fun(int.Parse(MgmGenTxtSearchbyformid.Text), false))//if record found or Picture displayed then execute below code.
                    {
                        try
                        {

                        }
                        catch (OracleException ex) { MessageBox.Show(ex.Message, "Oracle Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        catch (Exception ex) { MessageBox.Show(ex.Message, "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        finally
                        {
                            gl_dbaseclsobj.cls_con.Close();
                        }
                    }
                }
            }

            else 
            {
                MessageBox.Show("Please Choose any Option.", "Choose an Option", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            
            string str = "Select To_char( To_date('" + MgmGenDtTime.Value.Date + "','MM/DD/RRRR HH:MI:SS AM'), 'DD-MON-RRRR')" + " From dual";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(str);

            MgmGenTxtLeavedates.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();
            str = " Select To_date('" + MgmGenCtlDate1.Value.Date + "','MM/DD/RRRR HH:MI:SS AM')+" + int.Parse(numericUpDown1.Value.ToString()) + " From dual";

            //str = "Select  To_char(" + obj_4_rf.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper() + "+" + numericUpDown1.Value.ToString()+ ", 'DD-MON-RRRR') From Dual"; ;
            gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
            str = "Select To_char( To_date('" + gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString() + "','MM/DD/RRRR HH:MI:SS AM'), 'DD-MON-RRRR')" + " From dual";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
            MgmGenTxtLeavedates.Text = String.Concat(MgmGenTxtLeavedates.Text, "     TO     " + gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper()).ToUpper();
            /*
            obj_4_rf.Dml_UpdateAdapterfun("Select To_char( To_date('" + RecepCtlDate1.Value.Date + "','dd/mm/rrrr HH24:MI:SS'), 'DD-MON-RRRR')+"+numericUpDown1.Value.ToString()+" From Dual");
            RecepTxtLeavedates.Text = obj_4_rf.cls_dataset.Tables[0].Rows[0][0].ToString();
             */ 
        }
        private void MgmGenBtnAllow_Click(object sender, EventArgs e)
        {
            if (MgmGenCkbSetleave.Checked == true)
            {
                if (numericUpDown1.Value <= 0)
                {
                    MessageBox.Show("Please Select Number of Required Leave Days.", "Required Leave Days", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    if (MgmGenCtlDate1.Value.Day < DateTime.Today.Day)
                    {//You can not set Leave in Previous days.
                        //But you are allowed to set leave on TODAY date/sysdate.
                        MessageBox.Show("You are not allowed to Set Leave in Back days.\n You can Set Leave on Today/ Incoming days.", "Action Denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        DialogResult result = MessageBox.Show("Are you sure to set Leave on Selected Date.", "Leave Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            string str = "Select To_char( To_date('" + MgmGenCtlDate1.Value.Day + "/" + MgmGenCtlDate1.Value.Month + "/" + MgmGenCtlDate1.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                            gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
                            string mydate = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();
                            try
                            {
                                OracleCommand cmd = new OracleCommand();
                                gl_dbaseclsobj.cls_con.Open();
                                cmd.Connection = gl_dbaseclsobj.cls_con;
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.CommandText = "Req_2_Leave";

                                //Sending Formdid.
                                OracleParameter p0_days = new OracleParameter("Pv_days", OracleType.Number);
                                p0_days.Direction = ParameterDirection.Input;
                                p0_days.Value = Convert.ToInt32(numericUpDown1.Value);
                                cmd.Parameters.Add(p0_days);

                                OracleParameter p1_rfid = new OracleParameter("Pv_Formid", OracleType.Number);
                                p1_rfid.Direction = ParameterDirection.Input;
                                p1_rfid.Value = int.Parse(MgmGenTxtFormid.Text);
                                cmd.Parameters.Add(p1_rfid);

                                OracleParameter p2_rfid = new OracleParameter("Pv_particulardate", OracleType.DateTime);
                                p2_rfid.Direction = ParameterDirection.Input;
                                p2_rfid.Value = OracleDateTime.Parse(mydate);
                                cmd.Parameters.Add(p2_rfid);

                                OracleParameter p3_formid = new OracleParameter("Answer", OracleType.VarChar, 100);
                                p3_formid.Direction = ParameterDirection.Output;
                                cmd.Parameters.Add(p3_formid);
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("" + p3_formid.Value.ToString(), "Leave Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Action Incomplete", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            finally 
                            {
                                gl_dbaseclsobj.cls_con.Close();
                            }
                        }
                    }
                }
            }
            if (MgmGenCkbSetstatus.Checked == true)
            {
                if (MgmGenCmbStatus.Text == "")
                {
                    MessageBox.Show("Please choose a status to update result.", "Status Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    if (MgmGenDttimeHoliday.Value.Day < DateTime.Today.Day)
                    {//In previous dates you can not Update STATUS of An Employee.
                        //Rather You can UPDATE Status of An Employee Todays.
                        MessageBox.Show("You are not allowed to update Employee status in Back Days\nYou can update Record for Today/Incoming days.", "Action Denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    { 
                        DialogResult result = MessageBox.Show("Are you sure to update Employee status.", "Update Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            string str = "Select To_char( To_date('" + MgmGenDtTime.Value.Day + "/" + MgmGenDtTime.Value.Month + "/" + MgmGenDtTime.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                            gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
                            string mydate = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();
                            try
                            {
                                OracleCommand cmd = new OracleCommand();
                                gl_dbaseclsobj.cls_con.Open();
                                cmd.Connection = gl_dbaseclsobj.cls_con;
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.CommandText = "Set_Status";

                                //Sending Formdid.                            
                                OracleParameter p1_formid = new OracleParameter("Pv_Formid", OracleType.Number);
                                p1_formid.Direction = ParameterDirection.Input;
                                p1_formid.Value = int.Parse(MgmGenTxtFormid.Text);
                                cmd.Parameters.Add(p1_formid);

                                OracleParameter p2_date = new OracleParameter("Pv_Date", OracleType.DateTime);
                                p2_date.Direction = ParameterDirection.Input;
                                p2_date.Value = OracleDateTime.Parse(mydate);
                                cmd.Parameters.Add(p2_date);

                                OracleParameter p2_status = new OracleParameter("Pv_Status", OracleType.VarChar,20);
                                p2_status.Direction = ParameterDirection.Input;
                                p2_status.Value = MgmGenCmbStatus.Text;
                                cmd.Parameters.Add(p2_status);

                                OracleParameter p3_formid = new OracleParameter("Answer", OracleType.VarChar, 100);
                                p3_formid.Direction = ParameterDirection.Output;
                                cmd.Parameters.Add(p3_formid);
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("" + p3_formid.Value.ToString(), "Status Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Action Incomplete", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            finally
                            {
                                gl_dbaseclsobj.cls_con.Close();
                            }
                        }
                        else
                        {//Do nothing.

                        }                    
                    }
                }
            }

            if (MgmGenCkbSetholidays.Checked == true)
            {//Set holiday like on 25th dec,14Aug and etc.
                
                string str = "Select To_char( To_date('" + MgmGenDttimeHoliday.Value.Day + "/" + MgmGenDttimeHoliday.Value.Month + "/" + MgmGenDttimeHoliday.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
                str = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();
                if (MgmGenDttimeHoliday.Value.Day <= DateTime.Today.Day)
                {//In previous dates and TODAY Also You can not set Holidays. It should be clear.
                    MessageBox.Show("You are not allowed to Set Holidays in Back Days or Today.\nThis Option is Only For Incoming days.", "Action Denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DialogResult result = MessageBox.Show("Are you sure to set Holidays on '" + str + "'.", "Holiday Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        string script_query =
                                            " DECLARE " +
                                                " v_Formid Attendance.Form_id%type; " +
                                                " CURSOR C1 IS " +
                                                " Select Form_id From Employee Where formstatus_id=6; " +
                                            " Begin " +
                                                " open c1; " +
                                                " Fetch c1 into V_Formid; " +
                                                " While c1%Found Loop " +                                                    
                                                    " Insert into Attendance " +
                                                    " Values(V_Formid,'" + str + "','OFFICIAL HOLIDAY',00,00,00); " +
                                                    " Commit; " +
                                                    " Fetch c1 into V_Formid; " +
                            //Dbms_output.put_line('Number is Forced out');
                                                "  End Loop; " +
                                                " Close C1; " +
                                            " End; ";
                        if (gl_dbaseclsobj.Dml_Updatefun(script_query) != 0)
                        {
                            MessageBox.Show("All Employees have been Issued Holiday on '" + str + "'.", "Holiday Issued", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Already on Holiday", "Transaction Incomplete", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {//Do nothing.

                    }
                }
            }
        }
        private void Admin_Activated(object sender, EventArgs e)
        {
            this.Opacity = 1;
        }
        private void MgmDeptTxtDid_Leave(object sender, EventArgs e)
        {
            if (MgmDeptTxtDid.Text.Length <= 3)
            {
                MgmDeptTxtDid.Focus();            
            }
        }
        private void MgmDeptTxtJobid_Leave(object sender, EventArgs e)
        {
            if (MgmDeptTxtJobid.Text.Length < 2)
            {
                MessageBox.Show("Job Id can not be less than 02 Characters.", "Field validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MgmDeptTxtJobid.Focus();
            }
        }
        private void MgmDeptTxtMaxbpay_Leave(object sender, EventArgs e)
        {

            if (MgmDeptTxtMaxbpay.Text == "")
            {
                MgmDeptTxtMaxbpay.Focus();
                MgmDeptLblMaxbpay.ForeColor = Color.Red;
            }
            else if (int.Parse(MgmDeptTxtMaxbpay.Text.ToString()) < 3000)
            {
                MessageBox.Show("Max Basic Pay can not be less than Rs 3000.", "Field validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MgmDeptLblMaxbpay.ForeColor = Color.Red;
                MgmDeptTxtMaxbpay.Focus();
            }
        }
        private void MgmJobCkbdDeptwise_CheckedChanged(object sender, EventArgs e)
        {
            MgmJobCmb0Dept.Text = "";

            if (MgmJobCkbdDeptwise.Checked == true)
            {
                MgmJobGroupSelectedemp.Hide();
                int available_rows=MgmJobGridvu.RowCount;
                while(available_rows>0)                
                {
                    MgmJobGridvu.Rows.RemoveAt(available_rows-1);
                    available_rows--;
                }
                MgmJobCkbJobisnull.Enabled = false;
                MgmJobCkbdDeptwise.Checked = true;
                MgmJobCkbFormid.Enabled = false;
                MgmJobCmb0Dept.Visible = true;      //It enable Combo for work.
                gl_searchfor = 0;                   //Search According to department.
                MgmJobBtnSearch.Enabled = true;     //It enable button for work.
            }
            else
            {
                MgmJobCkbJobisnull.Enabled = true;
                MgmJobCkbdDeptwise.Checked = false;
                MgmJobCkbFormid.Enabled = true;
                MgmJobCmb0Dept.Visible = false;      //It Disable Combo.
                MgmJobBtnEdit.Enabled = false;
                gl_chk_editbtn = false;

                gl_searchfor = -2;                  //Re-Set it for next time usage.
                MgmJobBtnSearch.Enabled = false;    //It Disable button.
            }
        }

        private void MgmGenTxtSearchbyrf_KeyPress(object sender, KeyPressEventArgs e)
        {
            Charverified.check_charctersfun(MgmGenTxtSearchbyrf, 1, e, 8);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }

        private void MgmGenTxtSearchbyformid_KeyPress(object sender, KeyPressEventArgs e)
        {
            Charverified.check_charctersfun(MgmGenTxtSearchbyformid, 1, e, 4);//Int Allow. If Character are more than Its (Presented)limit, Then DO NOT Display Character.                                    
        }

        private void EmpSalGridvu_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i <= gl_grid_Rows; i++)
            {
                if (EmpSalGridvu.Rows[i].Cells[0].Selected == true)
                {
                    EmpSalGroupSalarydetail.Show();
                    EmpSalGroupdaysdetail.Show();

                    string strformid = EmpSalGridvu.Rows[i].Cells[1].Value.ToString();
                    show_salaryfun(0, strformid, false);//False mean, I don't want to Load Data into Grid, Rather show Attendance.
                    run_procefun(strformid);

                    gl_dbaseclsobj.Show_Datafun("Select Ename,Emp_Image From Employee Where Form_Id=" + strformid, 0);
                    if (gl_dbaseclsobj.cls_dr.Read())
                    {
                        EmpSalTxtName.Text = gl_dbaseclsobj.cls_dr.GetString(0).ToUpper();
                        //Picture loading code.
                        OracleLob blob = gl_dbaseclsobj.cls_dr.GetOracleLob(1);      //Column # .      Emp_image
                        Byte[] BLOBData = new Byte[blob.Length];
                        //Read blob data into byte array
                        int j = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));
                        //Get the primitive byte data into in-memory data stream
                        MemoryStream stmBLOBData = new MemoryStream(BLOBData);
                        //LOADING INTO PICBOX1                            
                        EmpSalPicbox.Image = Image.FromStream(stmBLOBData);
                        //MessageBox.Show("Now i'm reseting it.");
                        EmpSalPicbox.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                    gl_dbaseclsobj.Show_Datafun("Select Ename,Emp_Image From Employee Where Form_Id=" + strformid, 1);
                    //gl_dbaseclsobj.Show_Datafun("Select Ename, Emp_Image From Employee Where Form_Id=" + 2840, 1);

                    EmpSalCkbAllowamount.Checked = false;
                    EmpSalCkbDeducamount.Checked = false;

                }
            }
            
        }

        

       
        //End func
                
    }//End Class
}//End namespace.