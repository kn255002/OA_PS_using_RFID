using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;



namespace Deo_Module
{
    partial class Admin
    {
        // Required designer variable.
        //
        //
        string  gl_str = null;           //It can contain deptname,Emp/gen/Picbox path and etc.
        int     gl_grid_Rows = -1;         //It tells you how many row found in a grid.
        int     gl_form_id = 0;            //Whic Formid it will choose from query, pls retain it for feature use.
        bool    gl_chkbox_satuts = true;   //For checking 'Search' Ckbox status on Every Tab/Inner tab.
        short   gl_searchfor=-2;          //-2=nothing,-1=Search for where job is null,(0=form_id, 1=RFId, 2=Cnic and etc.
                                        //Byte can't have -1, so i choose short as datatype.

        private Update_Database gl_dbaseclsobj = new Update_Database();

        ///ONLY FOR Management/Job CONTROLS
        ///I uses these variable for Only this page.
        bool gl_chk_editbtn = false; //You did not press edit btn yet.


        //Data members only. 
        static string orcl9i = "Data Source=(DESCRIPTION="
             + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + Safeinfo.machinename + ")(PORT=1521)))"
             + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME="+Safeinfo.servicename+")));"
             + "User Id=pk_ocp;Password=me;";

        OracleConnection con = new OracleConnection(orcl9i);
        DataSet mydataset;

        /*****************************************************************************************/
        /**********************************Functions For Department/Scale Tab.**************************************/
        bool verify_scalefieldfun()
        {
            bool should_i_Update = true;
            if (MgmDeptTxtBps.Text == "")
            {
                MgmDeptLblBps.ForeColor = Color.Red;
                should_i_Update = false;
            }
            else
            { 
                MgmDeptLblBps.ForeColor = Color.Black; 
            }

            if(MgmDeptTxtJobid.Text=="")
            {
                MgmDeptLblJobid.ForeColor=Color.Red;
                should_i_Update = false;
            }
            else
            { 
                MgmDeptLblJobid.ForeColor = Color.Black; 
            }

            if(MgmDeptTxtMinbpay.Text=="")
            {
                MgmDeptLblMinbpay.ForeColor=Color.Red;
                should_i_Update = false;
            }
            else
            {
                MgmDeptLblMinbpay.ForeColor = Color.Black;                
            }


            if (MgmDeptTxtMaxbpay.Text == "")
            {
                MgmDeptLblMaxbpay.ForeColor = Color.Red;
                should_i_Update = false;
            }
            else
            {
                MgmDeptLblMaxbpay.ForeColor = Color.Black;
                if (MgmDeptTxtMinbpay.Text != "" && MgmDeptTxtMaxbpay.Text != "")
                {
                    if (int.Parse(MgmDeptTxtMaxbpay.Text.ToString()) < int.Parse(MgmDeptTxtMinbpay.Text.ToString()))
                    {
                        MgmDeptLblMinbpay.ForeColor = Color.Red;
                        MgmDeptLblMaxbpay.ForeColor = Color.Red;
                        should_i_Update = false;
                    }
                }
            }
            return should_i_Update;        
        }


        /*****************************************************************************************/
        /**********************************Now my Functions.**************************************/
        void show_deptfun()//Show data of Department Table.
        {
            try
            {
                con.Open();
                string dept_query = "Select * from Department order by 1";
                OracleDataAdapter admin_adapter = new OracleDataAdapter(dept_query, con);
                mydataset = new DataSet("Department");
                admin_adapter.Fill(mydataset, "Department");
                MgmDeptGridvu.DataSource = mydataset.Tables[0];
                gl_grid_Rows = -1;
                foreach (DataRow myrow in mydataset.Tables[0].Rows)
                {//keep record how many rows in Department table found.
                    ++gl_grid_Rows;               
                }

                if (MgmDeptGridvu.Columns[0].Name == "ChooseRow" || MgmDeptGridvu.Columns[1].Name == "DeleteRow")
                {   //Hey! Don't be worry we'll enable on pressing 'EDIT' Control.{                    
                    //MgmDeptGridvu.Columns.Remove(DeleteRow);    //Hey! Don't be worry we'll enable on pressing 'EDIT' Control.
                }                
                MgmDeptGridvu.ReadOnly = true;            
                MgmDeptGroupadddeptscale.Hide();
                //GridDeptvu.AllowUserToAddRows = false;
            }
            catch (OracleException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);                
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
            }
            finally
            {
                // MessageBox.Show("I'm going to close connection.");
                con.Close();
            }
        }                          
        void show_scalefun()//Show data of Scale/jobs Table
        {
            try
            {
                con.Open();
                string dept_query = "Select * from Scale order by 1";
                OracleDataAdapter admin_adapter = new OracleDataAdapter(dept_query, con);
                mydataset = new DataSet("Scale");
                admin_adapter.Fill(mydataset, "Scale");
                MgmDeptGridvu.DataSource = mydataset.Tables[0];
                gl_grid_Rows = -1;
                foreach (DataRow myrow in mydataset.Tables[0].Rows)
                {//keep record how many rows in Scale table found.
                    ++gl_grid_Rows;
                }

                if (MgmDeptGridvu.Columns[0].Name == "ChooseRow" || MgmDeptGridvu.Columns[1].Name == "DeleteRow")
                {//Hey! Don't be worry we'll enable on pressing 'EDIT' Control.                                        
                    //MgmDeptGridvu.Columns.Remove(DeleteRow);    //Hey! Don't be worry we'll enable on pressing 'EDIT' Control.
                }

                MgmDeptGridvu.ReadOnly = true;
                MgmDeptGroupadddeptscale.Hide();
                //GridDeptvu.AllowUserToAddRows = false;
            }
            catch (OracleException ex) { MessageBox.Show("Oracle Error Found.\n" + ex.ToString()); }
            catch (Exception ex) { MessageBox.Show("General Exception Found, Khuram.\n" + ex.ToString()); }
            finally
            {
                //MessageBox.Show("I'm going to disconnect.");
                con.Close();
            }
        
        }
        void show_Empfun()//Show data of Employees.
        {
            string search_query = null;
            switch (gl_searchfor)
            {//Now check here, Search for what, and make a Where clause.
                case -1: 
                {
                    search_query = "Select Form_id,Ename,bps,job_id,rfid_id,Hire_date,Joining_date,Emp_Type from Employee Where (Bps is null or Dept_id is null) And formstatus_id=4";                    
                    break; 
                }
                case 0:
                {

                    MgmJobBtnEdit.Enabled = true;   //Bcz it will return rows so enable 'Edit' Control.
                    if (gl_str ==null)
                    {//gl_str means here gl_str contain dept name value that you've selected.
                        search_query = "Select Form_id,Ename,bps,job_id,Dept_id,rfid_id,Hire_date,Joining_date,Admin_by from Employee Where (Dept_id is not  null And Bps is not null) And Formstatus_id=6";
                        //search_query = "Select Form_id,Ename,bps,job_id,rfid_id from Employee Where Dept_id =this selected name";                         
                    }
                    else
                    {   //Sub query to search dept_id.
                        search_query = "Select Form_id,Ename,bps,job_id,rfid_id,Emp_Type from Employee " + 
                                       "Where Dept_id=(select dept_id "+
                                                      "from Department "+
                                                      "where dept_name=" + gl_str + ") And Formstatus_id=6";
                    }
                    break;
                }
                case 1: 
                {
                    search_query = "Select Form_id,Ename,bps,job_id,Dept_id,rfid_id,Hire_date,Joining_date,Admin_by from Employee Where Form_id=" + gl_form_id + " And Formstatus_id=6";
                    //MgmJobBtnEdit.Enabled = true;   it is diable bcz it may return rows.
                    break; 
                }
                default: break;
            }//end switch.
            try
            {
                con.Open();
                OracleDataAdapter admin_adapter = new OracleDataAdapter(search_query, con);
                mydataset = new DataSet("Employee");
                admin_adapter.Fill(mydataset, "Employee");
                gl_grid_Rows = -1;      //putting gl_grid_rows=-1 or null.
                
                foreach (DataRow drow in mydataset.Tables[0].Rows)
                {
                    ++gl_grid_Rows; //As number of rows are there. it updates it self.
                }

                MgmJobGridvu.DataSource = mydataset.Tables[0];
                if (MgmJobGridvu.Rows.Count >= 1)
                {
                    MgmJobBtnEdit.Enabled = true;
                }
                else
                {
                    MgmJobBtnEdit.Enabled = false;
                }
            }
            catch (DataException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
            }
            finally
            {
                // MessageBox.Show("I'm going to close connection.");
                con.Close();
            }
        }
        void show_Attendfun(int chk_fun,string received)//Show data of Attendance for daily attendance.
        {
            string Attend_query = null;
            string Attenddetail_query = null;
            switch (chk_fun)
            {
                case 0:
                    {//Show All Employees details do not sperate according to dept. on this date.
                        string str = "Select To_char( To_date('" + MgmAttDTP.Value.Day + "/" + MgmAttDTP.Value.Month + "/" + MgmAttDTP.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                        gl_dbaseclsobj.Dml_UpdateAdapterfun(str);

                        str = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();

                        //Attend_query = "select Distinct e.Form_id,E.Ename,E.job_id,D.Dept_name " +
                        //    "From Employee E,Attendance A, Department D" +
                        //    " where (E.form_id=A.form_id And E.Dept_id=D.Dept_id) And E.Formstatus_id=6 And (A.Date_Time) like '" + str + "' "; //And date=this
                        //SHOW RESULT FOR ALL DEPARTMENTS.
                        Attend_query = "select Distinct E.Form_id,E.Ename,D.Dept_name,E.Job_id,B.HH,B.MM,B.SS "+
                                       "From Employee E,Department D,Attendance A, "+
                                          "( Select Form_id,Sum(Nvl(HH,0))HH,Sum(Nvl(MM,0))MM,Sum(Nvl(ss,0))SS " +
	                                         "From Attendance "+
	                                         "Where Form_Id in (Select Form_Id From Employee Where Dept_Id is not null and Formstatus_Id=6) "+
                                             "And TO_Char(Date_Time,'DD-MON-RR')='" + str + "' "+
	                                         "Group by Form_Id "+
                                          ")B  "+
                                        "Where  (E.Dept_id=D.Dept_id And E.Form_id=A.Form_id) "+
                                        "And (E.Formstatus_id=6) "+
                                        "And (E.Form_id=B.Form_id) "+
                                        "And TO_Char(A.Date_Time,'DD-MON-RR')='" + str + "' ";


                        break;
                    }
                case 1:
                    {//Show All Employees details According to Selected Department. on this date.\
                        string str = "Select To_char( To_date('" + MgmAttDTP.Value.Day + "/" + MgmAttDTP.Value.Month + "/" + MgmAttDTP.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                        gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
                        str = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();

                        //Attend_query = "select Distinct E.Form_id,E.Ename,E.job_id " +
                        //    "from Employee E,Attendance A " +
                        //    "where E.form_id=A.form_id And "+
                        //    "E.Dept_id=(Select Dept_id From Department where Dept_name='"+received+"') And "+"E.Formstatus_id=6 And Upper(A.Date_Time) like '" + str + "'"; //And date=this.
                        //SHOW RESULT FOR ONLY selected DEPARTMENT.
                        Attend_query =
                            "select Distinct E.Form_id,E.Ename,D.Dept_name,E.Job_id,B.HH,B.MM,B.SS " +
                            "From Employee E,Department D,Attendance A, "+
                            "("+
                            "Select Form_id,Sum(Nvl(HH,0))HH,Sum(Nvl(MM,0))MM,Sum(Nvl(ss,0))SS " +
                            "From Attendance "+
                            "Where Form_Id in "+
                                            "("+
			                                "Select Form_Id "+
                                            "From Employee "+
                                            "Where Dept_Id =(Select Dept_id From Department where Dept_name='" + received + "') " +
				                            "And   Formstatus_Id=6 "+
		                                    ") "+
                            "And TO_Char(Date_Time,'DD-MON-RR')='" + str + "' " +
		                    "Group by Form_Id "+
                            ")B "+
                            "Where "+
	                        "(E.Form_id=B.Form_id) "+
	                        "And "+
	                        "(E.Dept_id=D.Dept_id)"+
	                        "And"+
	                        "(E.Form_id=A.Form_id) "+
	                        "And"+
                            "(E.Formstatus_id=6 AND E.dept_Id=(Select Dept_id From Department where Dept_name='" + received + "') )" +
	                        "And "+
                            "TO_Char(A.Date_Time,'DD-MON-RR')='" + str + "'";                          
                        break;
                    }
                case 2:
                    {//Show All Employees details According to Provided Form_id.
                        string str = "Select To_char( To_date('" + MgmAttDTP.Value.Day + "/" + MgmAttDTP.Value.Month + "/" + MgmAttDTP.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                        gl_dbaseclsobj.Dml_UpdateAdapterfun(str);

                        str = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();
                /*
                        Attend_query = "select Distinct E.Form_id,E.Ename,E.job_id " +
                            "from Employee E,Attendance A " +
                            "where E.form_id=A.form_id And (E.Rfid_id='"+int.Parse(received)+"' OR E.Form_id="+int.Parse(received)+" )And E.Formstatus_id=6 And Upper(A.Date_Time) like '" + str + "'"; //And date=this
                 */

                        Attend_query = " Select Distinct E.Form_id,E.Ename,D.Dept_name,E.Job_id,B.HH,B.MM,B.SS " +
                                       " From Employee E,Department D,Attendance A, " +
                                       " ( Select Form_id,Sum(Nvl(HH,0))HH,Sum(Nvl(MM,0))MM,Sum(Nvl(ss,0))SS " +
                                       " From Attendance " +
                                       " Where Form_id=" +int.Parse(received)+
                                       " Group by Form_Id " +
                                       " )B " +
                                       " Where      " +
                                       " (E.Dept_id=D.Dept_id And E.Form_id=A.Form_id) " +
                                       " And " +
                                       " (E.Formstatus_id=6 And E.Form_id="+int.Parse(received) +
                                       " ) And " +
                                       " (E.Form_id=B.Form_id) " +
                                       " And " +
                                       " TO_Char(A.Date_Time,'DD-MON-RR')='"+str+"'";
                        break;                        
                    }
                //case 3:
                //    {//Show All Employees details According to Provided Form_id.
                //        string str = "Select To_char( To_date('" + MgmAttDTP.Value.Day + "/" + MgmAttDTP.Value.Month + "/" + MgmAttDTP.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                //        gl_dbaseclsobj.Dml_UpdateAdapterfun(str);

                //        str = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper() + "%";
                //        Attend_query = "select Distinct E.Form_id,E.Ename,E.job_id " +
                //            "from Employee E,Attendance A " +
                //            "where (E.form_id=A.form_id And"+" E.Form_id="+int.Parse(received)+") And E.Formstatus_id=6 And Upper(A.Date_Time) like '"+str+"'"; //And date=this
                //        break;
                //    }

                case 99:
                    {//Estring str = "Select To_char( To_date('" + MgmAttDTP.Value.Date + "','DD/MM/RRRR HH24:MI:SS'), 'DD-MON-RR')" + " From dual";
                        string str = "Select To_char( To_date('" + MgmAttDTP.Value.Day + "/" + MgmAttDTP.Value.Month + "/" + MgmAttDTP.Value.Year + "','dd/mm/RRRR'), 'DD-MON-RR')" + " From dual";
                       // string str = "Select To_char( To_date('" + MgmAttDTP.Value.Date + "','DD/MM/RRRR HH24:MI:SS'), 'DD-MON-RR')" + " From dual";
                        gl_dbaseclsobj.Dml_UpdateAdapterfun(str);

                        str = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();
                        Attenddetail_query = "select E.form_id,E.ename as Name,E.job_id,To_char(Date_Time)Date_Time,A.status,NVL(A.HH,0)||':'||NVL(A.MM,0)||':'||NVL(A.SS,0) Spent_Time " +
                            "from Employee E,Attendance A " +
                            "where E.form_id=A.form_id " +
                            "And A.form_id=" + gl_form_id +
                            " And E.Formstatus_id=6 "+
                            " And Upper(A.Date_Time) like '" + str + "%'" +
                            " Order by 1";//date=this date.
                        break;
                    }


            }//End switch.
            try
                {
                    if (chk_fun == 0 || chk_fun == 1 || chk_fun == 2 || chk_fun == 3)
                    {//Its mean Show All Employees details do not sperate according to dept.
                        gl_dbaseclsobj.Grid_Loaderfun(MgmAttGridAttendVu, Attend_query);
                        gl_grid_Rows = -1;
                        foreach (DataRow drow in gl_dbaseclsobj.cls_dataset.Tables[0].Rows)
                        {//Count rows, How many rows found?
                            ++gl_grid_Rows;
                        }

                        { //Now check Time (hh,mm,ss).
                            if (MgmAttGridAttendVu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                            {
                                int orig_min, orig_hr, orig_ss, Temp_ss;
                                orig_hr = orig_min = orig_ss = Temp_ss = 00;

                                for (int i = 0; i < MgmAttGridAttendVu.RowCount; i++)
                                {
                                    orig_hr = int.Parse(MgmAttGridAttendVu.Rows[i].Cells[5].Value.ToString());    //I've found HOUR column.
                                    orig_min = int.Parse(MgmAttGridAttendVu.Rows[i].Cells[6].Value.ToString());   //I've found MINute column.
                                    Temp_ss = int.Parse(MgmAttGridAttendVu.Rows[i].Cells[7].Value.ToString());   //I've found Second column.

                                    if (Temp_ss >= 60)
                                    {
                                        orig_min += Temp_ss / 60;      //Here if sec>60 then put minute into Orig_min.
                                        orig_ss = Temp_ss % 60;        //Here U make Original Seconds.
                                    }
                                    else
                                    {
                                        orig_ss = Temp_ss;
                                    }
                                    if (orig_min >= 60)
                                    {
                                        orig_hr += orig_min / 60;      //Here if sec>60 then put minute into Orig_min.
                                        orig_min = orig_min % 60;        //Here U make Original Seconds.
                                    }
                                    //Now i show here FINAL HR,MIN,SEC to display in grid.
                                    MgmAttGridAttendVu.Rows[i].Cells[5].Value = orig_hr;
                                    MgmAttGridAttendVu.Rows[i].Cells[6].Value = orig_min;
                                    MgmAttGridAttendVu.Rows[i].Cells[7].Value = orig_ss;
                                }
                            }
                        }




                    }
                    else if(chk_fun ==99)
                    {//Its mean Show Particular Emp Attend in its own Particular Grid.
                        gl_dbaseclsobj.Grid_Loaderfun(MgmAttGridAttdetailvu, Attenddetail_query);
                    }
                    if (gl_grid_Rows < 0)
                    {
                        MessageBox.Show("No Data Found.", "No Data Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MgmAttGroupDetailview.Hide();
                    }
                    
                    /*//Now run query as we choosed.
                    con.Open();
                    OracleDataAdapter admin_adapter = new OracleDataAdapter(Attend_query, con);
                    mydataset = new DataSet("Attendance");
                    admin_adapter.Fill(mydataset, "Attendance");
                     if (chk_fun == true)
                    {
                        gl_grid_Rows = -1;
                        foreach (DataRow drow in mydataset.Tables[0].Rows)
                        {
                            ++gl_grid_Rows;
                        }
                    }
                    if (gl_form_id == 0)
                    {
                        MgmAttGridAttendVu.DataSource = mydataset.Tables[0];
                    }
                    else            //Its mean Show Particular Emp Attend in its own Particular Grid.
                    {
                        MgmAttGridAttdetailvu.DataSource = mydataset.Tables[0];
                    }*/
                    //MessageBox.Show("Col name:" + GridAttendVu.Columns[0].Selected);
                }
                catch (OracleException Oex)
                {
                    MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
                }
                finally
                {
                    // MessageBox.Show("I'm going to close connection.");
                    con.Close();
                }
            //If ends
        }            
        
        void search_Empfun(short ck_flag ,long disid)
        {
            string srch_query = null;
            switch (ck_flag)
            {
                case -1://This case when serching by (disid=)Form_id AND JOB NOT ALLOCATED YET.
                    {
                        srch_query = "Select * From Employee Where Form_id=" + disid + " And Formstatus_id=4";
                        break;
                    }
                case 0: //This case when serching by (disid=)RFID_id.
                    {
                        int a = Convert.ToInt32(disid);
                        srch_query = "Select * From Employee Where Rfid_id=" + a + " And Formstatus_id=6";
                        break;
                    }
                case 1: //This case when serching by (disid=)Form_id.
                    {
                        int formid =Convert.ToInt32(disid);
                        if (MgmJobGridvu.CurrentRow.Cells[3].Value.ToString().Length>0)
                        {
                            srch_query = "Select * From Employee Where Form_id=" + formid + " And Formstatus_id=6";                            
                        }
                        else
                        {
                            srch_query = "Select * From Employee Where Form_id=" + formid + " And Formstatus_id=4";                            
                        }
                        break;
                        }
                
                case 2://This case when serching by (disid=)CNIC.
                        {
                            srch_query = "Select * From Employee Where Cnic=" + disid + " And Formstatus_id=6";
                            break;
                        }
                //No need to write default stmt.

            }//End of Switch.            
            try
            {
                con.Open();
                OracleCommand cmd = new OracleCommand(srch_query, con);
                OracleDataReader dr = cmd.ExecuteReader();
                Boolean recordExist = dr.Read();

                if (recordExist)
                {//If record found then work on following TABS.
                    
  
                    if (MgmJobTxtFormid1.Text != "")//If it is Mgm/Job's Tab page.
                    {//this will tell that search_Empfun() result will be only for Mgm/Job's Tab page.

                        MgmJobTxtName.Text = dr.GetString(1).ToUpper();
                        if (dr.GetOracleNumber(20).IsNull)
                        {//its mean you are searching for those records where Job is null;
                            //MgmJobTxtBps.Hide();
                            MgmJobTxtJobid.Hide();

                            MgmJobLblJoindate.Text = "Hire Date";
                            MgmJobLblJoindate.Show();
                            MgmJobCtlHiredate.Enabled = true;
                            MgmJobCtlHiredate.Show();

                            //MgmJobCmb0Bps.Visible = false;
                            //MgmJobLblBps.Visible = false;
                            
                            MgmJobCmb0Jobid.Show();
                            MgmJobCmbEmptype.Show();
                            //MgmJobCmb0Bps.Text ="";
                            MgmJobCmb0Jobid.Text ="";
                            MgmJobCmb0Dept.Text = "";
                            MgmJobTxtRftag.Text = "";
                            //MgmJobCmb0Bps.Text = "";     // You didn't alloted yet and its result for dbase already is null.
                            MgmJobCmb0Jobid.Text = "";  // You didn't alloted yet and its result for dbase already is null.
                            MgmJobCmb0Dept.Text = "";   // You didn't alloted yet and its result for dbase already is null.
                            MgmJobCmbEmptype.Text = ""; // You didn't alloted yet and its result for dbase already is null.
                            MgmJobTxtRftag.Text = dr.GetString(19).ToString();                                                        
                            if (dr.GetString(26).ToString().ToUpper() == "C")
                            {
                                MgmJobCmbEmptype.Text = "CONTRACT";
                            }
                            else
                            {
                                MgmJobCmbEmptype.Text = "PERMANENT";
                            }



                            //string scale_query= "Select Bps,Job_id From Scale";
                            //gl_dbaseclsobj.Dml_UpdateAdapterfun(scale_query);                            
                            //for (int i = 0; i < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; i++)
                            //{
                            //  MgmJobCmb0Bps.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]);
                            //  MgmJobCmb0Jobid.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][1]);                                
                            //  //MgmJobCmb0Jobid.Items.Insert(i, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][0]+"--"+gl_dbaseclsobj.cls_dataset.Tables[0].Rows[i][1]);                                
                            //}


                            //string dept_query = "Select Dept_Name From Department";
                            //gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_query);                            
                            //for (int j = 0; j < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; j++)
                            //{
                            //    MgmJobCmb1Dept.Items.Insert(j, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[j][0]);
                            //}
                        }
                        
                        else
                        {
                            //MgmJobCmb0Bps.Hide();
                            MgmJobCmb0Jobid.Hide();
                            MgmJobLblJoindate.Show();
                            MgmJobCtlHiredate.Show();
                            
                            //MgmJobTxtBps.Show();                            
                            MgmJobTxtJobid.Show();
                            MgmJobTxtRftag.Text = dr.GetString(19).ToString();
                            //MgmJobTxtBps.Text = dr.GetDecimal(20).ToString();
                            MgmJobTxtJobid.Text = dr.GetString(21).ToUpper();
                            string dept_query = "Select Dept_Name From Department";
                            gl_dbaseclsobj.Dml_UpdateAdapterfun(dept_query);
                            for (int j = 0; j < gl_dbaseclsobj.cls_dataset.Tables[0].Rows.Count; j++)
                            {
                                MgmJobCmb1Dept.Items.Insert(j, gl_dbaseclsobj.cls_dataset.Tables[0].Rows[j][0]);
                            }
                            //MgmJobCmb1Dept.Text = dr.GetDecimal(23).ToString();

                            if (dr.GetString(26).ToString().ToUpper() == "C")
                            {
                                MgmJobCmbEmptype.Text = "CONTRACT";
                            }
                            else
                            {
                                MgmJobCmbEmptype.Text = "PERMANENT";
                            }
                            
                        }
                            {//Picture loading code.
                                OracleLob blob = dr.GetOracleLob(29);      //Column # .      Emp_image
                                Byte[] BLOBData = new Byte[blob.Length];
                                //Read blob data into byte array
                                int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));
                                //Get the primitive byte data into in-memory data stream
                                MemoryStream stmBLOBData = new MemoryStream(BLOBData);
                                //LOADING INTO PICBOX1                            
                                MgmJobPicbox.Image = Image.FromStream(stmBLOBData);
                                //MessageBox.Show("Now i'm reseting it.");
                                MgmJobPicbox.SizeMode = PictureBoxSizeMode.StretchImage;
                            }
                    }//Display for Mgm/Job's Tab page ends.
                    

                    if (EmpGenTxtSearchrec.Text != "")
                    {//this will tell that search_Empfun() result will be only for Emp/General's Tab page.
                        if (EmpGenBtnEdit.Enabled == true)
                        {
                            EmpGenBtnSave.Enabled = false;
                        }
                        make_readonlyfun(0,true);       //Here 0 Tells Set readonly=True for Every control of Emp/Gen Tab, True mean by setting Readonly=true don't change background color.
                        EmpGenBtnCancel.Text = "Clear"; //Now make 'Cancel' Button to 'Clear' Button.
                        EmpGenBtnEdit.Enabled = true;   //Bcz Data found, so keep Enable "Edit" button
                        EmpGenBtnCancel.Enabled = true; //Bcz Data found, so keep Enable "Cancel" button

                        EmpGentxtEname.Text = dr.GetString(1).ToUpper();
                        EmpGenTxtFname.Text = dr.GetString(2).ToUpper();
                        EmpGenTxtCnic.Text = dr.GetOracleNumber(3).ToString();
                        EmpGenTxtDob.Text = dr.GetOracleDateTime(4).ToString();
                        EmpGenTxtReligion.Text = dr.GetString(6);
                        EmpGenTxtDomicile.Text = dr.GetString(7).ToUpper();
                        EmpGenTxtEmail.Text = dr.GetString(8);
                        EmpGenTxtPhno.Text = dr.GetDecimal(9).ToString();
                        EmpGenTxtBnkacc.Text = dr.GetDecimal(10).ToString();
                        EmpGenTxtBnkbranch.Text = dr.GetString(11).ToUpper();
                        EmpGenTxtPwd.Text = dr.GetString(17).ToUpper();
                        EmpGenTxtLogname.Text = dr.GetDecimal(0).ToString();
                        int temp = int.Parse(dr.GetDecimal(18).ToString());
                        switch (temp)
                        {
                            case 0: EmpGenRdbDisable.Checked = true; break;
                            case 1: EmpGenRdbEnable.Checked = true; break;
                        }
                        EmpGenTxtRfid.Text = dr.GetString(19);
                        EmpGenTxtHiredate.Text = dr.GetOracleDateTime(24).ToString();
                        EmpGenTxtTempadd.Text = dr.GetString(27).ToUpper();
                        EmpGenTxtPermadd.Text = dr.GetString(28).ToUpper();
                        {//Picture loading code.
                            OracleLob blob = dr.GetOracleLob(29);      //Column # .      Emp_image
                            Byte[] BLOBData = new Byte[blob.Length];
                            //Read blob data into byte array
                            int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));
                            //Get the primitive byte data into in-memory data stream
                            MemoryStream stmBLOBData = new MemoryStream(BLOBData);
                            //LOADING INTO PICBOX1
                            EmpGenPicbox.Image = Image.FromStream(stmBLOBData);
                            //MessageBox.Show("Now i'm reseting it.");
                            EmpGenPicbox.SizeMode = PictureBoxSizeMode.StretchImage;
                        }
                    }//Display for Emp/Gen's Tab page ends.

                    //Display for Emp/Adv's Tab page ends.(That also contain History grid).
                    
                    //TxtScannedby.Show();
                    //TxtScannedby.ReadOnly = true;
                    //LblScannedby.Show();
                        
                    //TxtPrivewformid.Text = dr.GetDecimal(0).ToString();     //Column # .      Form_id
                    //TxtScannedby.Text = dr.GetString(1);
                   
                }
                else
                {
                    EmpGenBtnEdit.Enabled = false; //Bcz No data found, so keep disable "Edit" button
                    MessageBox.Show("Sorry! No data found against your provided information.", "Record Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OracleException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                if (ex.Message.ToUpper() == "PARAMETER IS NOT VALID.")
                {
                    MessageBox.Show("Employee Picture not found.","Missing Value",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
                finally
                {
                    // MessageBox.Show("I'm going to close connection.");
                    con.Close();
                }
               
        }//End of Function.
        void mgmdept_hideshowfun(bool check)
        {
            if (check == true)//its mean You are calling this function for hide/Show controls for Insertion in Dept or Scale table.
            {
                if (MgmDeptRbDept.Checked == true)
                {
                    MgmDeptLblDeptscaleview.Text = "Departments detailed view.";
                    MgmDeptGroupadddeptscale.Text = "Add Department";
                    MgmDeptLblDeptScaleadd.Text = "Department Id must be Unique.";
                    {//Hide all detail of Scale/Jobs controls in Add Group
                        MgmDeptLblBps.Hide();                        MgmDeptTxtBps.Hide();
                        MgmDeptLblJobid.Hide();                      MgmDeptTxtJobid.Hide();
                        MgmDeptLblMinbpay.Hide();                    MgmDeptTxtMinbpay.Hide();
                        MgmDeptLblMaxbpay.Hide();                    MgmDeptTxtMaxbpay.Hide();
                        MgmDeptLblPerIncr.Hide();                    MgmDeptTxtPerIncr.Hide();
                        MgmDeptLblHouserent.Hide();                  MgmDeptTxtHouserent.Hide();
                        MgmDeptLblMedicalallow.Hide();               MgmDeptTxtMedicalallow.Hide();
                        MgmDeptLblConvallow.Hide();                  MgmDeptTxtConvallow.Hide();
                        MgmDeptLblPhoneallow.Hide();                 MgmDeptTxtPhoneallow.Hide();
                        MgmDeptLblOtherallow.Hide();                 MgmDeptTxtOtherallow.Hide();
                        MgmDeptLblLatededuc.Hide();                  MgmDeptTxtLatededuc.Hide();
                        MgmDeptLblLatesitallow.Hide();               MgmDeptTxtLatesitallow.Hide();
                        MgmDeptLblOtherdeduc.Hide();                 MgmDeptTxtOtherdeduc.Hide();

                        {//Now if you hide Jobs/Scale controls,Then u must show Dept controls.
                            MgmDeptLblDid.Show();   MgmDeptTxtDid.Show(); MgmDeptTxtDid.Text = "";
                            MgmDeptLblDname.Show(); MgmDeptTxtDname.Show(); MgmDeptTxtDname.Text = "";
                            MgmDeptLblMgrid.Show(); MgmDeptTxtMgrid.Show(); MgmDeptTxtMgrid.Text = "";
                        }
                    }//End hide/show control blocks.
                }// Dept end if
                else if (MgmDeptRbJobscales.Checked == true)
                {
                    MgmDeptLblDeptscaleview.Text = "Job Scale detailed view.";
                    MgmDeptGroupadddeptscale.Text = "Add Job scale";
                    MgmDeptLblDeptScaleadd.Text = "BPS and Job Id must be Unique. Keeping Null field mean default value allocate.";
                    {//Hide all detail of Departments controls in Add Group.
                        MgmDeptLblDid.Hide();
                        MgmDeptTxtDid.Hide();
                        MgmDeptLblDname.Hide();
                        MgmDeptTxtDname.Hide();
                        MgmDeptLblMgrid.Hide();
                        MgmDeptTxtMgrid.Hide();
                        {//Now if you hide depart controls,Then u must show Scale/Jobs controls.
                            MgmDeptLblBps.Show();           MgmDeptTxtBps.Show();           MgmDeptTxtBps.Text = "";
                            MgmDeptLblJobid.Show();         MgmDeptTxtJobid.Show();         MgmDeptTxtJobid.Text = "";
                            MgmDeptLblMinbpay.Show();       MgmDeptTxtMinbpay.Show();       MgmDeptTxtMinbpay.Text = "";
                            MgmDeptLblMaxbpay.Show();       MgmDeptTxtMaxbpay.Show();       MgmDeptTxtMaxbpay.Text = "";
                            MgmDeptLblPerIncr.Show();       MgmDeptTxtPerIncr.Show();       MgmDeptTxtPerIncr.Text = "";
                            MgmDeptLblHouserent.Show();     MgmDeptTxtHouserent.Show();     MgmDeptTxtHouserent.Text = "";
                            MgmDeptLblMedicalallow.Show();  MgmDeptTxtMedicalallow.Show();  MgmDeptTxtMedicalallow.Text = "";
                            MgmDeptLblConvallow.Show();     MgmDeptTxtConvallow.Show();     MgmDeptTxtConvallow.Text = "";
                            MgmDeptLblPhoneallow.Show();    MgmDeptTxtPhoneallow.Show();    MgmDeptTxtPhoneallow.Text = "";
                            MgmDeptLblOtherallow.Show();    MgmDeptTxtOtherallow.Show();    MgmDeptTxtOtherallow.Text = "";
                            MgmDeptLblLatededuc.Show();     MgmDeptTxtLatededuc.Show();     MgmDeptTxtLatededuc.Text="";
                            MgmDeptLblLatesitallow.Show();  MgmDeptTxtLatesitallow.Show();  MgmDeptTxtLatesitallow.Text="";
                            MgmDeptLblOtherdeduc.Show();    MgmDeptTxtOtherdeduc.Show();    MgmDeptTxtOtherdeduc.Text="";
                        }
                    }//End hide/show control blocks.
                }//Job/Scale end else.
            }//Hide+show control for Insertion purpose end.
            else if (check == false)//its mean You are calling this function for hide/Show controls for Updation in Dept or Scale table.
            {
                if (MgmDeptRbDept.Checked == true)
                {
                    MgmDeptLblDeptscaleview.Text = "Departments detailed view.";
                    MgmDeptGroupadddeptscale.Text = "Particular department view";
                    MgmDeptLblDeptScaleadd.Text = "Department Id must be Unique.";
                    {//Hide all detail of Scale/Jobs controls in Add Group
                        MgmDeptLblBps.Hide(); MgmDeptTxtBps.Hide();
                        MgmDeptLblJobid.Hide(); MgmDeptTxtJobid.Hide();
                        MgmDeptLblMinbpay.Hide(); MgmDeptTxtMinbpay.Hide();
                        MgmDeptLblMaxbpay.Hide(); MgmDeptTxtMaxbpay.Hide();
                        MgmDeptLblPerIncr.Hide(); MgmDeptTxtPerIncr.Hide();
                        MgmDeptLblHouserent.Hide(); MgmDeptTxtHouserent.Hide();
                        MgmDeptLblMedicalallow.Hide(); MgmDeptTxtMedicalallow.Hide();
                        MgmDeptLblConvallow.Hide(); MgmDeptTxtConvallow.Hide();
                        MgmDeptLblPhoneallow.Hide(); MgmDeptTxtPhoneallow.Hide();
                        MgmDeptLblOtherallow.Hide(); MgmDeptTxtOtherallow.Hide();

                        MgmDeptLblLatededuc.Hide(); MgmDeptTxtLatededuc.Hide();
                        MgmDeptLblLatesitallow.Hide(); MgmDeptTxtOtherdeduc.Hide();
                        MgmDeptLblOtherdeduc.Hide(); MgmDeptTxtLatesitallow.Hide();

                        {//Now if you hide Jobs/Scale controls,Then u must show Dept controls.
                            MgmDeptLblDid.Show();
                            MgmDeptTxtDid.Show(); MgmDeptTxtDid.Text = "";
                            MgmDeptLblDname.Show();
                            MgmDeptTxtDname.Show(); MgmDeptTxtDname.Text = "";
                            MgmDeptLblMgrid.Show();
                            MgmDeptTxtMgrid.Show(); MgmDeptTxtMgrid.Text = "";
                        }
                    }
                }

                else if (MgmDeptRbJobscales.Checked == true)
                {
                    MgmDeptLblDeptscaleview.Text = "Job Scale detailed view.";
                    MgmDeptGroupadddeptscale.Text = "Particular scale view";
                    MgmDeptLblDeptScaleadd.Text = "BPS and Job Id must be Unique. Keeping Null field mean default value allocate.";
                    {//Hide all detail of Departments controls in Add Group.
                        MgmDeptLblDid.Hide(); MgmDeptTxtDid.Hide();
                        MgmDeptLblDname.Hide(); MgmDeptTxtDname.Hide();
                        MgmDeptLblMgrid.Hide(); MgmDeptTxtMgrid.Hide();
                      {//Now if you hide depart controls,Then u must show Scale/Jobs controls.
                        MgmDeptLblBps.Show();
                        MgmDeptTxtBps.Show(); MgmDeptTxtBps.Text = "";
                        MgmDeptLblJobid.Show();
                        MgmDeptTxtJobid.Show(); MgmDeptTxtJobid.Text = "";
                        MgmDeptLblMinbpay.Show();
                        MgmDeptTxtMinbpay.Show(); MgmDeptTxtMinbpay.Text = "";
                        MgmDeptLblMaxbpay.Show();
                        MgmDeptTxtMaxbpay.Show(); MgmDeptTxtMaxbpay.Text = "";
                        MgmDeptLblPerIncr.Show();
                        MgmDeptTxtPerIncr.Show(); MgmDeptTxtPerIncr.Text = "";
                        MgmDeptLblHouserent.Show();
                        MgmDeptTxtHouserent.Show(); MgmDeptTxtHouserent.Text = "";
                        MgmDeptLblMedicalallow.Show();
                        MgmDeptTxtMedicalallow.Show(); MgmDeptTxtMedicalallow.Text = "";
                        MgmDeptLblConvallow.Show();
                        MgmDeptTxtConvallow.Show(); MgmDeptTxtConvallow.Text = "";
                        MgmDeptLblPhoneallow.Show();
                        MgmDeptTxtPhoneallow.Show(); MgmDeptTxtPhoneallow.Text = "";
                        MgmDeptLblOtherallow.Show();
                        MgmDeptTxtOtherallow.Show(); MgmDeptTxtOtherallow.Text = "";

                        MgmDeptLblLatededuc.Show(); MgmDeptTxtLatededuc.Show();
                        MgmDeptLblLatesitallow.Show(); MgmDeptTxtOtherdeduc.Show();
                        MgmDeptLblOtherdeduc.Show(); MgmDeptTxtLatesitallow.Show();

                      }
                    }
                }//End hide/show control blocks.
            }//Hide+show control for Update purpose end.


        }
        void make_readonlyfun(byte Readonly_4Tab, bool ck)
        {//This is for Emp/Gen Tab's Page. It'll make all controls read only with window original color.
            //Bcz when u make control readonly it changes color. so i've sett color. Now after making readonly a control, color will not change.   
            switch (Readonly_4Tab)
            {
                case 0:
                    {//Case 0 For Emp/Gen Tab.
                        if (ck == true)
                        {   //True mean on this page,By Default Everything wll be readonly, And color will look like same.(I mean no change in color.)
                            EmpGentxtEname.ReadOnly = true;
                            EmpGenTxtFname.ReadOnly = true;
                            EmpGenTxtCnic.ReadOnly = true;
                            EmpGenTxtDob.ReadOnly = true;
                            EmpGenTxtReligion.ReadOnly = true;
                            EmpGenTxtDomicile.ReadOnly = true;
                            EmpGenTxtHiredate.ReadOnly = true;
                            EmpGenTxtRfid.ReadOnly = true;
                            EmpGenTxtBnkacc.ReadOnly = true;
                            EmpGenTxtBnkbranch.ReadOnly = true;
                            EmpGenTxtEmail.ReadOnly = true;
                            EmpGenTxtPhno.ReadOnly = true;
                            EmpGenTxtTempadd.ReadOnly = true;
                            EmpGenTxtPermadd.ReadOnly = true;
                            EmpGenTxtLogname.ReadOnly = true;
                            EmpGenTxtPwd.ReadOnly = true;
                            EmpGenRdbDisable.Enabled = false;
                            EmpGenRdbEnable.Enabled = false;
                            EmpGenBtnBrowse.Enabled = false;
                            EmpGentxtEname.BackColor = Color.White;
                            EmpGenTxtFname.BackColor = Color.White;
                            EmpGenTxtCnic.BackColor = Color.White;
                            EmpGenTxtDob.BackColor = Color.White;
                            EmpGenTxtReligion.BackColor = Color.White;
                            EmpGenTxtDomicile.BackColor = Color.White;
                            EmpGenTxtHiredate.BackColor = Color.White;
                            EmpGenTxtRfid.BackColor = Color.White;
                            EmpGenTxtBnkacc.BackColor = Color.White;
                            EmpGenTxtBnkbranch.BackColor = Color.White;
                            EmpGenTxtEmail.BackColor = Color.White;
                            EmpGenTxtPhno.BackColor = Color.White;
                            EmpGenTxtTempadd.BackColor = Color.White;
                            EmpGenTxtPermadd.BackColor = Color.White;
                            EmpGenTxtLogname.BackColor = Color.White;
                            EmpGenTxtPwd.BackColor = Color.White;
                            EmpGenRdbDisable.BackColor = Color.Transparent;
                            EmpGenRdbEnable.BackColor = Color.Transparent;
                            EmpGenBtnBrowse.BackColor = Color.Transparent;
                        }
                        else
                        { //ck=False mean on this page,Some control will be readonly, With Color changed. And some will be edit able.

                            EmpGentxtEname.ReadOnly = true;
                            EmpGenTxtFname.ReadOnly = true;
                            EmpGenTxtCnic.ReadOnly = true;
                            EmpGenTxtDob.ReadOnly = true;
                            EmpGenTxtReligion.ReadOnly = true;
                            EmpGenTxtDomicile.ReadOnly = true;
                            EmpGenTxtHiredate.ReadOnly = true;
                            EmpGenTxtRfid.ReadOnly = true;
                            EmpGenTxtBnkacc.ReadOnly = false;
                            EmpGenTxtBnkbranch.ReadOnly = false;
                            EmpGenTxtEmail.ReadOnly = false;
                            EmpGenTxtPhno.ReadOnly = false;
                            EmpGenTxtTempadd.ReadOnly = false;
                            EmpGenTxtPermadd.ReadOnly = false;
                            EmpGenTxtLogname.ReadOnly = true;
                            EmpGenTxtPwd.ReadOnly = false;
                            EmpGenRdbDisable.Enabled = true;
                            EmpGenRdbEnable.Enabled = true;
                            EmpGenBtnBrowse.Enabled = true;
                            //Now setting color.                
                            EmpGentxtEname.BackColor = Color.Linen;
                            EmpGenTxtFname.BackColor = Color.Linen;// Linen;
                            EmpGenTxtCnic.BackColor = Color.Linen;
                            EmpGenTxtDob.BackColor = Color.Linen;
                            EmpGenTxtReligion.BackColor = Color.Linen;
                            EmpGenTxtDomicile.BackColor = Color.Linen;
                            EmpGenTxtHiredate.BackColor = Color.Linen;
                            EmpGenTxtRfid.BackColor = Color.Linen;
                            EmpGenTxtBnkacc.BackColor = Color.White;
                            EmpGenTxtBnkbranch.BackColor = Color.White;
                            EmpGenTxtEmail.BackColor = Color.White;
                            EmpGenTxtPhno.BackColor = Color.White;
                            EmpGenTxtTempadd.BackColor = Color.White;
                            EmpGenTxtPermadd.BackColor = Color.White;
                            EmpGenTxtLogname.BackColor = Color.White;
                            EmpGenTxtPwd.BackColor = Color.White;
                            EmpGenRdbEnable.BackColor = Color.Transparent;
                            EmpGenRdbDisable.BackColor = Color.Transparent;
                            EmpGenBtnBrowse.BackColor = Color.White;
                        }
                        break;
                    }//case 0 Ends.Here Readonly set color and etc setted for Emp/Gen Tab.
                case 1:
                    {
                        break;
                    }//case 1 Ends.
                    }//End Switch.
            }//End fun.
        void clear_controlfun(int tabpage)
        {
            switch (tabpage)
            {
                case 0:
                    {   //Its mean you are executing following code for Emp/Gen Tab.
                        for (int i = 0; i < TabGeneral.Controls.Count; i++)
                        {
                            if (TabGeneral.Controls[i].ToString().ToUpper().StartsWith("SYSTEM.WINDOWS.FORMS.TEXTBOX") == true)//Found only those controls, who are TextBox datatype. and then set them NULL.
                            {
                                TabGeneral.Controls[i].ResetText();
                            }
                        }
                        EmpGenPicbox.Image = null;
                        EmpGenCkbCnic.Checked = false;
                        EmpGenCkbFormid.Checked = false;
                        EmpGenCkbRfid.Checked = false;
                        EmpGenTxtSearchrec.Text = "";
                        EmpGenTxtLogname.Text = "";
                        EmpGenTxtPwd.Text = "";
                        EmpGenRdbEnable.Checked = false;
                        EmpGenRdbDisable.Checked = false;
                        EmpGenBtnCancel.Enabled = false;
                        break;
                    }
                case 1:
                    {   
                        // EmpAdvTxtUpdateby                       
                        
                        break;
                    }
            }
        }
        void show_monthlysettingfun(int chk_fun, string received)
        {
            string Select_query = null;            
            switch (chk_fun)
            {
                case 0:
                    {//Show All Employees details do not sperate according to dept.
                        Select_query = "Select E.Form_id,E.Ename,S.Min_basic_pay,S.percent_incr,S.House_inper,S.Medical_inper,S.conveyance_inper,S.Phone_inper,S.Latesit_inper,S.OtherAll_inper,S.Lateded_inper,S.Otherded_inper " +
                            " From Employee E, Scale S " +
                            " Where E.Bps=S.Bps And E.Job_id=S.Job_id And E.Formstatus_id=6";
                        break;
                    }
                case 1:
                    {//Show All Employees details According to Selected Department.
                        Select_query = "Select E.Form_id,E.Ename,S.Min_basic_pay,S.percent_incr,S.House_inper,S.Medical_inper,S.conveyance_inper,S.Phone_inper,S.Latesit_inper,S.OtherAll_inper,S.Lateded_inper,S.Otherded_inper " +
                            " From Employee E, Scale S " +
                            " Where E.Bps=S.Bps And E.Job_id=S.Job_id And "+
                            " E.Dept_id=(Select Dept_id From Department where Dept_name='" + received + "') And E.Formstatus_id=6";
                        break;
                    }
                case 2:
                    {//Show All Employees details According to Provided Rf ID.

                        Select_query = "Select E.Form_id,E.Ename,S.Min_basic_pay,S.percent_incr,S.House_inper,S.Medical_inper,S.conveyance_inper,S.Phone_inper,S.Latesit_inper,S.OtherAll_inper,S.Lateded_inper,S.Otherded_inper " +
                            " From Employee E, Scale S " +
                            " Where E.Bps=S.Bps And E.Job_id=S.Job_id And E.Rfid_id=" + int.Parse(received) + " And E.Formstatus_id=6";
                        break;
                    }
                case 3:
                    {//Show All Employees details According to Provided Form_id.
                        Select_query = "Select E.Form_id,E.Ename,S.Min_basic_pay,S.percent_incr,S.House_inper,S.Medical_inper,S.conveyance_inper,S.Phone_inper,S.Latesit_inper,S.OtherAll_inper,S.Lateded_inper,S.Otherded_inper " +
                           " From Employee E, Scale S " +
                           " Where E.Bps=S.Bps And E.Job_id=S.Job_id And E.Form_id=" + int.Parse(received) + " And E.Formstatus_id=6";
                        break;
                    }                   
            }//End switch.
            try
            {
                if (chk_fun == 0 || chk_fun == 1 || chk_fun == 2 || chk_fun == 3)
                {//Its mean Show All Employees details do not sperate according to dept.
                    gl_dbaseclsobj.Grid_Loaderfun(EmpMonGridvu, Select_query);
                    gl_grid_Rows = -1;
                    foreach (DataRow drow in gl_dbaseclsobj.cls_dataset.Tables[0].Rows)
                    {//Count rows, How many rows found?
                        ++gl_grid_Rows;
                    }
                }
               
                if (gl_grid_Rows < 0)
                {
                    MessageBox.Show("No Data Found.", "No Data Found", MessageBoxButtons.OK, MessageBoxIcon.Information);                    
                }                
            }
            catch (OracleException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
            }
            finally
            {
                // MessageBox.Show("I'm going to close connection.");
                con.Close();
            }
            //If ends
        }
        void show_EmpAttfun(string uniqueid)
        {
            string str = "Select To_char( To_date('" + EmpAttCtlDatetime.Value.Day + "/" + EmpAttCtlDatetime.Value.Month + "/" + EmpAttCtlDatetime.Value.Year + "','dd/mm/RRRR'), 'MON-RR')" + " From dual";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
            string mydate = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();

            OracleCommand cmd = new OracleCommand();
            gl_dbaseclsobj.cls_con.Open();
            cmd.Connection = gl_dbaseclsobj.cls_con;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "Proc_daily_Time_calculate";

            //Sending Formdid.
            OracleParameter p1_formid = new OracleParameter("Pv_Formid", OracleType.Number);
            p1_formid.Direction = ParameterDirection.Input;
            p1_formid.Value = uniqueid;
            cmd.Parameters.Add(p1_formid);

            OracleParameter P2_Datelike = new OracleParameter("Pv_Datelike", OracleType.DateTime);
            P2_Datelike.Direction = ParameterDirection.Input;
            P2_Datelike.Value = OracleDateTime.Parse(mydate);
            cmd.Parameters.Add(P2_Datelike);

            OracleParameter p_refcursor = new OracleParameter("my_cursor", OracleType.Cursor);
            p_refcursor.Direction = ParameterDirection.Output;
            cmd.Parameters.Add(p_refcursor);

            DataSet Ds = new DataSet();
            OracleDataAdapter Da = new OracleDataAdapter(cmd);
            //gl_dbaseclsobj.cls_odadptr = new OracleDataAdapter(cmd);
            Da.Fill(Ds);            
            EmpAttGridvu.DataSource = Ds.Tables[0].DefaultView;
            { //Now check Time (hh,mm,ss).
                if (EmpAttGridvu.RowCount >= 1) //Before Loading of this option Data. FIRST REMOVE PREVIOUS/Already Displayed DATA.
                {
                    int orig_min, orig_hr, orig_ss, Temp_ss;
                    orig_hr = orig_min = orig_ss = Temp_ss = 00;
                    
                    for (int i = 0; i < EmpAttGridvu.RowCount; i++)
                    {
                        object aab = EmpAttGridvu.Rows[i].Cells[7].Value;
                        orig_hr = int.Parse(EmpAttGridvu.Rows[i].Cells[6].Value.ToString());    //I've found HOUR column.
                        orig_min = int.Parse(EmpAttGridvu.Rows[i].Cells[7].Value.ToString());   //I've found MINute column.
                        Temp_ss = int.Parse(EmpAttGridvu.Rows[i].Cells[8].Value.ToString());    //I've found Second column.

                        if (Temp_ss >= 60)
                        {
                            orig_min += Temp_ss / 60;      //Here if sec>60 then put minute into Orig_min.
                            orig_ss = Temp_ss % 60;        //Here U make Original Seconds.
                        }
                        else
                        {
                            orig_ss = Temp_ss;
                        }
                        if (orig_min >= 60)
                        {
                            orig_hr += orig_min / 60;      //Here if sec>60 then put minute into Orig_min.
                            orig_min = orig_min % 60;        //Here U make Original Seconds.
                        }
                        //Now i show here FINAL HR,MIN,SEC to display in grid.
                        EmpAttGridvu.Rows[i].Cells[6].Value = orig_hr;
                        EmpAttGridvu.Rows[i].Cells[7].Value = orig_min;
                        EmpAttGridvu.Rows[i].Cells[8].Value = orig_ss;
                    }
                }             
            }

            try
            {
                string query = "Select * From Employee Where Rfid_Id='" + uniqueid + "' Or Form_id=" + uniqueid;
                gl_dbaseclsobj.Show_Datafun(query, 0);
                Boolean recordExist = gl_dbaseclsobj.cls_dr.Read();
                {//Picture loading code.
                    OracleLob blob = gl_dbaseclsobj.cls_dr.GetOracleLob(29);      //Column # .      Emp_image
                    Byte[] BLOBData = new Byte[blob.Length];
                    //Read blob data into byte array
                    int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));
                    //Get the primitive byte data into in-memory data stream
                    MemoryStream stmBLOBData = new MemoryStream(BLOBData);
                    //LOADING INTO PICBOX1                            
                    EmpAttPicBox.Image = Image.FromStream(stmBLOBData);
                    //MessageBox.Show("Now i'm reseting it.");
                    EmpAttPicBox.SizeMode = PictureBoxSizeMode.StretchImage;
                }
                gl_dbaseclsobj.Show_Datafun(query, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Might be picture loading problem.\n" + ex.Message, "Oracle Error");
            }
            finally { gl_dbaseclsobj.cls_con.Close(); }

        }//End Show_EmpAttfun()
        void monthly_setfun(int Formid)
        {
            int ha, ma, ca, pa, la, oa, ld, od;

            if (EmpMonCkbHouserent.Checked == true) ha = 1; else ha = 0;
            if (EmpMonCkbMedallowance.Checked == true) ma = 1; else ma = 0;
            if (EmpMonCkbConvAllowance.Checked == true) ca = 1; else ca = 0;
            if (EmpMonCkbPhoneall.Checked == true) pa = 1; else pa = 0;
            if (EmpMonCkbLatesitall.Checked == true) la = 1; else la = 0;
            if (EmpMonCkbOtherall.Checked == true) oa = 1; else oa = 0;
            if (EmpMonCkbLateded.Checked == true) ld = 1; else ld = 0;
            if (EmpMonCkbOtherded.Checked == true) od = 1; else od = 0;
            
            OracleCommand cmd = new OracleCommand();
            gl_dbaseclsobj.cls_con.Open();
            cmd.Connection = gl_dbaseclsobj.cls_con;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "Proc_Set_Mon_Allow";
            

            //Sending Formdid.
            OracleParameter p1_formid = new OracleParameter("Pv_Formid", OracleType.Number);
            p1_formid.Direction=ParameterDirection.Input;
            p1_formid.Value=Formid;
            cmd.Parameters.Add(p1_formid);

            //Sending House Allowance True or not.
            OracleParameter p2_ha = new OracleParameter("Pv_House", OracleType.Number);
            p2_ha.Direction = ParameterDirection.Input;
            p2_ha.Value = ha;
            cmd.Parameters.Add(p2_ha);

            //Sending Medical Allowance True or not.
            OracleParameter p3_ma = new OracleParameter("Pv_Medical", OracleType.Number);
            p3_ma.Direction = ParameterDirection.Input;
            p3_ma.Value = ma;
            cmd.Parameters.Add(p3_ma);

            //Sending Convence Allowance True or not.
            OracleParameter p4_ca = new OracleParameter("Pv_Convey", OracleType.Number);
            p4_ca.Direction = ParameterDirection.Input;
            p4_ca.Value = ca;
            cmd.Parameters.Add(p4_ca);

            //Sending Phone Allowance True or not.
            OracleParameter p5_pa = new OracleParameter("Pv_Phone", OracleType.Number);
            p5_pa.Direction = ParameterDirection.Input;
            p5_pa.Value = pa;
            cmd.Parameters.Add(p5_pa);

            //Sending Latesitting Allowance True or not.
            OracleParameter p6_la = new OracleParameter("Pv_Latesit", OracleType.Number);
            p6_la.Direction = ParameterDirection.Input;
            p6_la.Value = la;
            cmd.Parameters.Add(p6_la);

            //Sending Other Allowance True or not.
            OracleParameter p7_oa = new OracleParameter("Pv_Otherallow", OracleType.Number);
            p7_oa.Direction = ParameterDirection.Input;
            p7_oa.Value = oa;
            cmd.Parameters.Add(p7_oa);

            //Sending Late deduction True or not.
            OracleParameter p8_ld = new OracleParameter("Pv_Lateded", OracleType.Number);
            p8_ld.Direction = ParameterDirection.Input;
            p8_ld.Value = ld;
            cmd.Parameters.Add(p8_ld);

            //Sending Other deduction True or not.
            OracleParameter p9_od = new OracleParameter("Pv_Otherdeduc", OracleType.Number);
            p9_od.Direction = ParameterDirection.Input;
            p9_od.Value = od;
            cmd.Parameters.Add(p9_od);

            //Receiving Total Allowances in a text box.
            OracleParameter p10_TotalAllowances = new OracleParameter("Pv_TotalAllowances", OracleType.Number);
            p10_TotalAllowances.Direction = ParameterDirection.Output;
            cmd.Parameters.Add(p10_TotalAllowances);

            //Receiving Total deduction in a text box.
            OracleParameter p11_TotalDeductions = new OracleParameter("Pv_TotalDeductions", OracleType.Number);
            p11_TotalDeductions.Direction = ParameterDirection.Output;
            cmd.Parameters.Add(p11_TotalDeductions);
            
            //Receiving Total in a text box.
            OracleParameter p12_Total = new OracleParameter("Pv_Total", OracleType.Number);
            p12_Total.Direction = ParameterDirection.Output;
            cmd.Parameters.Add(p12_Total);

            OracleParameter p13_Netsal = new OracleParameter("Pv_Netsalary", OracleType.Number);
            p13_Netsal.Direction = ParameterDirection.Output;
            cmd.Parameters.Add(p13_Netsal);

            OracleParameter p13_Answer = new OracleParameter("Answer", OracleType.VarChar,200);
            p13_Answer.Direction = ParameterDirection.Output;
            cmd.Parameters.Add(p13_Answer);

            
            try
            {
                cmd.ExecuteNonQuery();
                EmpMonTxtAllowtotal.Text = cmd.Parameters["Pv_TotalAllowances"].Value.ToString();
                EmpMonTxtDedtotal.Text = cmd.Parameters["Pv_TotalDeductions"].Value.ToString();
                EmpMonTxtTotalamount.Text = cmd.Parameters["Pv_Total"].Value.ToString();
                MessageBox.Show("" + cmd.Parameters["Answer"].Value,"Record Status",MessageBoxButtons.OK, MessageBoxIcon.Information);
                if(cmd.Parameters["Answer"].Value.ToString().ToUpper()=="RECORD UPDATED.")
                {
                    EmpMonGroupSet0.Hide();
                }

            }            
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message, "Oracle Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "External Eror Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                gl_dbaseclsobj.cls_con.Close();
            }

            



        }
        void show_salaryfun(int chk_fun, string received,bool purpose)
        {    //This function work for TWO purpose
             // 1). Display data into Data Grid.
             // 2). Display Attendance.
            if (purpose)
            {//Purpose True mean just display data into Data Grid.
                string str = "Select To_char( To_date('" + EmpSalDatetime.Value.Day + "/" + EmpSalDatetime.Value.Month + "/" + EmpSalDatetime.Value.Year + "','dd/mm/RRRR'), 'MON-RRRR')" + " From dual";
                gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
                string mydate = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();


                string Select_query = null;
                switch (chk_fun)
                {
                    case 0:
                        {//Show All Employees details do not sperate according to dept.
                            Select_query = "Select Emp.Form_id,Emp.eNAME Name,Scl.Min_basic_pay,Sry.Sal_Mon,Sry.Basic_pay,Sry.Total_Salary " +
                                " From Employee Emp, Scale Scl, Department Dpt, Salary sry " +
                                " Where ( Emp.Bps=Scl.Bps And Emp.Job_id=Scl.Job_id) And " +
                                " (Emp.Dept_id=Dpt.Dept_id And " +
                                " Emp.Form_id=Sry.Form_id)  And " +
                                " ( Emp.Form_id=Sry.Form_id And To_char(Sry.Sal_mon,'MON-RRRR') Like ('%" + mydate + "')) " +
                                " And Emp.Formstatus_id=6 ";
                            break;
                        }
                    case 1:
                        {//Show All Employees details According to Selected Department.
                            Select_query = "Select Emp.Form_id,Emp.eNAME Name,Scl.Min_basic_pay,To_char(Sry.Sal_Mon,'MONTH-RRRR')Salary_Month,Sry.Basic_pay,Sry.Total_Salary " +
                                 " From Employee Emp, Scale Scl, Department Dpt, Salary sry " +
                                " Where ( Emp.Bps=Scl.Bps And Emp.Job_id=Scl.Job_id) And " +
                                " (Emp.Dept_id=Dpt.Dept_id And " +
                                " Emp.Form_id=Sry.Form_id)  And " +
                                " ( Emp.Form_id=Sry.Form_id And To_char(Sry.Sal_mon,'MON-RRRR') Like ('%" + mydate + "') )And (Dpt.Dept_Name='" + received + "' And Emp.Formstatus_id=6)";
                            break;
                        }
                    case 2:
                        {//Show All Employees details According to Provided Rf ID.
                            Select_query = "Select Emp.Form_id,Emp.eNAME Name,Scl.Min_basic_pay,To_char(Sry.Sal_Mon,'MONTH-RRRR')Salary_Month,Sry.Basic_pay,Sry.Total_Salary " +
                                 " From Employee Emp, Scale Scl, Department Dpt, Salary sry " +
                                " Where ( Emp.Bps=Scl.Bps And Emp.Job_id=Scl.Job_id) And " +
                                " (Emp.Dept_id=Dpt.Dept_id And " +
                                " Emp.Form_id=Sry.Form_id)  And " +
                                " ( Emp.Form_id=Sry.Form_id And To_char(Sry.Sal_mon,'MON-RRRR') Like ('%" + mydate + "')) " +
                                " And (Emp.Rfid_id=" + int.Parse(received) + " And Emp.Formstatus_id=6)";
                            break;
                        }
                    case 3:
                        {//Show All Employees details According to Provided Form_id.
                            Select_query = "Select Emp.Form_id,Emp.eNAME Name,Scl.Min_basic_pay,To_char(Sry.Sal_Mon,'MONTH-RRRR')Salary_Month,Sry.Basic_pay,Sry.Total_Salary " +
                                 " From Employee Emp, Scale Scl, Department Dpt, Salary sry " +
                                " Where ( Emp.Bps=Scl.Bps And Emp.Job_id=Scl.Job_id) And " +
                                " (Emp.Dept_id=Dpt.Dept_id And " +
                                " Emp.Form_id=Sry.Form_id)  And " +
                                " ( Emp.Form_id=Sry.Form_id And To_char(Sry.Sal_mon,'MON-RRRR') Like ('%" + mydate + "')) " +
                                " And (Emp.Form_id=" + int.Parse(received) + " And Emp.Formstatus_id=6)";
                            break;
                        }
                }//End switch.

                    if (chk_fun == 0 || chk_fun == 1 || chk_fun == 2 || chk_fun == 3)
                    {//Its mean Show All Employees details do not sperate according to dept.
                        gl_dbaseclsobj.Grid_Loaderfun(EmpSalGridvu, Select_query);
                        gl_grid_Rows = -1;
                        foreach (DataRow drow in gl_dbaseclsobj.cls_dataset.Tables[0].Rows)
                        {//Count rows, How many rows found?
                            ++gl_grid_Rows;
                        }
                    }

                    if (gl_grid_Rows < 0)
                    {
                        MessageBox.Show("No Data Found.", "No Data Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        EmpSalPicbox.Image = null;
                        EmpSalTxtName.Text = "";
                        EmpSalGroupSalarydetail.Hide();
                        EmpSalGroupdaysdetail.Hide();
                        EmpSalGroupAlloDeductdetail.Hide();
                    }
                }//If true ends.
            if (!purpose)
            {//Purpose False mean Now Display Attendance.
                string str = "Select To_char( To_date('" + EmpSalDatetime.Value.Day + "/" + EmpSalDatetime.Value.Month + "/" + EmpSalDatetime.Value.Year + "','dd/mm/RRRR'), 'MON-RRRR')" + " From dual";
                gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
                string mydate = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();
                try
                {
                    {
                        gl_dbaseclsobj.cls_con.Open();
                        OracleCommand cmd = new OracleCommand();
                        cmd.Connection = gl_dbaseclsobj.cls_con;

                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "Tell_Attendance";

                        OracleParameter p1 = new OracleParameter("Pv_Formid", OracleType.Number);
                        p1.Direction = ParameterDirection.Input;
                        p1.Value = int.Parse(received);
                        cmd.Parameters.Add(p1);

                        OracleParameter p2 = new OracleParameter("Pv_date", OracleType.DateTime);
                        p2.Direction = ParameterDirection.Input;
                        p2.Value = OracleDateTime.Parse(mydate);
                        cmd.Parameters.Add(p2);

                        OracleParameter p3 = new OracleParameter("Pv_Present", OracleType.Number);
                        p3.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p3);

                        OracleParameter p4 = new OracleParameter("Pv_Absent", OracleType.Number);
                        p4.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p4);

                        OracleParameter p5 = new OracleParameter("Pv_Leave", OracleType.Number);
                        p5.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p5);

                        OracleParameter p6 = new OracleParameter("Pv_Holidays", OracleType.Number);

                        p6.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p6);

                        OracleParameter p7 = new OracleParameter("Pv_Totaldays", OracleType.Number);
                        p7.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p7);

                        OracleParameter p8 = new OracleParameter("Answer", OracleType.VarChar, 200);
                        p8.Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(p8);

                        cmd.ExecuteNonQuery();
                        EmpSalTxtPresentdays.Text = cmd.Parameters["Pv_Present"].Value.ToString();
                        EmpSalTxtAbsentdays.Text = cmd.Parameters["Pv_Absent"].Value.ToString();
                        EmpSalTxtLeavedays.Text = cmd.Parameters["Pv_Leave"].Value.ToString();
                        EmpSalTxtHolidays.Text = cmd.Parameters["Pv_Holidays"].Value.ToString();
                        EmpSalTxtTotaldays.Text = cmd.Parameters["Pv_Totaldays"].Value.ToString();
                        gl_dbaseclsobj.cls_con.Close();
                    }
                }
                catch (OracleException ex) { MessageBox.Show(ex.Message, "Oracle Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                catch (Exception ex) { MessageBox.Show(ex.Message, "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
               finally
               {
                   gl_dbaseclsobj.cls_con.Close();
               } 
            }
        }
                    //////////////////******Here Allowances+Deduction code work.******/////////////////                   
        void run_procefun(string received)
        {
            string str = "Select To_char( To_date('" + EmpSalDatetime.Value.Day + "/" + EmpSalDatetime.Value.Month + "/" + EmpSalDatetime.Value.Year + "','dd/mm/RRRR'), 'MON-RRRR')" + " From dual";
            gl_dbaseclsobj.Dml_UpdateAdapterfun(str);
            string mydate = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString();
            try
            {
                
                
                gl_dbaseclsobj.cls_con.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = gl_dbaseclsobj.cls_con;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "Show_Mon_Allow";

                //Sending Formdid.
                OracleParameter p1_formid = new OracleParameter("Pv_Formid", OracleType.Number);
                p1_formid.Direction = ParameterDirection.Input;
                p1_formid.Value = int.Parse(received);
                cmd.Parameters.Add(p1_formid);
                                
                OracleParameter p2 = new OracleParameter("Pv_Month_Year", OracleType.DateTime);
                p2.Direction = ParameterDirection.Input;
                p2.Value = OracleDateTime.Parse(mydate);
                cmd.Parameters.Add(p2);

                OracleParameter p3 = new OracleParameter("Pv_Allowances_Amount", OracleType.Number);
                p3.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(p3);

                OracleParameter p4 = new OracleParameter("Pv_Deductions_Amount", OracleType.Number);
                p4.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(p4);  

                OracleParameter p5 = new OracleParameter("Pv_Netsalary", OracleType.Number);
                p5.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(p5);

                OracleParameter p6 = new OracleParameter("Pv_Basicpay", OracleType.Number);
                p6.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(p6);                                                   
                
                OracleParameter p7 = new OracleParameter("Pv_house", OracleType.Number);
                p7.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(p7);
                
                OracleParameter p8 = new OracleParameter("PV_medical", OracleType.Number);
                p8.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(p8);
                
                OracleParameter p9 = new OracleParameter("PV_convey", OracleType.Number);
                p9.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(p9);
                  
                  OracleParameter p10 = new OracleParameter("PV_Phone", OracleType.Number);
                  p10.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p10);
                  
                  OracleParameter p11 = new OracleParameter("Pv_latesit", OracleType.Number);
                  p11.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p11);
                  
                  OracleParameter p12 = new OracleParameter("Pv_otherallow", OracleType.Number);
                  p12.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p12);
                                 
                  OracleParameter P13 = new OracleParameter("Pv_Absentded", OracleType.Number);
                  P13.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(P13);

                  OracleParameter p14 = new OracleParameter("Pv_Lateded", OracleType.Number);
                  p14.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p14);

                  
                  OracleParameter p15 = new OracleParameter("Pv_otherdeduc", OracleType.Number);
                  p15.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p15);

                  OracleParameter p16 = new OracleParameter("Pv_Presentdays", OracleType.Number);
                  p16.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p16);

                  OracleParameter p17 = new OracleParameter("Pv_Absentdays", OracleType.Number);
                  p17.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p17);

                  OracleParameter p18 = new OracleParameter("Pv_Leavedays", OracleType.Number);
                  p18.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p18);

                  OracleParameter p19 = new OracleParameter("Pv_Holidays", OracleType.Number);
                  p19.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p19);

                  OracleParameter p20 = new OracleParameter("Pv_Totaldays", OracleType.Number);
                  p20.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p20);

                  OracleParameter p21 = new OracleParameter("Answer", OracleType.VarChar, 200);
                  p21.Direction = ParameterDirection.Output;
                  cmd.Parameters.Add(p21);


                cmd.ExecuteNonQuery();
                //MessageBox.Show("" + cmd.Parameters["PV_House"].Value.ToString());

                
                EmpSalTxtAllowamount.Text = cmd.Parameters["Pv_Allowances_Amount"].Value.ToString();
                EmpSalTxtDeducamount.Text = cmd.Parameters["Pv_Deductions_Amount"].Value.ToString();
                EmpSalTxtNetsalary.Text = cmd.Parameters["Pv_Netsalary"].Value.ToString();
                EmpSalTxtBasicpay.Text = cmd.Parameters["Pv_Basicpay"].Value.ToString();
                //Allowances.
                EmpSalTxtHouseallow.Text = cmd.Parameters["Pv_house"].Value.ToString();
                EmpSalTxtMedicalallow.Text = cmd.Parameters["PV_medical"].Value.ToString();
                EmpSalTxtConveyanceallow.Text = cmd.Parameters["PV_convey"].Value.ToString();
                EmpSalTxtPhoneallow.Text = cmd.Parameters["PV_Phone"].Value.ToString();
                EmpSalTxtLatesittingallow.Text = cmd.Parameters["Pv_latesit"].Value.ToString();
                EmpSalTxtOtherallow.Text = cmd.Parameters["Pv_otherallow"].Value.ToString();

                EmpSalTxtLatededuc.Text = cmd.Parameters["Pv_lateded"].Value.ToString();
                EmpSalTxtAdvadeduc.Text = cmd.Parameters["Pv_Absentded"].Value.ToString();
                EmpSalTxtOtherdeduc.Text = cmd.Parameters["Pv_otherdeduc"].Value.ToString();


                
            }
            catch (OracleException ex) { MessageBox.Show(ex.Message, "Oracle Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                gl_dbaseclsobj.cls_con.Close();
            }
        }
        bool rfdisp_fun(int id,bool ck)
        {//Bool mention run function for what purpose, For Rfid search or For Formid search.
            string srch_rftagid = null;
            if(ck)
            {//This is for RFID.
                srch_rftagid = "Select * From Employee Where Rfid_id=" + id + " And Formstatus_id=6";
            }
            else
            {//This for Formid.
                srch_rftagid = "Select * From Employee Where Form_id=" + id + " And Formstatus_id=6";
            }           
            gl_dbaseclsobj.Show_Datafun(srch_rftagid, 0);            
            if (gl_dbaseclsobj.cls_dr.Read())
            {//If record found then Display into particular fields.
                MgmGenGroupLeave.Show();
                MgmGenCkbSetleave.Checked = true;
                MgmGenTxtFormid.Text=gl_dbaseclsobj.cls_dr.GetDecimal(0).ToString();
                MgmGenTxtName.Text = gl_dbaseclsobj.cls_dr.GetString(1);
                MgmGenTxtBps.Text = gl_dbaseclsobj.cls_dr.GetDecimal(20).ToString();
                

                {   //Picture loading code.
                    OracleLob blob = gl_dbaseclsobj.cls_dr.GetOracleLob(29);      //Column # .      Emp_image
                    Byte[] BLOBData = new Byte[blob.Length];
                    //Read blob data into byte array
                    int i = blob.Read(BLOBData, 0, System.Convert.ToInt32(blob.Length));
                    //Get the primitive byte data into in-memory data stream
                    MemoryStream stmBLOBData = new MemoryStream(BLOBData);
                    //LOADING INTO PICBOX1                            
                    MgmGenPicture.Image = Image.FromStream(stmBLOBData);
                    //MessageBox.Show("Now i'm reseting it.");
                    MgmGenPicture.SizeMode = PictureBoxSizeMode.StretchImage;
                }
                gl_dbaseclsobj.Show_Datafun(srch_rftagid, 1);//Now close database connection.
                //Now here below Load Department name.
                

                if (ck)
                {//This is for RFID.
                    gl_dbaseclsobj.Dml_UpdateAdapterfun("Select Dept_name from Department Where dept_id=(Select dept_id from Employee where Rfid_id='" + id + "')");
                }
                else
                {//This for Formid.
                    gl_dbaseclsobj.Dml_UpdateAdapterfun("Select Dept_name from Department Where dept_id=(Select dept_id from Employee where Form_id='" + id + "')");
                }  


                MgmGenTxtDepartment.Text = gl_dbaseclsobj.cls_dataset.Tables[0].Rows[0][0].ToString().ToUpper();
                return true;
            }//end if
            else
            {
                MessageBox.Show("Sorry No data found.", "Record Not Exist", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }



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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Admin));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.BtnLogout = new System.Windows.Forms.Button();
            this.TabAdmin = new System.Windows.Forms.TabControl();
            this.TabMgm = new System.Windows.Forms.TabPage();
            this.Management = new System.Windows.Forms.TabControl();
            this.TabDepartments = new System.Windows.Forms.TabPage();
            this.richTextBox7 = new System.Windows.Forms.Label();
            this.MgmDeptGroupDetailvu = new System.Windows.Forms.GroupBox();
            this.MgmDeptGridvu = new System.Windows.Forms.DataGridView();
            this.MgmDeptLblDeptscaleview = new System.Windows.Forms.Label();
            this.MgmDeptGroupoption = new System.Windows.Forms.GroupBox();
            this.MgmDeptLblDeptScalechoose = new System.Windows.Forms.Label();
            this.MgmDeptRbJobscales = new System.Windows.Forms.RadioButton();
            this.MgmDeptRbDept = new System.Windows.Forms.RadioButton();
            this.MgmDeptBtnAddrow = new System.Windows.Forms.Button();
            this.MgmDeptBtnEdit = new System.Windows.Forms.Button();
            this.MgmDeptGroupadddeptscale = new System.Windows.Forms.GroupBox();
            this.MgmDeptLblMinbpay = new System.Windows.Forms.Label();
            this.MgmDeptLblJobid = new System.Windows.Forms.Label();
            this.MgmDeptTxtDname = new System.Windows.Forms.TextBox();
            this.MgmDeptTxtMinbpay = new System.Windows.Forms.TextBox();
            this.MgmDeptLblOtherdeduc = new System.Windows.Forms.Label();
            this.MgmDeptTxtOtherdeduc = new System.Windows.Forms.TextBox();
            this.MgmDeptLblLatededuc = new System.Windows.Forms.Label();
            this.MgmDeptTxtLatededuc = new System.Windows.Forms.TextBox();
            this.MgmDeptLblLatesitallow = new System.Windows.Forms.Label();
            this.MgmDeptTxtLatesitallow = new System.Windows.Forms.TextBox();
            this.MgmDeptBtnDelete = new System.Windows.Forms.Button();
            this.MgmDeptBtnCancel = new System.Windows.Forms.Button();
            this.MgmDeptLblOtherallow = new System.Windows.Forms.Label();
            this.MgmDeptTxtOtherallow = new System.Windows.Forms.TextBox();
            this.MgmDeptLblPhoneallow = new System.Windows.Forms.Label();
            this.MgmDeptTxtPhoneallow = new System.Windows.Forms.TextBox();
            this.MgmDeptLblConvallow = new System.Windows.Forms.Label();
            this.MgmDeptTxtConvallow = new System.Windows.Forms.TextBox();
            this.MgmDeptLblMedicalallow = new System.Windows.Forms.Label();
            this.MgmDeptTxtMedicalallow = new System.Windows.Forms.TextBox();
            this.MgmDeptLblHouserent = new System.Windows.Forms.Label();
            this.MgmDeptTxtHouserent = new System.Windows.Forms.TextBox();
            this.MgmDeptLblPerIncr = new System.Windows.Forms.Label();
            this.MgmDeptLblDeptScaleadd = new System.Windows.Forms.Label();
            this.MgmDeptTxtPerIncr = new System.Windows.Forms.TextBox();
            this.MgmDeptLblDid = new System.Windows.Forms.Label();
            this.MgmDeptLblMaxbpay = new System.Windows.Forms.Label();
            this.MgmDeptLblDname = new System.Windows.Forms.Label();
            this.MgmDeptLblMgrid = new System.Windows.Forms.Label();
            this.MgmDeptTxtMaxbpay = new System.Windows.Forms.TextBox();
            this.MgmDeptBtnSave = new System.Windows.Forms.Button();
            this.MgmDeptTxtMgrid = new System.Windows.Forms.TextBox();
            this.MgmDeptTxtDid = new System.Windows.Forms.TextBox();
            this.MgmDeptLblBps = new System.Windows.Forms.Label();
            this.MgmDeptTxtBps = new System.Windows.Forms.TextBox();
            this.MgmDeptTxtJobid = new System.Windows.Forms.TextBox();
            this.MgmDeptBtnRefresh = new System.Windows.Forms.Button();
            this.TabJob = new System.Windows.Forms.TabPage();
            this.LblJob_update = new System.Windows.Forms.Label();
            this.MgmJobGroupSelectedemp = new System.Windows.Forms.GroupBox();
            this.MgmJobGroupeEmpdetail = new System.Windows.Forms.GroupBox();
            this.MgmJobCmb1Dept = new System.Windows.Forms.ComboBox();
            this.MgmJobLblJoindate = new System.Windows.Forms.Label();
            this.MgmJobLblJobid = new System.Windows.Forms.Label();
            this.MgmJobCtlHiredate = new System.Windows.Forms.DateTimePicker();
            this.MgmJobLbl1Dept = new System.Windows.Forms.Label();
            this.MgmJobLblRftag = new System.Windows.Forms.Label();
            this.MgmJobTxtName = new System.Windows.Forms.TextBox();
            this.MgmJobCmbEmptype = new System.Windows.Forms.ComboBox();
            this.MgmJobLblName = new System.Windows.Forms.Label();
            this.MgmJobLblEmptype = new System.Windows.Forms.Label();
            this.MgmJobTxtRftag = new System.Windows.Forms.TextBox();
            this.MgmJobCmb0Jobid = new System.Windows.Forms.ComboBox();
            this.MgmJobTxtJobid = new System.Windows.Forms.TextBox();
            this.MgmJobTxtPermotedby = new System.Windows.Forms.TextBox();
            this.MgmJobLblPermotedby = new System.Windows.Forms.Label();
            this.MgmJobBtnSaverec = new System.Windows.Forms.Button();
            this.MgmJobBtnCancel = new System.Windows.Forms.Button();
            this.MgmJobTxtFormid1 = new System.Windows.Forms.TextBox();
            this.MgmJobLblFormid = new System.Windows.Forms.Label();
            this.MgmJobPicbox = new System.Windows.Forms.PictureBox();
            this.MgmJobLblPic = new System.Windows.Forms.Label();
            this.MgmJobCkbPermotion = new System.Windows.Forms.CheckBox();
            this.MgmJobGroupRestricted = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.MgmJobGroupRestrictedpermote = new System.Windows.Forms.GroupBox();
            this.label25 = new System.Windows.Forms.Label();
            this.label29 = new System.Windows.Forms.Label();
            this.pictureBox6 = new System.Windows.Forms.PictureBox();
            this.MgmJobPanelPermotion = new System.Windows.Forms.Panel();
            this.MgmJobCmb3Dept = new System.Windows.Forms.ComboBox();
            this.MgmJobLblComments = new System.Windows.Forms.Label();
            this.MgmJobLbl1Jobid = new System.Windows.Forms.Label();
            this.MgmJobLbl2Dept = new System.Windows.Forms.Label();
            this.MgmJobTxtComments = new System.Windows.Forms.TextBox();
            this.MgmJobCmb1Bps = new System.Windows.Forms.ComboBox();
            this.MgmJobCmb1Jobid = new System.Windows.Forms.ComboBox();
            this.MgmJobLblBps1 = new System.Windows.Forms.Label();
            this.MgmJobGroupSearch = new System.Windows.Forms.GroupBox();
            this.MgmJobLblCaution = new System.Windows.Forms.Label();
            this.MgmJobTxtformid0 = new System.Windows.Forms.TextBox();
            this.MgmJobCmb0Dept = new System.Windows.Forms.ComboBox();
            this.MgmJobCkbFormid = new System.Windows.Forms.CheckBox();
            this.MgmJobBtnSearch = new System.Windows.Forms.Button();
            this.MgmJobCkbdDeptwise = new System.Windows.Forms.CheckBox();
            this.MgmJobCkbJobisnull = new System.Windows.Forms.CheckBox();
            this.MgmJobBtnEdit = new System.Windows.Forms.Button();
            this.MgmJobGridvu = new System.Windows.Forms.DataGridView();
            this.TabAttend = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.MgmAttGroupSearch = new System.Windows.Forms.GroupBox();
            this.MgmAttCkbFormid = new System.Windows.Forms.CheckBox();
            this.MgmAttLblDatetime = new System.Windows.Forms.Label();
            this.MgmAttDTP = new System.Windows.Forms.DateTimePicker();
            this.MgmAttBtnSearch = new System.Windows.Forms.Button();
            this.MgmAttCmbDept = new System.Windows.Forms.ComboBox();
            this.MgmAttBtnClear = new System.Windows.Forms.Button();
            this.MgmAttTxtFormid = new System.Windows.Forms.TextBox();
            this.MgmAttCkbRftag = new System.Windows.Forms.CheckBox();
            this.MgmAttCkbDept = new System.Windows.Forms.CheckBox();
            this.MgmAttTxtRftag = new System.Windows.Forms.TextBox();
            this.MgmAttGroupDetailview = new System.Windows.Forms.GroupBox();
            this.MgmAttGridAttdetailvu = new System.Windows.Forms.DataGridView();
            this.MgmAttGroupEmpatt = new System.Windows.Forms.GroupBox();
            this.MgmAttGridAttendVu = new System.Windows.Forms.DataGridView();
            this.Detail_View = new System.Windows.Forms.DataGridViewButtonColumn();
            this.TabMgmGeneral = new System.Windows.Forms.TabPage();
            this.label23 = new System.Windows.Forms.Label();
            this.MgmGenGroup = new System.Windows.Forms.GroupBox();
            this.MgmGenLblFormid = new System.Windows.Forms.Label();
            this.MgmGenTxtFormid = new System.Windows.Forms.TextBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.MgmGenTxtSearchbyformid = new System.Windows.Forms.TextBox();
            this.MgmGenTxtSearchbyrf = new System.Windows.Forms.TextBox();
            this.MgmGenBtnClear = new System.Windows.Forms.Button();
            this.label36 = new System.Windows.Forms.Label();
            this.MgmGenBtnSearch = new System.Windows.Forms.Button();
            this.MgmGenCkbSearchbyrf = new System.Windows.Forms.CheckBox();
            this.MgmGenCkbSearchbyformid = new System.Windows.Forms.CheckBox();
            this.MgmGenGroupLeave = new System.Windows.Forms.GroupBox();
            this.MgmGenGroupBoxStatus = new System.Windows.Forms.GroupBox();
            this.MgmGenLblSelectdate = new System.Windows.Forms.Label();
            this.label34 = new System.Windows.Forms.Label();
            this.MgmGenLblStatus = new System.Windows.Forms.Label();
            this.MgmGenCmbStatus = new System.Windows.Forms.ComboBox();
            this.MgmGenDtTime = new System.Windows.Forms.DateTimePicker();
            this.MgmGenBtnAllow = new System.Windows.Forms.Button();
            this.MgmGenGroupBoxLeave = new System.Windows.Forms.GroupBox();
            this.label24 = new System.Windows.Forms.Label();
            this.MgmGenTxtLeavedates = new System.Windows.Forms.TextBox();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.MgmGenCtlDate1 = new System.Windows.Forms.DateTimePicker();
            this.RecepLblLeavedates = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.label27 = new System.Windows.Forms.Label();
            this.MgmGenGroupboxSetholiday = new System.Windows.Forms.GroupBox();
            this.label30 = new System.Windows.Forms.Label();
            this.MgmGenLblReason = new System.Windows.Forms.Label();
            this.MgmGenCmbReason = new System.Windows.Forms.ComboBox();
            this.MgmGenDttimeHoliday = new System.Windows.Forms.DateTimePicker();
            this.MgmGenLblSelecteddate = new System.Windows.Forms.Label();
            this.MgmGenCkbSetstatus = new System.Windows.Forms.CheckBox();
            this.MgmGenCkbSetholidays = new System.Windows.Forms.CheckBox();
            this.MgmGenCkbSetleave = new System.Windows.Forms.CheckBox();
            this.MgmGenLblDepartment = new System.Windows.Forms.Label();
            this.MgmGenTxtDepartment = new System.Windows.Forms.TextBox();
            this.MgmGenPicture = new System.Windows.Forms.PictureBox();
            this.MgmGenTxtBps = new System.Windows.Forms.TextBox();
            this.MgmGenLblBps = new System.Windows.Forms.Label();
            this.MgmGenLblName = new System.Windows.Forms.Label();
            this.MgmGenTxtName = new System.Windows.Forms.TextBox();
            this.TabEmp = new System.Windows.Forms.TabPage();
            this.button4 = new System.Windows.Forms.Button();
            this.Employees = new System.Windows.Forms.TabControl();
            this.TabGeneral = new System.Windows.Forms.TabPage();
            this.label35 = new System.Windows.Forms.Label();
            this.EmpGenTxtUpdateby = new System.Windows.Forms.TextBox();
            this.EmpGenLblUpdateby = new System.Windows.Forms.Label();
            this.EmpGenLblRfid = new System.Windows.Forms.Label();
            this.EmpGenTxtRfid = new System.Windows.Forms.TextBox();
            this.EmpGenLblHiredate = new System.Windows.Forms.Label();
            this.EmpGenTxtHiredate = new System.Windows.Forms.TextBox();
            this.EmpGenLblBnkbranch = new System.Windows.Forms.Label();
            this.EmpGenTxtBnkbranch = new System.Windows.Forms.TextBox();
            this.EmpGenLblBnkacc = new System.Windows.Forms.Label();
            this.EmpGenTxtBnkacc = new System.Windows.Forms.TextBox();
            this.EmpGenGroupSearch = new System.Windows.Forms.GroupBox();
            this.EmpGenBtnEdit = new System.Windows.Forms.Button();
            this.EmpGenBtnCancel = new System.Windows.Forms.Button();
            this.EmpGenLblSearchoption = new System.Windows.Forms.Label();
            this.EmpGenCkbCnic = new System.Windows.Forms.CheckBox();
            this.EmpGenBtnSearch = new System.Windows.Forms.Button();
            this.EmpGenBtnSave = new System.Windows.Forms.Button();
            this.EmpGenCkbRfid = new System.Windows.Forms.CheckBox();
            this.EmpGenLblSearchrec = new System.Windows.Forms.Label();
            this.EmpGenCkbFormid = new System.Windows.Forms.CheckBox();
            this.EmpGenTxtSearchrec = new System.Windows.Forms.TextBox();
            this.EmpGenLblPic = new System.Windows.Forms.Label();
            this.EmpGenLblTempadd = new System.Windows.Forms.Label();
            this.EmpGenTxtTempadd = new System.Windows.Forms.TextBox();
            this.EmpGenLblPermadd = new System.Windows.Forms.Label();
            this.EmpGenTxtPermadd = new System.Windows.Forms.TextBox();
            this.EmpGenLblPhno = new System.Windows.Forms.Label();
            this.EmpGenTxtPhno = new System.Windows.Forms.TextBox();
            this.EmpGenLblEmail = new System.Windows.Forms.Label();
            this.EmpGenTxtEmail = new System.Windows.Forms.TextBox();
            this.EmpGenLblDomicile = new System.Windows.Forms.Label();
            this.EmpGenTxtDomicile = new System.Windows.Forms.TextBox();
            this.EmpGenLblReligion = new System.Windows.Forms.Label();
            this.EmpGenTxtReligion = new System.Windows.Forms.TextBox();
            this.EmpGenLblDob = new System.Windows.Forms.Label();
            this.EmpGenLblCnic = new System.Windows.Forms.Label();
            this.EmpGenLblFname = new System.Windows.Forms.Label();
            this.EmpGenLblEname = new System.Windows.Forms.Label();
            this.EmpGenTxtCnic = new System.Windows.Forms.TextBox();
            this.EmpGenTxtDob = new System.Windows.Forms.TextBox();
            this.EmpGenTxtFname = new System.Windows.Forms.TextBox();
            this.EmpGentxtEname = new System.Windows.Forms.TextBox();
            this.EmpGenBtnBrowse = new System.Windows.Forms.Button();
            this.EmpGenPicbox = new System.Windows.Forms.PictureBox();
            this.EmpGenGroupEmpaccount = new System.Windows.Forms.GroupBox();
            this.EmpGenTxtPwd = new System.Windows.Forms.TextBox();
            this.EmpGenLblPwd = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.EmpGenLblAccstatus = new System.Windows.Forms.Label();
            this.EmpGenRdbDisable = new System.Windows.Forms.RadioButton();
            this.EmpGenRdbEnable = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.EmpGenTxtLogname = new System.Windows.Forms.TextBox();
            this.EmpGenLblLogname = new System.Windows.Forms.Label();
            this.TabMonthlySettings = new System.Windows.Forms.TabPage();
            this.label33 = new System.Windows.Forms.Label();
            this.EmpMonGroupSet0 = new System.Windows.Forms.GroupBox();
            this.EmpMonGroup1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.EmpMonCkbPhoneall = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbOtherded = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbLateded = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbHouserent = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbMedallowance = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbOtherall = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbConvAllowance = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbLatesitall = new System.Windows.Forms.CheckBox();
            this.EmpMonTxtName = new System.Windows.Forms.TextBox();
            this.EmpMonLblAllowtotal = new System.Windows.Forms.Label();
            this.EmpMonLblJobid = new System.Windows.Forms.Label();
            this.EmpMonTxtDedtotal = new System.Windows.Forms.TextBox();
            this.EmpMonTxtJobid = new System.Windows.Forms.TextBox();
            this.EmpMonTxtAllowtotal = new System.Windows.Forms.TextBox();
            this.EmpMonBtnApply = new System.Windows.Forms.Button();
            this.EmpMonLblName = new System.Windows.Forms.Label();
            this.EmpMonLblDedtotal = new System.Windows.Forms.Label();
            this.EmpMonTxtTotalamount = new System.Windows.Forms.TextBox();
            this.EmpMonBtnCancel0 = new System.Windows.Forms.Button();
            this.EmpMonTxtBps = new System.Windows.Forms.TextBox();
            this.EmpMonLblBps = new System.Windows.Forms.Label();
            this.EmpMonLblTotalamount = new System.Windows.Forms.Label();
            this.EmpMonGridvu = new System.Windows.Forms.DataGridView();
            this.Edit = new System.Windows.Forms.DataGridViewButtonColumn();
            this.EmpMonGroupsearch = new System.Windows.Forms.GroupBox();
            this.EmpMonTxtFormid = new System.Windows.Forms.TextBox();
            this.EmpMonTxtRfid = new System.Windows.Forms.TextBox();
            this.EmpMonCkbDepartment = new System.Windows.Forms.CheckBox();
            this.EmpMonCbmDepartment = new System.Windows.Forms.ComboBox();
            this.EmpMonBtnClear = new System.Windows.Forms.Button();
            this.EmpMonLblSearchoption = new System.Windows.Forms.Label();
            this.EmpMonBtnSearch = new System.Windows.Forms.Button();
            this.EmpMonCkbByrfid = new System.Windows.Forms.CheckBox();
            this.EmpMonCkbByformid = new System.Windows.Forms.CheckBox();
            this.EmpMonGroup2 = new System.Windows.Forms.GroupBox();
            this.label28 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.EmpMonBtnOk = new System.Windows.Forms.Button();
            this.TabSalary = new System.Windows.Forms.TabPage();
            this.label32 = new System.Windows.Forms.Label();
            this.EmpSalGroupSalary = new System.Windows.Forms.GroupBox();
            this.EmpSalGroupAlloDeductdetail = new System.Windows.Forms.GroupBox();
            this.EmpSalLblHouseallow = new System.Windows.Forms.Label();
            this.EmpSalLblOtherdeduc = new System.Windows.Forms.Label();
            this.EmpSalLblAbsentdeduc = new System.Windows.Forms.Label();
            this.EmpSalLblLatededuc = new System.Windows.Forms.Label();
            this.EmpSalTxtOtherdeduc = new System.Windows.Forms.TextBox();
            this.EmpSalTxtAdvadeduc = new System.Windows.Forms.TextBox();
            this.EmpSalTxtLatededuc = new System.Windows.Forms.TextBox();
            this.EmpSalLblOtherallow = new System.Windows.Forms.Label();
            this.EmpSalLblLatesittingallow = new System.Windows.Forms.Label();
            this.EmpSalLblPhoneallow = new System.Windows.Forms.Label();
            this.EmpSalLblConvallow = new System.Windows.Forms.Label();
            this.EmpSalLblMedicalallow = new System.Windows.Forms.Label();
            this.EmpSalTxtOtherallow = new System.Windows.Forms.TextBox();
            this.EmpSalTxtConveyanceallow = new System.Windows.Forms.TextBox();
            this.EmpSalTxtLatesittingallow = new System.Windows.Forms.TextBox();
            this.EmpSalTxtMedicalallow = new System.Windows.Forms.TextBox();
            this.EmpSalTxtPhoneallow = new System.Windows.Forms.TextBox();
            this.EmpSalTxtHouseallow = new System.Windows.Forms.TextBox();
            this.EmpSalGroupSalarydetail = new System.Windows.Forms.GroupBox();
            this.EmpSalCkbAllowamount = new System.Windows.Forms.CheckBox();
            this.EmpSalLblBasicpay = new System.Windows.Forms.Label();
            this.EmpSalTxtBasicpay = new System.Windows.Forms.TextBox();
            this.EmpSalTxtAllowamount = new System.Windows.Forms.TextBox();
            this.EmpSalLblNetsalary = new System.Windows.Forms.Label();
            this.EmpSalTxtDeducamount = new System.Windows.Forms.TextBox();
            this.EmpSalCkbDeducamount = new System.Windows.Forms.CheckBox();
            this.EmpSalLblFollow0 = new System.Windows.Forms.Label();
            this.EmpSalTxtNetsalary = new System.Windows.Forms.TextBox();
            this.EmpSalGroupdaysdetail = new System.Windows.Forms.GroupBox();
            this.EmpSalTxtTotaldays = new System.Windows.Forms.TextBox();
            this.EmpSalLblTotaldays = new System.Windows.Forms.Label();
            this.EmpSalLblFollow1 = new System.Windows.Forms.Label();
            this.EmpSalLblPresentdays = new System.Windows.Forms.Label();
            this.EmpSalTxtPresentdays = new System.Windows.Forms.TextBox();
            this.EmpSalTxtAbsentdays = new System.Windows.Forms.TextBox();
            this.EmpSalLblAbsentdays = new System.Windows.Forms.Label();
            this.EmpSalLblLeavedays = new System.Windows.Forms.Label();
            this.EmpSalTxtLeavedays = new System.Windows.Forms.TextBox();
            this.EmpSalTxtHolidays = new System.Windows.Forms.TextBox();
            this.EmpSalLblHolidays = new System.Windows.Forms.Label();
            this.EmpSalPicbox = new System.Windows.Forms.PictureBox();
            this.EmpSalTxtName = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.EmpSalDatetime = new System.Windows.Forms.DateTimePicker();
            this.EmpSalCkbRftag = new System.Windows.Forms.CheckBox();
            this.EmpSalCkbFormid = new System.Windows.Forms.CheckBox();
            this.EmpSalTxtFormid = new System.Windows.Forms.TextBox();
            this.EmpSalBtnSearch = new System.Windows.Forms.Button();
            this.EmpSalTxtRftag = new System.Windows.Forms.TextBox();
            this.EmpSalLblFollow2 = new System.Windows.Forms.Label();
            this.EmpSalCkbDept = new System.Windows.Forms.CheckBox();
            this.EmpSalBtnCancel0 = new System.Windows.Forms.Button();
            this.EmpSalCombDept = new System.Windows.Forms.ComboBox();
            this.EmpSalGridvu = new System.Windows.Forms.DataGridView();
            this.ViewSalaryDetail = new System.Windows.Forms.DataGridViewButtonColumn();
            this.TabAttendance = new System.Windows.Forms.TabPage();
            this.label31 = new System.Windows.Forms.Label();
            this.EmpAttGroup0 = new System.Windows.Forms.GroupBox();
            this.EmpAttBtnClear = new System.Windows.Forms.Button();
            this.EmpAttTxtSearchbyformid = new System.Windows.Forms.TextBox();
            this.EmpAttTxtSearchbyrfid = new System.Windows.Forms.TextBox();
            this.EmpAttCkbSearchbyrfid = new System.Windows.Forms.CheckBox();
            this.EmpAttCkbSearchbyformid = new System.Windows.Forms.CheckBox();
            this.EmpAttPicBox = new System.Windows.Forms.PictureBox();
            this.EmpAttGridvu = new System.Windows.Forms.DataGridView();
            this.CheckDetaill = new System.Windows.Forms.DataGridViewButtonColumn();
            this.EmAttBtnSearch = new System.Windows.Forms.Button();
            this.EmpAttCtlDatetime = new System.Windows.Forms.DateTimePicker();
            this.EmpAttGroup1 = new System.Windows.Forms.GroupBox();
            this.EmpAttGridvudetail = new System.Windows.Forms.DataGridView();
            this.TabScanner = new System.Windows.Forms.TabPage();
            this.TabDEO = new System.Windows.Forms.TabPage();
            this.TabHr = new System.Windows.Forms.TabPage();
            this.TabReception = new System.Windows.Forms.TabPage();
            this.TabReportViewer = new System.Windows.Forms.TabPage();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.TabReport_Employee = new System.Windows.Forms.TabPage();
            this.label37 = new System.Windows.Forms.Label();
            this.CrysReport_Employee = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.CryReport_Employee1 = new OA_PS_RFID.CryReport_Employee();
            this.ChoosRow = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.LblNullogin = new System.Windows.Forms.Label();
            this.LblLoginstatus = new System.Windows.Forms.Label();
            this.LblLoginame = new System.Windows.Forms.Label();
            this.LblNullname = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.textBox17 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label10 = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.textBox19 = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.textBox20 = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.textBox21 = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.textBox22 = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.textBox23 = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.textBox24 = new System.Windows.Forms.TextBox();
            this.textBox25 = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label18 = new System.Windows.Forms.Label();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.label19 = new System.Windows.Forms.Label();
            this.textBox26 = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.employeeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.managementToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.scanningProcessToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.scanningToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataEntryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.hRVerifyingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.receptonToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.indexToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.searchToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.TabAdmin.SuspendLayout();
            this.TabMgm.SuspendLayout();
            this.Management.SuspendLayout();
            this.TabDepartments.SuspendLayout();
            this.MgmDeptGroupDetailvu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmDeptGridvu)).BeginInit();
            this.MgmDeptGroupoption.SuspendLayout();
            this.MgmDeptGroupadddeptscale.SuspendLayout();
            this.TabJob.SuspendLayout();
            this.MgmJobGroupSelectedemp.SuspendLayout();
            this.MgmJobGroupeEmpdetail.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmJobPicbox)).BeginInit();
            this.MgmJobGroupRestricted.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            this.MgmJobGroupRestrictedpermote.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox6)).BeginInit();
            this.MgmJobPanelPermotion.SuspendLayout();
            this.MgmJobGroupSearch.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmJobGridvu)).BeginInit();
            this.TabAttend.SuspendLayout();
            this.MgmAttGroupSearch.SuspendLayout();
            this.MgmAttGroupDetailview.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmAttGridAttdetailvu)).BeginInit();
            this.MgmAttGroupEmpatt.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmAttGridAttendVu)).BeginInit();
            this.TabMgmGeneral.SuspendLayout();
            this.MgmGenGroup.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.MgmGenGroupLeave.SuspendLayout();
            this.MgmGenGroupBoxStatus.SuspendLayout();
            this.MgmGenGroupBoxLeave.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.MgmGenGroupboxSetholiday.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmGenPicture)).BeginInit();
            this.TabEmp.SuspendLayout();
            this.Employees.SuspendLayout();
            this.TabGeneral.SuspendLayout();
            this.EmpGenGroupSearch.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpGenPicbox)).BeginInit();
            this.EmpGenGroupEmpaccount.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.TabMonthlySettings.SuspendLayout();
            this.EmpMonGroupSet0.SuspendLayout();
            this.EmpMonGroup1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpMonGridvu)).BeginInit();
            this.EmpMonGroupsearch.SuspendLayout();
            this.EmpMonGroup2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.TabSalary.SuspendLayout();
            this.EmpSalGroupSalary.SuspendLayout();
            this.EmpSalGroupAlloDeductdetail.SuspendLayout();
            this.EmpSalGroupSalarydetail.SuspendLayout();
            this.EmpSalGroupdaysdetail.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpSalPicbox)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpSalGridvu)).BeginInit();
            this.TabAttendance.SuspendLayout();
            this.EmpAttGroup0.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpAttPicBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.EmpAttGridvu)).BeginInit();
            this.EmpAttGroup1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpAttGridvudetail)).BeginInit();
            this.TabReportViewer.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.TabReport_Employee.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnLogout
            // 
            this.BtnLogout.Location = new System.Drawing.Point(893, 604);
            this.BtnLogout.Name = "BtnLogout";
            this.BtnLogout.Size = new System.Drawing.Size(74, 23);
            this.BtnLogout.TabIndex = 0;
            this.BtnLogout.Text = "Log out";
            this.BtnLogout.UseVisualStyleBackColor = true;
            this.BtnLogout.Click += new System.EventHandler(this.BtnLogout_Click);
            // 
            // TabAdmin
            // 
            this.TabAdmin.Controls.Add(this.TabMgm);
            this.TabAdmin.Controls.Add(this.TabEmp);
            this.TabAdmin.Controls.Add(this.TabScanner);
            this.TabAdmin.Controls.Add(this.TabDEO);
            this.TabAdmin.Controls.Add(this.TabHr);
            this.TabAdmin.Controls.Add(this.TabReception);
            this.TabAdmin.Controls.Add(this.TabReportViewer);
            this.TabAdmin.Location = new System.Drawing.Point(14, 28);
            this.TabAdmin.Name = "TabAdmin";
            this.TabAdmin.SelectedIndex = 0;
            this.TabAdmin.Size = new System.Drawing.Size(1002, 656);
            this.TabAdmin.TabIndex = 1;
            this.TabAdmin.Selected += new System.Windows.Forms.TabControlEventHandler(this.TabAdmin_Selected);
            // 
            // TabMgm
            // 
            this.TabMgm.BackColor = System.Drawing.SystemColors.Desktop;
            this.TabMgm.Controls.Add(this.Management);
            this.TabMgm.Controls.Add(this.BtnLogout);
            this.TabMgm.Location = new System.Drawing.Point(4, 22);
            this.TabMgm.Name = "TabMgm";
            this.TabMgm.Padding = new System.Windows.Forms.Padding(3);
            this.TabMgm.Size = new System.Drawing.Size(994, 630);
            this.TabMgm.TabIndex = 1;
            this.TabMgm.Text = "Management";
            // 
            // Management
            // 
            this.Management.Controls.Add(this.TabDepartments);
            this.Management.Controls.Add(this.TabJob);
            this.Management.Controls.Add(this.TabAttend);
            this.Management.Controls.Add(this.TabMgmGeneral);
            this.Management.Location = new System.Drawing.Point(22, 21);
            this.Management.Name = "Management";
            this.Management.SelectedIndex = 0;
            this.Management.Size = new System.Drawing.Size(945, 577);
            this.Management.TabIndex = 12;
            this.Management.Selected += new System.Windows.Forms.TabControlEventHandler(this.Management_Selected);
            // 
            // TabDepartments
            // 
            this.TabDepartments.BackColor = System.Drawing.Color.AliceBlue;
            this.TabDepartments.Controls.Add(this.richTextBox7);
            this.TabDepartments.Controls.Add(this.MgmDeptGroupDetailvu);
            this.TabDepartments.Controls.Add(this.MgmDeptGroupoption);
            this.TabDepartments.Controls.Add(this.MgmDeptBtnAddrow);
            this.TabDepartments.Controls.Add(this.MgmDeptBtnEdit);
            this.TabDepartments.Controls.Add(this.MgmDeptGroupadddeptscale);
            this.TabDepartments.Controls.Add(this.MgmDeptBtnRefresh);
            this.TabDepartments.Location = new System.Drawing.Point(4, 22);
            this.TabDepartments.Name = "TabDepartments";
            this.TabDepartments.Padding = new System.Windows.Forms.Padding(3);
            this.TabDepartments.Size = new System.Drawing.Size(937, 551);
            this.TabDepartments.TabIndex = 0;
            this.TabDepartments.Text = "Departments";
            // 
            // richTextBox7
            // 
            this.richTextBox7.BackColor = System.Drawing.Color.WhiteSmoke;
            this.richTextBox7.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox7.Location = new System.Drawing.Point(0, 0);
            this.richTextBox7.Name = "richTextBox7";
            this.richTextBox7.Size = new System.Drawing.Size(937, 24);
            this.richTextBox7.TabIndex = 78;
            this.richTextBox7.Text = "                                                             MIA SCALE VIEW";
            // 
            // MgmDeptGroupDetailvu
            // 
            this.MgmDeptGroupDetailvu.Controls.Add(this.MgmDeptGridvu);
            this.MgmDeptGroupDetailvu.Controls.Add(this.MgmDeptLblDeptscaleview);
            this.MgmDeptGroupDetailvu.Location = new System.Drawing.Point(12, 151);
            this.MgmDeptGroupDetailvu.Name = "MgmDeptGroupDetailvu";
            this.MgmDeptGroupDetailvu.Size = new System.Drawing.Size(914, 232);
            this.MgmDeptGroupDetailvu.TabIndex = 15;
            this.MgmDeptGroupDetailvu.TabStop = false;
            this.MgmDeptGroupDetailvu.Text = "Detailed View";
            // 
            // MgmDeptGridvu
            // 
            this.MgmDeptGridvu.AllowUserToAddRows = false;
            this.MgmDeptGridvu.AllowUserToDeleteRows = false;
            this.MgmDeptGridvu.AllowUserToOrderColumns = true;
            this.MgmDeptGridvu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MgmDeptGridvu.Location = new System.Drawing.Point(9, 37);
            this.MgmDeptGridvu.Name = "MgmDeptGridvu";
            this.MgmDeptGridvu.ReadOnly = true;
            this.MgmDeptGridvu.Size = new System.Drawing.Size(898, 184);
            this.MgmDeptGridvu.TabIndex = 7;
            // 
            // MgmDeptLblDeptscaleview
            // 
            this.MgmDeptLblDeptscaleview.AutoSize = true;
            this.MgmDeptLblDeptscaleview.Location = new System.Drawing.Point(8, 17);
            this.MgmDeptLblDeptscaleview.Name = "MgmDeptLblDeptscaleview";
            this.MgmDeptLblDeptscaleview.Size = new System.Drawing.Size(119, 13);
            this.MgmDeptLblDeptscaleview.TabIndex = 11;
            this.MgmDeptLblDeptscaleview.Text = "Job Scale detailed view";
            // 
            // MgmDeptGroupoption
            // 
            this.MgmDeptGroupoption.BackColor = System.Drawing.Color.AliceBlue;
            this.MgmDeptGroupoption.Controls.Add(this.MgmDeptLblDeptScalechoose);
            this.MgmDeptGroupoption.Controls.Add(this.MgmDeptRbJobscales);
            this.MgmDeptGroupoption.Controls.Add(this.MgmDeptRbDept);
            this.MgmDeptGroupoption.Location = new System.Drawing.Point(12, 43);
            this.MgmDeptGroupoption.Name = "MgmDeptGroupoption";
            this.MgmDeptGroupoption.Size = new System.Drawing.Size(261, 103);
            this.MgmDeptGroupoption.TabIndex = 14;
            this.MgmDeptGroupoption.TabStop = false;
            this.MgmDeptGroupoption.Text = "Options";
            // 
            // MgmDeptLblDeptScalechoose
            // 
            this.MgmDeptLblDeptScalechoose.AutoSize = true;
            this.MgmDeptLblDeptScalechoose.BackColor = System.Drawing.Color.White;
            this.MgmDeptLblDeptScalechoose.Dock = System.Windows.Forms.DockStyle.Top;
            this.MgmDeptLblDeptScalechoose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblDeptScalechoose.Location = new System.Drawing.Point(3, 16);
            this.MgmDeptLblDeptScalechoose.Name = "MgmDeptLblDeptScalechoose";
            this.MgmDeptLblDeptScalechoose.Size = new System.Drawing.Size(245, 13);
            this.MgmDeptLblDeptScalechoose.TabIndex = 15;
            this.MgmDeptLblDeptScalechoose.Text = "Following option can be selected for detailed view.";
            // 
            // MgmDeptRbJobscales
            // 
            this.MgmDeptRbJobscales.AutoSize = true;
            this.MgmDeptRbJobscales.Checked = true;
            this.MgmDeptRbJobscales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptRbJobscales.Location = new System.Drawing.Point(30, 44);
            this.MgmDeptRbJobscales.Name = "MgmDeptRbJobscales";
            this.MgmDeptRbJobscales.Size = new System.Drawing.Size(87, 17);
            this.MgmDeptRbJobscales.TabIndex = 1;
            this.MgmDeptRbJobscales.TabStop = true;
            this.MgmDeptRbJobscales.Text = "Job Scales";
            this.MgmDeptRbJobscales.UseVisualStyleBackColor = true;
            this.MgmDeptRbJobscales.CheckedChanged += new System.EventHandler(this.MgmDeptRbJobscales_CheckedChanged);
            // 
            // MgmDeptRbDept
            // 
            this.MgmDeptRbDept.AutoSize = true;
            this.MgmDeptRbDept.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptRbDept.Location = new System.Drawing.Point(30, 67);
            this.MgmDeptRbDept.Name = "MgmDeptRbDept";
            this.MgmDeptRbDept.Size = new System.Drawing.Size(96, 17);
            this.MgmDeptRbDept.TabIndex = 0;
            this.MgmDeptRbDept.Text = "Departments";
            this.MgmDeptRbDept.UseVisualStyleBackColor = true;
            // 
            // MgmDeptBtnAddrow
            // 
            this.MgmDeptBtnAddrow.Location = new System.Drawing.Point(443, 122);
            this.MgmDeptBtnAddrow.Name = "MgmDeptBtnAddrow";
            this.MgmDeptBtnAddrow.Size = new System.Drawing.Size(74, 23);
            this.MgmDeptBtnAddrow.TabIndex = 9;
            this.MgmDeptBtnAddrow.Text = "Add Row";
            this.MgmDeptBtnAddrow.UseVisualStyleBackColor = true;
            this.MgmDeptBtnAddrow.Click += new System.EventHandler(this.BtnAddrow_Click);
            // 
            // MgmDeptBtnEdit
            // 
            this.MgmDeptBtnEdit.Location = new System.Drawing.Point(363, 122);
            this.MgmDeptBtnEdit.Name = "MgmDeptBtnEdit";
            this.MgmDeptBtnEdit.Size = new System.Drawing.Size(74, 23);
            this.MgmDeptBtnEdit.TabIndex = 8;
            this.MgmDeptBtnEdit.Text = "Edit";
            this.MgmDeptBtnEdit.UseVisualStyleBackColor = true;
            this.MgmDeptBtnEdit.Click += new System.EventHandler(this.BtnEdit_Click);
            // 
            // MgmDeptGroupadddeptscale
            // 
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblMinbpay);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblJobid);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtDname);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtMinbpay);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblOtherdeduc);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtOtherdeduc);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblLatededuc);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtLatededuc);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblLatesitallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtLatesitallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptBtnDelete);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptBtnCancel);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblOtherallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtOtherallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblPhoneallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtPhoneallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblConvallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtConvallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblMedicalallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtMedicalallow);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblHouserent);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtHouserent);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblPerIncr);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblDeptScaleadd);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtPerIncr);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblDid);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblMaxbpay);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblDname);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblMgrid);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtMaxbpay);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptBtnSave);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtMgrid);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtDid);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptLblBps);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtBps);
            this.MgmDeptGroupadddeptscale.Controls.Add(this.MgmDeptTxtJobid);
            this.MgmDeptGroupadddeptscale.Location = new System.Drawing.Point(12, 386);
            this.MgmDeptGroupadddeptscale.Name = "MgmDeptGroupadddeptscale";
            this.MgmDeptGroupadddeptscale.Size = new System.Drawing.Size(914, 159);
            this.MgmDeptGroupadddeptscale.TabIndex = 13;
            this.MgmDeptGroupadddeptscale.TabStop = false;
            this.MgmDeptGroupadddeptscale.Text = "Add Department";
            // 
            // MgmDeptLblMinbpay
            // 
            this.MgmDeptLblMinbpay.AutoSize = true;
            this.MgmDeptLblMinbpay.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblMinbpay.Location = new System.Drawing.Point(318, 47);
            this.MgmDeptLblMinbpay.Name = "MgmDeptLblMinbpay";
            this.MgmDeptLblMinbpay.Size = new System.Drawing.Size(87, 13);
            this.MgmDeptLblMinbpay.TabIndex = 14;
            this.MgmDeptLblMinbpay.Text = "Min Basic Pay";
            // 
            // MgmDeptLblJobid
            // 
            this.MgmDeptLblJobid.AutoSize = true;
            this.MgmDeptLblJobid.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblJobid.Location = new System.Drawing.Point(135, 47);
            this.MgmDeptLblJobid.Name = "MgmDeptLblJobid";
            this.MgmDeptLblJobid.Size = new System.Drawing.Size(42, 13);
            this.MgmDeptLblJobid.TabIndex = 12;
            this.MgmDeptLblJobid.Text = "Job Id";
            // 
            // MgmDeptTxtDname
            // 
            this.MgmDeptTxtDname.Location = new System.Drawing.Point(313, 50);
            this.MgmDeptTxtDname.Name = "MgmDeptTxtDname";
            this.MgmDeptTxtDname.Size = new System.Drawing.Size(186, 20);
            this.MgmDeptTxtDname.TabIndex = 1;
            this.MgmDeptTxtDname.Leave += new System.EventHandler(this.MgmDeptTxtDname_Leave);
            this.MgmDeptTxtDname.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtDname_KeyPress);
            // 
            // MgmDeptTxtMinbpay
            // 
            this.MgmDeptTxtMinbpay.Location = new System.Drawing.Point(407, 43);
            this.MgmDeptTxtMinbpay.Name = "MgmDeptTxtMinbpay";
            this.MgmDeptTxtMinbpay.Size = new System.Drawing.Size(94, 20);
            this.MgmDeptTxtMinbpay.TabIndex = 13;
            this.MgmDeptTxtMinbpay.Leave += new System.EventHandler(this.MgmDeptTxtMinbpay_Leave);
            this.MgmDeptTxtMinbpay.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtMinbpay_KeyPress);
            // 
            // MgmDeptLblOtherdeduc
            // 
            this.MgmDeptLblOtherdeduc.AutoSize = true;
            this.MgmDeptLblOtherdeduc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblOtherdeduc.Location = new System.Drawing.Point(318, 104);
            this.MgmDeptLblOtherdeduc.Name = "MgmDeptLblOtherdeduc";
            this.MgmDeptLblOtherdeduc.Size = new System.Drawing.Size(100, 13);
            this.MgmDeptLblOtherdeduc.TabIndex = 36;
            this.MgmDeptLblOtherdeduc.Text = "Other Deduction";
            // 
            // MgmDeptTxtOtherdeduc
            // 
            this.MgmDeptTxtOtherdeduc.Location = new System.Drawing.Point(457, 100);
            this.MgmDeptTxtOtherdeduc.Name = "MgmDeptTxtOtherdeduc";
            this.MgmDeptTxtOtherdeduc.Size = new System.Drawing.Size(44, 20);
            this.MgmDeptTxtOtherdeduc.TabIndex = 35;
            this.MgmDeptTxtOtherdeduc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtOtherdeduc_KeyPress);
            // 
            // MgmDeptLblLatededuc
            // 
            this.MgmDeptLblLatededuc.AutoSize = true;
            this.MgmDeptLblLatededuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblLatededuc.Location = new System.Drawing.Point(167, 104);
            this.MgmDeptLblLatededuc.Name = "MgmDeptLblLatededuc";
            this.MgmDeptLblLatededuc.Size = new System.Drawing.Size(94, 13);
            this.MgmDeptLblLatededuc.TabIndex = 34;
            this.MgmDeptLblLatededuc.Text = "Late Deduction";
            // 
            // MgmDeptTxtLatededuc
            // 
            this.MgmDeptTxtLatededuc.Location = new System.Drawing.Point(267, 100);
            this.MgmDeptTxtLatededuc.Name = "MgmDeptTxtLatededuc";
            this.MgmDeptTxtLatededuc.Size = new System.Drawing.Size(45, 20);
            this.MgmDeptTxtLatededuc.TabIndex = 33;
            this.MgmDeptTxtLatededuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtLatededuc_KeyPress);
            // 
            // MgmDeptLblLatesitallow
            // 
            this.MgmDeptLblLatesitallow.AutoSize = true;
            this.MgmDeptLblLatesitallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblLatesitallow.Location = new System.Drawing.Point(701, 74);
            this.MgmDeptLblLatesitallow.Name = "MgmDeptLblLatesitallow";
            this.MgmDeptLblLatesitallow.Size = new System.Drawing.Size(134, 13);
            this.MgmDeptLblLatesitallow.TabIndex = 32;
            this.MgmDeptLblLatesitallow.Text = "Late Sitting Allowance";
            // 
            // MgmDeptTxtLatesitallow
            // 
            this.MgmDeptTxtLatesitallow.Location = new System.Drawing.Point(849, 70);
            this.MgmDeptTxtLatesitallow.Name = "MgmDeptTxtLatesitallow";
            this.MgmDeptTxtLatesitallow.Size = new System.Drawing.Size(44, 20);
            this.MgmDeptTxtLatesitallow.TabIndex = 31;
            this.MgmDeptTxtLatesitallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtLatesitallow_KeyPress);
            // 
            // MgmDeptBtnDelete
            // 
            this.MgmDeptBtnDelete.Location = new System.Drawing.Point(833, 100);
            this.MgmDeptBtnDelete.Name = "MgmDeptBtnDelete";
            this.MgmDeptBtnDelete.Size = new System.Drawing.Size(74, 23);
            this.MgmDeptBtnDelete.TabIndex = 30;
            this.MgmDeptBtnDelete.Text = "Delete";
            this.MgmDeptBtnDelete.UseVisualStyleBackColor = true;
            this.MgmDeptBtnDelete.Click += new System.EventHandler(this.MgmDeptBtnDelete_Click);
            // 
            // MgmDeptBtnCancel
            // 
            this.MgmDeptBtnCancel.Location = new System.Drawing.Point(755, 129);
            this.MgmDeptBtnCancel.Name = "MgmDeptBtnCancel";
            this.MgmDeptBtnCancel.Size = new System.Drawing.Size(74, 23);
            this.MgmDeptBtnCancel.TabIndex = 29;
            this.MgmDeptBtnCancel.Text = "Cancel";
            this.MgmDeptBtnCancel.UseVisualStyleBackColor = true;
            this.MgmDeptBtnCancel.Click += new System.EventHandler(this.MgmDeptBtnCancel_Click);
            // 
            // MgmDeptLblOtherallow
            // 
            this.MgmDeptLblOtherallow.AutoSize = true;
            this.MgmDeptLblOtherallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblOtherallow.Location = new System.Drawing.Point(7, 104);
            this.MgmDeptLblOtherallow.Name = "MgmDeptLblOtherallow";
            this.MgmDeptLblOtherallow.Size = new System.Drawing.Size(100, 13);
            this.MgmDeptLblOtherallow.TabIndex = 28;
            this.MgmDeptLblOtherallow.Text = "Other Allowance";
            // 
            // MgmDeptTxtOtherallow
            // 
            this.MgmDeptTxtOtherallow.Location = new System.Drawing.Point(113, 100);
            this.MgmDeptTxtOtherallow.Name = "MgmDeptTxtOtherallow";
            this.MgmDeptTxtOtherallow.Size = new System.Drawing.Size(44, 20);
            this.MgmDeptTxtOtherallow.TabIndex = 27;
            this.MgmDeptTxtOtherallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtOtherallow_KeyPress);
            // 
            // MgmDeptLblPhoneallow
            // 
            this.MgmDeptLblPhoneallow.AutoSize = true;
            this.MgmDeptLblPhoneallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblPhoneallow.Location = new System.Drawing.Point(520, 74);
            this.MgmDeptLblPhoneallow.Name = "MgmDeptLblPhoneallow";
            this.MgmDeptLblPhoneallow.Size = new System.Drawing.Size(105, 13);
            this.MgmDeptLblPhoneallow.TabIndex = 26;
            this.MgmDeptLblPhoneallow.Text = "Phone Allowance";
            // 
            // MgmDeptTxtPhoneallow
            // 
            this.MgmDeptTxtPhoneallow.Location = new System.Drawing.Point(627, 70);
            this.MgmDeptTxtPhoneallow.Name = "MgmDeptTxtPhoneallow";
            this.MgmDeptTxtPhoneallow.Size = new System.Drawing.Size(46, 20);
            this.MgmDeptTxtPhoneallow.TabIndex = 25;
            this.MgmDeptTxtPhoneallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtPhoneallow_KeyPress);
            // 
            // MgmDeptLblConvallow
            // 
            this.MgmDeptLblConvallow.AutoSize = true;
            this.MgmDeptLblConvallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblConvallow.Location = new System.Drawing.Point(318, 74);
            this.MgmDeptLblConvallow.Name = "MgmDeptLblConvallow";
            this.MgmDeptLblConvallow.Size = new System.Drawing.Size(139, 13);
            this.MgmDeptLblConvallow.TabIndex = 24;
            this.MgmDeptLblConvallow.Text = "Conveyance Allowance";
            // 
            // MgmDeptTxtConvallow
            // 
            this.MgmDeptTxtConvallow.Location = new System.Drawing.Point(455, 70);
            this.MgmDeptTxtConvallow.Name = "MgmDeptTxtConvallow";
            this.MgmDeptTxtConvallow.Size = new System.Drawing.Size(46, 20);
            this.MgmDeptTxtConvallow.TabIndex = 23;
            this.MgmDeptTxtConvallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtConvallow_KeyPress);
            // 
            // MgmDeptLblMedicalallow
            // 
            this.MgmDeptLblMedicalallow.AutoSize = true;
            this.MgmDeptLblMedicalallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblMedicalallow.Location = new System.Drawing.Point(148, 74);
            this.MgmDeptLblMedicalallow.Name = "MgmDeptLblMedicalallow";
            this.MgmDeptLblMedicalallow.Size = new System.Drawing.Size(113, 13);
            this.MgmDeptLblMedicalallow.TabIndex = 22;
            this.MgmDeptLblMedicalallow.Text = "Medical Allowance";
            // 
            // MgmDeptTxtMedicalallow
            // 
            this.MgmDeptTxtMedicalallow.Location = new System.Drawing.Point(267, 70);
            this.MgmDeptTxtMedicalallow.Name = "MgmDeptTxtMedicalallow";
            this.MgmDeptTxtMedicalallow.Size = new System.Drawing.Size(45, 20);
            this.MgmDeptTxtMedicalallow.TabIndex = 21;
            this.MgmDeptTxtMedicalallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtMedicalallow_KeyPress);
            // 
            // MgmDeptLblHouserent
            // 
            this.MgmDeptLblHouserent.AutoSize = true;
            this.MgmDeptLblHouserent.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblHouserent.Location = new System.Drawing.Point(7, 74);
            this.MgmDeptLblHouserent.Name = "MgmDeptLblHouserent";
            this.MgmDeptLblHouserent.Size = new System.Drawing.Size(74, 13);
            this.MgmDeptLblHouserent.TabIndex = 20;
            this.MgmDeptLblHouserent.Text = "House Rent";
            // 
            // MgmDeptTxtHouserent
            // 
            this.MgmDeptTxtHouserent.Location = new System.Drawing.Point(84, 70);
            this.MgmDeptTxtHouserent.Name = "MgmDeptTxtHouserent";
            this.MgmDeptTxtHouserent.Size = new System.Drawing.Size(45, 20);
            this.MgmDeptTxtHouserent.TabIndex = 19;
            this.MgmDeptTxtHouserent.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtHouserent_KeyPress);
            // 
            // MgmDeptLblPerIncr
            // 
            this.MgmDeptLblPerIncr.AutoSize = true;
            this.MgmDeptLblPerIncr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblPerIncr.Location = new System.Drawing.Point(701, 47);
            this.MgmDeptLblPerIncr.Name = "MgmDeptLblPerIncr";
            this.MgmDeptLblPerIncr.Size = new System.Drawing.Size(146, 13);
            this.MgmDeptLblPerIncr.TabIndex = 18;
            this.MgmDeptLblPerIncr.Text = "Increment in Percentage";
            // 
            // MgmDeptLblDeptScaleadd
            // 
            this.MgmDeptLblDeptScaleadd.AutoSize = true;
            this.MgmDeptLblDeptScaleadd.Location = new System.Drawing.Point(7, 22);
            this.MgmDeptLblDeptScaleadd.Name = "MgmDeptLblDeptScaleadd";
            this.MgmDeptLblDeptScaleadd.Size = new System.Drawing.Size(152, 13);
            this.MgmDeptLblDeptScaleadd.TabIndex = 8;
            this.MgmDeptLblDeptScaleadd.Text = "Department Id must be unique.";
            // 
            // MgmDeptTxtPerIncr
            // 
            this.MgmDeptTxtPerIncr.Location = new System.Drawing.Point(849, 43);
            this.MgmDeptTxtPerIncr.Name = "MgmDeptTxtPerIncr";
            this.MgmDeptTxtPerIncr.Size = new System.Drawing.Size(44, 20);
            this.MgmDeptTxtPerIncr.TabIndex = 17;
            this.MgmDeptTxtPerIncr.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtPerIncr_KeyPress);
            // 
            // MgmDeptLblDid
            // 
            this.MgmDeptLblDid.AutoSize = true;
            this.MgmDeptLblDid.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblDid.Location = new System.Drawing.Point(6, 54);
            this.MgmDeptLblDid.Name = "MgmDeptLblDid";
            this.MgmDeptLblDid.Size = new System.Drawing.Size(87, 13);
            this.MgmDeptLblDid.TabIndex = 7;
            this.MgmDeptLblDid.Text = "Department Id";
            // 
            // MgmDeptLblMaxbpay
            // 
            this.MgmDeptLblMaxbpay.AutoSize = true;
            this.MgmDeptLblMaxbpay.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblMaxbpay.Location = new System.Drawing.Point(520, 47);
            this.MgmDeptLblMaxbpay.Name = "MgmDeptLblMaxbpay";
            this.MgmDeptLblMaxbpay.Size = new System.Drawing.Size(90, 13);
            this.MgmDeptLblMaxbpay.TabIndex = 16;
            this.MgmDeptLblMaxbpay.Text = "Max Basic Pay";
            // 
            // MgmDeptLblDname
            // 
            this.MgmDeptLblDname.AutoSize = true;
            this.MgmDeptLblDname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblDname.Location = new System.Drawing.Point(199, 54);
            this.MgmDeptLblDname.Name = "MgmDeptLblDname";
            this.MgmDeptLblDname.Size = new System.Drawing.Size(108, 13);
            this.MgmDeptLblDname.TabIndex = 6;
            this.MgmDeptLblDname.Text = "Department Name";
            // 
            // MgmDeptLblMgrid
            // 
            this.MgmDeptLblMgrid.AutoSize = true;
            this.MgmDeptLblMgrid.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblMgrid.Location = new System.Drawing.Point(520, 54);
            this.MgmDeptLblMgrid.Name = "MgmDeptLblMgrid";
            this.MgmDeptLblMgrid.Size = new System.Drawing.Size(71, 13);
            this.MgmDeptLblMgrid.TabIndex = 5;
            this.MgmDeptLblMgrid.Text = "Manager Id";
            // 
            // MgmDeptTxtMaxbpay
            // 
            this.MgmDeptTxtMaxbpay.Location = new System.Drawing.Point(613, 43);
            this.MgmDeptTxtMaxbpay.Name = "MgmDeptTxtMaxbpay";
            this.MgmDeptTxtMaxbpay.Size = new System.Drawing.Size(77, 20);
            this.MgmDeptTxtMaxbpay.TabIndex = 15;
            this.MgmDeptTxtMaxbpay.Leave += new System.EventHandler(this.MgmDeptTxtMaxbpay_Leave);
            this.MgmDeptTxtMaxbpay.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtMaxbpay_KeyPress);
            // 
            // MgmDeptBtnSave
            // 
            this.MgmDeptBtnSave.Location = new System.Drawing.Point(755, 100);
            this.MgmDeptBtnSave.Name = "MgmDeptBtnSave";
            this.MgmDeptBtnSave.Size = new System.Drawing.Size(74, 23);
            this.MgmDeptBtnSave.TabIndex = 3;
            this.MgmDeptBtnSave.Text = "Save";
            this.MgmDeptBtnSave.UseVisualStyleBackColor = true;
            this.MgmDeptBtnSave.Click += new System.EventHandler(this.MgmDeptBtnSave_Click);
            // 
            // MgmDeptTxtMgrid
            // 
            this.MgmDeptTxtMgrid.Location = new System.Drawing.Point(591, 50);
            this.MgmDeptTxtMgrid.Name = "MgmDeptTxtMgrid";
            this.MgmDeptTxtMgrid.Size = new System.Drawing.Size(152, 20);
            this.MgmDeptTxtMgrid.TabIndex = 2;
            this.MgmDeptTxtMgrid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtMgrid_KeyPress);
            // 
            // MgmDeptTxtDid
            // 
            this.MgmDeptTxtDid.Location = new System.Drawing.Point(96, 50);
            this.MgmDeptTxtDid.Name = "MgmDeptTxtDid";
            this.MgmDeptTxtDid.Size = new System.Drawing.Size(97, 20);
            this.MgmDeptTxtDid.TabIndex = 0;
            this.MgmDeptTxtDid.Leave += new System.EventHandler(this.MgmDeptTxtDid_Leave);
            this.MgmDeptTxtDid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtDid_KeyPress);
            // 
            // MgmDeptLblBps
            // 
            this.MgmDeptLblBps.AutoSize = true;
            this.MgmDeptLblBps.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MgmDeptLblBps.Location = new System.Drawing.Point(7, 47);
            this.MgmDeptLblBps.Name = "MgmDeptLblBps";
            this.MgmDeptLblBps.Size = new System.Drawing.Size(31, 13);
            this.MgmDeptLblBps.TabIndex = 10;
            this.MgmDeptLblBps.Text = "BPS";
            // 
            // MgmDeptTxtBps
            // 
            this.MgmDeptTxtBps.Location = new System.Drawing.Point(38, 43);
            this.MgmDeptTxtBps.Name = "MgmDeptTxtBps";
            this.MgmDeptTxtBps.Size = new System.Drawing.Size(44, 20);
            this.MgmDeptTxtBps.TabIndex = 9;
            this.MgmDeptTxtBps.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtBps_KeyPress);
            // 
            // MgmDeptTxtJobid
            // 
            this.MgmDeptTxtJobid.Location = new System.Drawing.Point(180, 43);
            this.MgmDeptTxtJobid.Name = "MgmDeptTxtJobid";
            this.MgmDeptTxtJobid.Size = new System.Drawing.Size(132, 20);
            this.MgmDeptTxtJobid.TabIndex = 11;
            this.MgmDeptTxtJobid.Leave += new System.EventHandler(this.MgmDeptTxtJobid_Leave);
            this.MgmDeptTxtJobid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmDeptTxtJobid_KeyPress);
            // 
            // MgmDeptBtnRefresh
            // 
            this.MgmDeptBtnRefresh.Location = new System.Drawing.Point(283, 122);
            this.MgmDeptBtnRefresh.Name = "MgmDeptBtnRefresh";
            this.MgmDeptBtnRefresh.Size = new System.Drawing.Size(74, 23);
            this.MgmDeptBtnRefresh.TabIndex = 10;
            this.MgmDeptBtnRefresh.Text = "Refresh";
            this.MgmDeptBtnRefresh.UseVisualStyleBackColor = true;
            this.MgmDeptBtnRefresh.Click += new System.EventHandler(this.BtnRefresh_Click);
            // 
            // TabJob
            // 
            this.TabJob.BackColor = System.Drawing.Color.AliceBlue;
            this.TabJob.Controls.Add(this.LblJob_update);
            this.TabJob.Controls.Add(this.MgmJobGroupSelectedemp);
            this.TabJob.Controls.Add(this.MgmJobGroupSearch);
            this.TabJob.Controls.Add(this.MgmJobBtnEdit);
            this.TabJob.Controls.Add(this.MgmJobGridvu);
            this.TabJob.Location = new System.Drawing.Point(4, 22);
            this.TabJob.Name = "TabJob";
            this.TabJob.Size = new System.Drawing.Size(937, 551);
            this.TabJob.TabIndex = 2;
            this.TabJob.Text = "Job Update";
            // 
            // LblJob_update
            // 
            this.LblJob_update.BackColor = System.Drawing.Color.WhiteSmoke;
            this.LblJob_update.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblJob_update.Location = new System.Drawing.Point(0, 0);
            this.LblJob_update.Name = "LblJob_update";
            this.LblJob_update.Size = new System.Drawing.Size(937, 24);
            this.LblJob_update.TabIndex = 80;
            this.LblJob_update.Text = "                                                             MIA JOB UPDATE";
            // 
            // MgmJobGroupSelectedemp
            // 
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobGroupeEmpdetail);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobTxtPermotedby);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobLblPermotedby);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobBtnSaverec);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobBtnCancel);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobTxtFormid1);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobLblFormid);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobPicbox);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobLblPic);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobCkbPermotion);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobGroupRestricted);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobGroupRestrictedpermote);
            this.MgmJobGroupSelectedemp.Controls.Add(this.MgmJobPanelPermotion);
            this.MgmJobGroupSelectedemp.Location = new System.Drawing.Point(12, 262);
            this.MgmJobGroupSelectedemp.Name = "MgmJobGroupSelectedemp";
            this.MgmJobGroupSelectedemp.Size = new System.Drawing.Size(914, 271);
            this.MgmJobGroupSelectedemp.TabIndex = 9;
            this.MgmJobGroupSelectedemp.TabStop = false;
            this.MgmJobGroupSelectedemp.Text = "Selected Employee Detail";
            this.MgmJobGroupSelectedemp.Visible = false;
            // 
            // MgmJobGroupeEmpdetail
            // 
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobCmb1Dept);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobLblJoindate);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobLblJobid);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobCtlHiredate);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobLbl1Dept);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobLblRftag);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobTxtName);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobCmbEmptype);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobLblName);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobLblEmptype);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobTxtRftag);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobCmb0Jobid);
            this.MgmJobGroupeEmpdetail.Controls.Add(this.MgmJobTxtJobid);
            this.MgmJobGroupeEmpdetail.Location = new System.Drawing.Point(10, 35);
            this.MgmJobGroupeEmpdetail.Name = "MgmJobGroupeEmpdetail";
            this.MgmJobGroupeEmpdetail.Size = new System.Drawing.Size(726, 96);
            this.MgmJobGroupeEmpdetail.TabIndex = 27;
            this.MgmJobGroupeEmpdetail.TabStop = false;
            this.MgmJobGroupeEmpdetail.Text = "Employee Detail";
            // 
            // MgmJobCmb1Dept
            // 
            this.MgmJobCmb1Dept.FormattingEnabled = true;
            this.MgmJobCmb1Dept.Location = new System.Drawing.Point(487, 57);
            this.MgmJobCmb1Dept.Name = "MgmJobCmb1Dept";
            this.MgmJobCmb1Dept.Size = new System.Drawing.Size(133, 21);
            this.MgmJobCmb1Dept.TabIndex = 9;
            this.MgmJobCmb1Dept.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MgmJobCmb1Dept_MouseDown);
            // 
            // MgmJobLblJoindate
            // 
            this.MgmJobLblJoindate.AutoSize = true;
            this.MgmJobLblJoindate.Location = new System.Drawing.Point(415, 33);
            this.MgmJobLblJoindate.Name = "MgmJobLblJoindate";
            this.MgmJobLblJoindate.Size = new System.Drawing.Size(66, 13);
            this.MgmJobLblJoindate.TabIndex = 37;
            this.MgmJobLblJoindate.Text = "Joining Date";
            // 
            // MgmJobLblJobid
            // 
            this.MgmJobLblJobid.AutoSize = true;
            this.MgmJobLblJobid.Location = new System.Drawing.Point(13, 61);
            this.MgmJobLblJobid.Name = "MgmJobLblJobid";
            this.MgmJobLblJobid.Size = new System.Drawing.Size(32, 13);
            this.MgmJobLblJobid.TabIndex = 6;
            this.MgmJobLblJobid.Text = "Jobid";
            // 
            // MgmJobCtlHiredate
            // 
            this.MgmJobCtlHiredate.Location = new System.Drawing.Point(487, 29);
            this.MgmJobCtlHiredate.Name = "MgmJobCtlHiredate";
            this.MgmJobCtlHiredate.Size = new System.Drawing.Size(119, 20);
            this.MgmJobCtlHiredate.TabIndex = 36;
            // 
            // MgmJobLbl1Dept
            // 
            this.MgmJobLbl1Dept.AutoSize = true;
            this.MgmJobLbl1Dept.Location = new System.Drawing.Point(416, 61);
            this.MgmJobLbl1Dept.Name = "MgmJobLbl1Dept";
            this.MgmJobLbl1Dept.Size = new System.Drawing.Size(62, 13);
            this.MgmJobLbl1Dept.TabIndex = 10;
            this.MgmJobLbl1Dept.Text = "Department";
            // 
            // MgmJobLblRftag
            // 
            this.MgmJobLblRftag.AutoSize = true;
            this.MgmJobLblRftag.Location = new System.Drawing.Point(207, 33);
            this.MgmJobLblRftag.Name = "MgmJobLblRftag";
            this.MgmJobLblRftag.Size = new System.Drawing.Size(43, 13);
            this.MgmJobLblRftag.TabIndex = 35;
            this.MgmJobLblRftag.Text = "RF Tag";
            // 
            // MgmJobTxtName
            // 
            this.MgmJobTxtName.Location = new System.Drawing.Point(55, 29);
            this.MgmJobTxtName.Name = "MgmJobTxtName";
            this.MgmJobTxtName.ReadOnly = true;
            this.MgmJobTxtName.Size = new System.Drawing.Size(147, 20);
            this.MgmJobTxtName.TabIndex = 1;
            // 
            // MgmJobCmbEmptype
            // 
            this.MgmJobCmbEmptype.FormattingEnabled = true;
            this.MgmJobCmbEmptype.Items.AddRange(new object[] {
            "CONTRACT",
            "PERMANENT"});
            this.MgmJobCmbEmptype.Location = new System.Drawing.Point(262, 57);
            this.MgmJobCmbEmptype.Name = "MgmJobCmbEmptype";
            this.MgmJobCmbEmptype.Size = new System.Drawing.Size(147, 21);
            this.MgmJobCmbEmptype.TabIndex = 34;
            // 
            // MgmJobLblName
            // 
            this.MgmJobLblName.AutoSize = true;
            this.MgmJobLblName.Location = new System.Drawing.Point(13, 33);
            this.MgmJobLblName.Name = "MgmJobLblName";
            this.MgmJobLblName.Size = new System.Drawing.Size(35, 13);
            this.MgmJobLblName.TabIndex = 2;
            this.MgmJobLblName.Text = "Name";
            // 
            // MgmJobLblEmptype
            // 
            this.MgmJobLblEmptype.AutoSize = true;
            this.MgmJobLblEmptype.Location = new System.Drawing.Point(207, 61);
            this.MgmJobLblEmptype.Name = "MgmJobLblEmptype";
            this.MgmJobLblEmptype.Size = new System.Drawing.Size(55, 13);
            this.MgmJobLblEmptype.TabIndex = 33;
            this.MgmJobLblEmptype.Text = "Emp Type";
            // 
            // MgmJobTxtRftag
            // 
            this.MgmJobTxtRftag.Location = new System.Drawing.Point(260, 29);
            this.MgmJobTxtRftag.Name = "MgmJobTxtRftag";
            this.MgmJobTxtRftag.Size = new System.Drawing.Size(149, 20);
            this.MgmJobTxtRftag.TabIndex = 32;
            // 
            // MgmJobCmb0Jobid
            // 
            this.MgmJobCmb0Jobid.FormattingEnabled = true;
            this.MgmJobCmb0Jobid.Location = new System.Drawing.Point(55, 57);
            this.MgmJobCmb0Jobid.Name = "MgmJobCmb0Jobid";
            this.MgmJobCmb0Jobid.Size = new System.Drawing.Size(147, 21);
            this.MgmJobCmb0Jobid.TabIndex = 30;
            this.MgmJobCmb0Jobid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MgmJobCmb0Jobid_MouseDown);
            // 
            // MgmJobTxtJobid
            // 
            this.MgmJobTxtJobid.Location = new System.Drawing.Point(55, 57);
            this.MgmJobTxtJobid.Name = "MgmJobTxtJobid";
            this.MgmJobTxtJobid.ReadOnly = true;
            this.MgmJobTxtJobid.Size = new System.Drawing.Size(147, 20);
            this.MgmJobTxtJobid.TabIndex = 21;
            // 
            // MgmJobTxtPermotedby
            // 
            this.MgmJobTxtPermotedby.AcceptsReturn = true;
            this.MgmJobTxtPermotedby.Location = new System.Drawing.Point(808, 191);
            this.MgmJobTxtPermotedby.Name = "MgmJobTxtPermotedby";
            this.MgmJobTxtPermotedby.ReadOnly = true;
            this.MgmJobTxtPermotedby.Size = new System.Drawing.Size(94, 20);
            this.MgmJobTxtPermotedby.TabIndex = 28;
            // 
            // MgmJobLblPermotedby
            // 
            this.MgmJobLblPermotedby.AutoSize = true;
            this.MgmJobLblPermotedby.Location = new System.Drawing.Point(746, 194);
            this.MgmJobLblPermotedby.Name = "MgmJobLblPermotedby";
            this.MgmJobLblPermotedby.Size = new System.Drawing.Size(62, 13);
            this.MgmJobLblPermotedby.TabIndex = 27;
            this.MgmJobLblPermotedby.Text = "Updated by";
            // 
            // MgmJobBtnSaverec
            // 
            this.MgmJobBtnSaverec.Enabled = false;
            this.MgmJobBtnSaverec.Location = new System.Drawing.Point(749, 225);
            this.MgmJobBtnSaverec.Name = "MgmJobBtnSaverec";
            this.MgmJobBtnSaverec.Size = new System.Drawing.Size(68, 23);
            this.MgmJobBtnSaverec.TabIndex = 8;
            this.MgmJobBtnSaverec.Text = "Save";
            this.MgmJobBtnSaverec.UseVisualStyleBackColor = true;
            this.MgmJobBtnSaverec.Click += new System.EventHandler(this.MgmJobBtnSave_Click);
            // 
            // MgmJobBtnCancel
            // 
            this.MgmJobBtnCancel.Enabled = false;
            this.MgmJobBtnCancel.Location = new System.Drawing.Point(834, 225);
            this.MgmJobBtnCancel.Name = "MgmJobBtnCancel";
            this.MgmJobBtnCancel.Size = new System.Drawing.Size(68, 23);
            this.MgmJobBtnCancel.TabIndex = 20;
            this.MgmJobBtnCancel.Text = "Cancel";
            this.MgmJobBtnCancel.UseVisualStyleBackColor = true;
            this.MgmJobBtnCancel.Click += new System.EventHandler(this.MgmJobBtnCancel_Click);
            // 
            // MgmJobTxtFormid1
            // 
            this.MgmJobTxtFormid1.Location = new System.Drawing.Point(831, 162);
            this.MgmJobTxtFormid1.Name = "MgmJobTxtFormid1";
            this.MgmJobTxtFormid1.ReadOnly = true;
            this.MgmJobTxtFormid1.Size = new System.Drawing.Size(72, 20);
            this.MgmJobTxtFormid1.TabIndex = 13;
            // 
            // MgmJobLblFormid
            // 
            this.MgmJobLblFormid.AutoSize = true;
            this.MgmJobLblFormid.Location = new System.Drawing.Point(789, 166);
            this.MgmJobLblFormid.Name = "MgmJobLblFormid";
            this.MgmJobLblFormid.Size = new System.Drawing.Size(42, 13);
            this.MgmJobLblFormid.TabIndex = 14;
            this.MgmJobLblFormid.Text = "Form Id";
            // 
            // MgmJobPicbox
            // 
            this.MgmJobPicbox.BackColor = System.Drawing.Color.Gray;
            this.MgmJobPicbox.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.MgmJobPicbox.Location = new System.Drawing.Point(792, 20);
            this.MgmJobPicbox.Name = "MgmJobPicbox";
            this.MgmJobPicbox.Size = new System.Drawing.Size(110, 138);
            this.MgmJobPicbox.TabIndex = 0;
            this.MgmJobPicbox.TabStop = false;
            // 
            // MgmJobLblPic
            // 
            this.MgmJobLblPic.AutoSize = true;
            this.MgmJobLblPic.Location = new System.Drawing.Point(750, 19);
            this.MgmJobLblPic.Name = "MgmJobLblPic";
            this.MgmJobLblPic.Size = new System.Drawing.Size(40, 13);
            this.MgmJobLblPic.TabIndex = 7;
            this.MgmJobLblPic.Text = "Picture";
            // 
            // MgmJobCkbPermotion
            // 
            this.MgmJobCkbPermotion.AutoSize = true;
            this.MgmJobCkbPermotion.Location = new System.Drawing.Point(10, 141);
            this.MgmJobCkbPermotion.Name = "MgmJobCkbPermotion";
            this.MgmJobCkbPermotion.Size = new System.Drawing.Size(73, 17);
            this.MgmJobCkbPermotion.TabIndex = 18;
            this.MgmJobCkbPermotion.Text = "Permotion";
            this.MgmJobCkbPermotion.UseVisualStyleBackColor = true;
            this.MgmJobCkbPermotion.CheckedChanged += new System.EventHandler(this.MgmJobCkbPermotion_CheckedChanged);
            // 
            // MgmJobGroupRestricted
            // 
            this.MgmJobGroupRestricted.BackColor = System.Drawing.Color.AliceBlue;
            this.MgmJobGroupRestricted.Controls.Add(this.label3);
            this.MgmJobGroupRestricted.Controls.Add(this.label21);
            this.MgmJobGroupRestricted.Controls.Add(this.pictureBox5);
            this.MgmJobGroupRestricted.Location = new System.Drawing.Point(11, 24);
            this.MgmJobGroupRestricted.Name = "MgmJobGroupRestricted";
            this.MgmJobGroupRestricted.Size = new System.Drawing.Size(731, 110);
            this.MgmJobGroupRestricted.TabIndex = 63;
            this.MgmJobGroupRestricted.TabStop = false;
            this.MgmJobGroupRestricted.Text = "Restricted Detail";
            this.MgmJobGroupRestricted.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(125, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 60;
            this.label3.Text = "SORRY!";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.ForeColor = System.Drawing.Color.Red;
            this.label21.Location = new System.Drawing.Point(125, 54);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(301, 13);
            this.label21.TabIndex = 32;
            this.label21.Text = "Before 25th of this Month, You can not Permote to Employees.";
            // 
            // pictureBox5
            // 
            this.pictureBox5.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox5.BackgroundImage")));
            this.pictureBox5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox5.Location = new System.Drawing.Point(10, 19);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(102, 81);
            this.pictureBox5.TabIndex = 58;
            this.pictureBox5.TabStop = false;
            // 
            // MgmJobGroupRestrictedpermote
            // 
            this.MgmJobGroupRestrictedpermote.BackColor = System.Drawing.Color.AliceBlue;
            this.MgmJobGroupRestrictedpermote.Controls.Add(this.label25);
            this.MgmJobGroupRestrictedpermote.Controls.Add(this.label29);
            this.MgmJobGroupRestrictedpermote.Controls.Add(this.pictureBox6);
            this.MgmJobGroupRestrictedpermote.Location = new System.Drawing.Point(10, 164);
            this.MgmJobGroupRestrictedpermote.Name = "MgmJobGroupRestrictedpermote";
            this.MgmJobGroupRestrictedpermote.Size = new System.Drawing.Size(568, 101);
            this.MgmJobGroupRestrictedpermote.TabIndex = 64;
            this.MgmJobGroupRestrictedpermote.TabStop = false;
            this.MgmJobGroupRestrictedpermote.Text = "Restricted to Permote";
            this.MgmJobGroupRestrictedpermote.Visible = false;
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label25.ForeColor = System.Drawing.Color.Red;
            this.label25.Location = new System.Drawing.Point(125, 27);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(54, 13);
            this.label25.TabIndex = 60;
            this.label25.Text = "SORRY!";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.ForeColor = System.Drawing.Color.Red;
            this.label29.Location = new System.Drawing.Point(125, 54);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(311, 13);
            this.label29.TabIndex = 32;
            this.label29.Text = "Before 25th of this Month, You can not Permote to an Employee.";
            // 
            // pictureBox6
            // 
            this.pictureBox6.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox6.BackgroundImage")));
            this.pictureBox6.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox6.Location = new System.Drawing.Point(10, 16);
            this.pictureBox6.Name = "pictureBox6";
            this.pictureBox6.Size = new System.Drawing.Size(102, 81);
            this.pictureBox6.TabIndex = 58;
            this.pictureBox6.TabStop = false;
            // 
            // MgmJobPanelPermotion
            // 
            this.MgmJobPanelPermotion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobCmb3Dept);
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobLblComments);
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobLbl1Jobid);
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobLbl2Dept);
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobTxtComments);
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobCmb1Bps);
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobCmb1Jobid);
            this.MgmJobPanelPermotion.Controls.Add(this.MgmJobLblBps1);
            this.MgmJobPanelPermotion.Location = new System.Drawing.Point(10, 166);
            this.MgmJobPanelPermotion.Name = "MgmJobPanelPermotion";
            this.MgmJobPanelPermotion.Size = new System.Drawing.Size(568, 99);
            this.MgmJobPanelPermotion.TabIndex = 19;
            this.MgmJobPanelPermotion.Visible = false;
            // 
            // MgmJobCmb3Dept
            // 
            this.MgmJobCmb3Dept.FormattingEnabled = true;
            this.MgmJobCmb3Dept.Location = new System.Drawing.Point(400, 26);
            this.MgmJobCmb3Dept.Name = "MgmJobCmb3Dept";
            this.MgmJobCmb3Dept.Size = new System.Drawing.Size(147, 21);
            this.MgmJobCmb3Dept.TabIndex = 21;
            this.MgmJobCmb3Dept.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MgmJobCmb3Dept_MouseDown);
            // 
            // MgmJobLblComments
            // 
            this.MgmJobLblComments.AutoSize = true;
            this.MgmJobLblComments.Location = new System.Drawing.Point(21, 61);
            this.MgmJobLblComments.Name = "MgmJobLblComments";
            this.MgmJobLblComments.Size = new System.Drawing.Size(56, 13);
            this.MgmJobLblComments.TabIndex = 25;
            this.MgmJobLblComments.Text = "Comments";
            // 
            // MgmJobLbl1Jobid
            // 
            this.MgmJobLbl1Jobid.AutoSize = true;
            this.MgmJobLbl1Jobid.Location = new System.Drawing.Point(128, 30);
            this.MgmJobLbl1Jobid.Name = "MgmJobLbl1Jobid";
            this.MgmJobLbl1Jobid.Size = new System.Drawing.Size(32, 13);
            this.MgmJobLbl1Jobid.TabIndex = 20;
            this.MgmJobLbl1Jobid.Text = "Jobid";
            // 
            // MgmJobLbl2Dept
            // 
            this.MgmJobLbl2Dept.AutoSize = true;
            this.MgmJobLbl2Dept.Location = new System.Drawing.Point(338, 30);
            this.MgmJobLbl2Dept.Name = "MgmJobLbl2Dept";
            this.MgmJobLbl2Dept.Size = new System.Drawing.Size(62, 13);
            this.MgmJobLbl2Dept.TabIndex = 22;
            this.MgmJobLbl2Dept.Text = "Department";
            // 
            // MgmJobTxtComments
            // 
            this.MgmJobTxtComments.AcceptsReturn = true;
            this.MgmJobTxtComments.Location = new System.Drawing.Point(98, 57);
            this.MgmJobTxtComments.Name = "MgmJobTxtComments";
            this.MgmJobTxtComments.Size = new System.Drawing.Size(449, 20);
            this.MgmJobTxtComments.TabIndex = 26;
            // 
            // MgmJobCmb1Bps
            // 
            this.MgmJobCmb1Bps.FormattingEnabled = true;
            this.MgmJobCmb1Bps.Location = new System.Drawing.Point(51, 26);
            this.MgmJobCmb1Bps.Name = "MgmJobCmb1Bps";
            this.MgmJobCmb1Bps.Size = new System.Drawing.Size(74, 21);
            this.MgmJobCmb1Bps.TabIndex = 23;
            this.MgmJobCmb1Bps.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MgmJobCmb1Bps_MouseDown);
            // 
            // MgmJobCmb1Jobid
            // 
            this.MgmJobCmb1Jobid.FormattingEnabled = true;
            this.MgmJobCmb1Jobid.Location = new System.Drawing.Point(161, 26);
            this.MgmJobCmb1Jobid.Name = "MgmJobCmb1Jobid";
            this.MgmJobCmb1Jobid.Size = new System.Drawing.Size(170, 21);
            this.MgmJobCmb1Jobid.TabIndex = 24;
            this.MgmJobCmb1Jobid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MgmJobCmb1Jobid_MouseDown);
            // 
            // MgmJobLblBps1
            // 
            this.MgmJobLblBps1.AutoSize = true;
            this.MgmJobLblBps1.Location = new System.Drawing.Point(21, 30);
            this.MgmJobLblBps1.Name = "MgmJobLblBps1";
            this.MgmJobLblBps1.Size = new System.Drawing.Size(25, 13);
            this.MgmJobLblBps1.TabIndex = 20;
            this.MgmJobLblBps1.Text = "Bps";
            // 
            // MgmJobGroupSearch
            // 
            this.MgmJobGroupSearch.Controls.Add(this.MgmJobLblCaution);
            this.MgmJobGroupSearch.Controls.Add(this.MgmJobTxtformid0);
            this.MgmJobGroupSearch.Controls.Add(this.MgmJobCmb0Dept);
            this.MgmJobGroupSearch.Controls.Add(this.MgmJobCkbFormid);
            this.MgmJobGroupSearch.Controls.Add(this.MgmJobBtnSearch);
            this.MgmJobGroupSearch.Controls.Add(this.MgmJobCkbdDeptwise);
            this.MgmJobGroupSearch.Controls.Add(this.MgmJobCkbJobisnull);
            this.MgmJobGroupSearch.Location = new System.Drawing.Point(13, 41);
            this.MgmJobGroupSearch.Name = "MgmJobGroupSearch";
            this.MgmJobGroupSearch.Size = new System.Drawing.Size(319, 182);
            this.MgmJobGroupSearch.TabIndex = 8;
            this.MgmJobGroupSearch.TabStop = false;
            this.MgmJobGroupSearch.Text = "Search Options";
            // 
            // MgmJobLblCaution
            // 
            this.MgmJobLblCaution.AutoSize = true;
            this.MgmJobLblCaution.Location = new System.Drawing.Point(10, 32);
            this.MgmJobLblCaution.Name = "MgmJobLblCaution";
            this.MgmJobLblCaution.Size = new System.Drawing.Size(242, 13);
            this.MgmJobLblCaution.TabIndex = 26;
            this.MgmJobLblCaution.Text = "Following Search option can be used for a record.";
            // 
            // MgmJobTxtformid0
            // 
            this.MgmJobTxtformid0.Location = new System.Drawing.Point(171, 116);
            this.MgmJobTxtformid0.Name = "MgmJobTxtformid0";
            this.MgmJobTxtformid0.Size = new System.Drawing.Size(129, 20);
            this.MgmJobTxtformid0.TabIndex = 7;
            this.MgmJobTxtformid0.Visible = false;
            // 
            // MgmJobCmb0Dept
            // 
            this.MgmJobCmb0Dept.FormattingEnabled = true;
            this.MgmJobCmb0Dept.Location = new System.Drawing.Point(171, 89);
            this.MgmJobCmb0Dept.Name = "MgmJobCmb0Dept";
            this.MgmJobCmb0Dept.Size = new System.Drawing.Size(129, 21);
            this.MgmJobCmb0Dept.TabIndex = 6;
            this.MgmJobCmb0Dept.Visible = false;
            this.MgmJobCmb0Dept.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MgmJobCmb0Dept_MouseDown);
            // 
            // MgmJobCkbFormid
            // 
            this.MgmJobCkbFormid.AutoSize = true;
            this.MgmJobCkbFormid.Enabled = false;
            this.MgmJobCkbFormid.Location = new System.Drawing.Point(10, 112);
            this.MgmJobCkbFormid.Name = "MgmJobCkbFormid";
            this.MgmJobCkbFormid.Size = new System.Drawing.Size(111, 17);
            this.MgmJobCkbFormid.TabIndex = 5;
            this.MgmJobCkbFormid.Text = "Search by Form id";
            this.MgmJobCkbFormid.UseVisualStyleBackColor = true;
            this.MgmJobCkbFormid.CheckedChanged += new System.EventHandler(this.MgmJobLblFormid_CheckedChanged);
            // 
            // MgmJobBtnSearch
            // 
            this.MgmJobBtnSearch.Location = new System.Drawing.Point(225, 142);
            this.MgmJobBtnSearch.Name = "MgmJobBtnSearch";
            this.MgmJobBtnSearch.Size = new System.Drawing.Size(75, 23);
            this.MgmJobBtnSearch.TabIndex = 6;
            this.MgmJobBtnSearch.Text = "Search";
            this.MgmJobBtnSearch.UseVisualStyleBackColor = true;
            this.MgmJobBtnSearch.Click += new System.EventHandler(this.MgmJobBtnSearch_Click);
            // 
            // MgmJobCkbdDeptwise
            // 
            this.MgmJobCkbdDeptwise.AutoSize = true;
            this.MgmJobCkbdDeptwise.Enabled = false;
            this.MgmJobCkbdDeptwise.Location = new System.Drawing.Point(10, 89);
            this.MgmJobCkbdDeptwise.Name = "MgmJobCkbdDeptwise";
            this.MgmJobCkbdDeptwise.Size = new System.Drawing.Size(157, 17);
            this.MgmJobCkbdDeptwise.TabIndex = 4;
            this.MgmJobCkbdDeptwise.Text = "Search via department wise";
            this.MgmJobCkbdDeptwise.UseVisualStyleBackColor = true;
            this.MgmJobCkbdDeptwise.CheckedChanged += new System.EventHandler(this.MgmJobCkbdDeptwise_CheckedChanged);
            // 
            // MgmJobCkbJobisnull
            // 
            this.MgmJobCkbJobisnull.AutoSize = true;
            this.MgmJobCkbJobisnull.Location = new System.Drawing.Point(10, 66);
            this.MgmJobCkbJobisnull.Name = "MgmJobCkbJobisnull";
            this.MgmJobCkbJobisnull.Size = new System.Drawing.Size(143, 17);
            this.MgmJobCkbJobisnull.TabIndex = 3;
            this.MgmJobCkbJobisnull.Text = "Search where Job is Null";
            this.MgmJobCkbJobisnull.UseVisualStyleBackColor = true;
            this.MgmJobCkbJobisnull.CheckStateChanged += new System.EventHandler(this.MgmJobLblJobisnull_CheckStateChanged);
            // 
            // MgmJobBtnEdit
            // 
            this.MgmJobBtnEdit.Location = new System.Drawing.Point(257, 229);
            this.MgmJobBtnEdit.Name = "MgmJobBtnEdit";
            this.MgmJobBtnEdit.Size = new System.Drawing.Size(75, 23);
            this.MgmJobBtnEdit.TabIndex = 5;
            this.MgmJobBtnEdit.Text = "Edit";
            this.MgmJobBtnEdit.UseVisualStyleBackColor = true;
            this.MgmJobBtnEdit.Click += new System.EventHandler(this.MgmJobBtnEdit_Click);
            // 
            // MgmJobGridvu
            // 
            this.MgmJobGridvu.AllowUserToAddRows = false;
            this.MgmJobGridvu.AllowUserToDeleteRows = false;
            this.MgmJobGridvu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MgmJobGridvu.Location = new System.Drawing.Point(363, 50);
            this.MgmJobGridvu.Name = "MgmJobGridvu";
            this.MgmJobGridvu.ReadOnly = true;
            this.MgmJobGridvu.Size = new System.Drawing.Size(555, 202);
            this.MgmJobGridvu.TabIndex = 2;
            // 
            // TabAttend
            // 
            this.TabAttend.BackColor = System.Drawing.Color.AliceBlue;
            this.TabAttend.Controls.Add(this.label1);
            this.TabAttend.Controls.Add(this.MgmAttGroupSearch);
            this.TabAttend.Controls.Add(this.MgmAttGroupDetailview);
            this.TabAttend.Controls.Add(this.MgmAttGroupEmpatt);
            this.TabAttend.Location = new System.Drawing.Point(4, 22);
            this.TabAttend.Name = "TabAttend";
            this.TabAttend.Size = new System.Drawing.Size(937, 551);
            this.TabAttend.TabIndex = 3;
            this.TabAttend.Text = "Attendance";
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(-3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(940, 24);
            this.label1.TabIndex = 81;
            this.label1.Text = "                                                 MIA DAILY ATTENDANCE VIEW";
            // 
            // MgmAttGroupSearch
            // 
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttCkbFormid);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttLblDatetime);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttDTP);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttBtnSearch);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttCmbDept);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttBtnClear);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttTxtFormid);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttCkbRftag);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttCkbDept);
            this.MgmAttGroupSearch.Controls.Add(this.MgmAttTxtRftag);
            this.MgmAttGroupSearch.Location = new System.Drawing.Point(619, 38);
            this.MgmAttGroupSearch.Name = "MgmAttGroupSearch";
            this.MgmAttGroupSearch.Size = new System.Drawing.Size(279, 215);
            this.MgmAttGroupSearch.TabIndex = 17;
            this.MgmAttGroupSearch.TabStop = false;
            this.MgmAttGroupSearch.Text = "Search Options";
            // 
            // MgmAttCkbFormid
            // 
            this.MgmAttCkbFormid.AutoSize = true;
            this.MgmAttCkbFormid.Enabled = false;
            this.MgmAttCkbFormid.Location = new System.Drawing.Point(21, 97);
            this.MgmAttCkbFormid.Name = "MgmAttCkbFormid";
            this.MgmAttCkbFormid.Size = new System.Drawing.Size(61, 17);
            this.MgmAttCkbFormid.TabIndex = 23;
            this.MgmAttCkbFormid.Text = "Form Id";
            this.MgmAttCkbFormid.UseVisualStyleBackColor = true;
            this.MgmAttCkbFormid.CheckedChanged += new System.EventHandler(this.MgmAttCkbFormid_CheckedChanged);
            // 
            // MgmAttLblDatetime
            // 
            this.MgmAttLblDatetime.AutoSize = true;
            this.MgmAttLblDatetime.Location = new System.Drawing.Point(21, 129);
            this.MgmAttLblDatetime.Name = "MgmAttLblDatetime";
            this.MgmAttLblDatetime.Size = new System.Drawing.Size(30, 13);
            this.MgmAttLblDatetime.TabIndex = 11;
            this.MgmAttLblDatetime.Text = "Date";
            // 
            // MgmAttDTP
            // 
            this.MgmAttDTP.Location = new System.Drawing.Point(102, 125);
            this.MgmAttDTP.Name = "MgmAttDTP";
            this.MgmAttDTP.Size = new System.Drawing.Size(151, 20);
            this.MgmAttDTP.TabIndex = 10;
            // 
            // MgmAttBtnSearch
            // 
            this.MgmAttBtnSearch.Location = new System.Drawing.Point(102, 167);
            this.MgmAttBtnSearch.Name = "MgmAttBtnSearch";
            this.MgmAttBtnSearch.Size = new System.Drawing.Size(74, 23);
            this.MgmAttBtnSearch.TabIndex = 18;
            this.MgmAttBtnSearch.Text = "Search";
            this.MgmAttBtnSearch.UseVisualStyleBackColor = true;
            this.MgmAttBtnSearch.Click += new System.EventHandler(this.MgmAttBtnSearch_Click);
            // 
            // MgmAttCmbDept
            // 
            this.MgmAttCmbDept.FormattingEnabled = true;
            this.MgmAttCmbDept.Location = new System.Drawing.Point(102, 37);
            this.MgmAttCmbDept.Name = "MgmAttCmbDept";
            this.MgmAttCmbDept.Size = new System.Drawing.Size(151, 21);
            this.MgmAttCmbDept.TabIndex = 15;
            this.MgmAttCmbDept.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MgmAttCmbDept_MouseDown);
            // 
            // MgmAttBtnClear
            // 
            this.MgmAttBtnClear.Location = new System.Drawing.Point(182, 167);
            this.MgmAttBtnClear.Name = "MgmAttBtnClear";
            this.MgmAttBtnClear.Size = new System.Drawing.Size(74, 23);
            this.MgmAttBtnClear.TabIndex = 13;
            this.MgmAttBtnClear.Text = "Clear";
            this.MgmAttBtnClear.UseVisualStyleBackColor = true;
            this.MgmAttBtnClear.Click += new System.EventHandler(this.MgmAttBtnClear_Click);
            // 
            // MgmAttTxtFormid
            // 
            this.MgmAttTxtFormid.AcceptsReturn = true;
            this.MgmAttTxtFormid.Enabled = false;
            this.MgmAttTxtFormid.Location = new System.Drawing.Point(102, 95);
            this.MgmAttTxtFormid.Name = "MgmAttTxtFormid";
            this.MgmAttTxtFormid.Size = new System.Drawing.Size(151, 20);
            this.MgmAttTxtFormid.TabIndex = 19;
            this.MgmAttTxtFormid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmAttTxtFormid_KeyPress);
            // 
            // MgmAttCkbRftag
            // 
            this.MgmAttCkbRftag.AutoSize = true;
            this.MgmAttCkbRftag.ForeColor = System.Drawing.Color.Black;
            this.MgmAttCkbRftag.Location = new System.Drawing.Point(21, 68);
            this.MgmAttCkbRftag.Name = "MgmAttCkbRftag";
            this.MgmAttCkbRftag.Size = new System.Drawing.Size(62, 17);
            this.MgmAttCkbRftag.TabIndex = 22;
            this.MgmAttCkbRftag.Text = "RF Tag";
            this.MgmAttCkbRftag.UseVisualStyleBackColor = true;
            this.MgmAttCkbRftag.Visible = false;
            this.MgmAttCkbRftag.CheckedChanged += new System.EventHandler(this.MgmAttLblRftag_CheckedChanged);
            // 
            // MgmAttCkbDept
            // 
            this.MgmAttCkbDept.AutoSize = true;
            this.MgmAttCkbDept.Checked = true;
            this.MgmAttCkbDept.CheckState = System.Windows.Forms.CheckState.Checked;
            this.MgmAttCkbDept.Location = new System.Drawing.Point(21, 39);
            this.MgmAttCkbDept.Name = "MgmAttCkbDept";
            this.MgmAttCkbDept.Size = new System.Drawing.Size(81, 17);
            this.MgmAttCkbDept.TabIndex = 21;
            this.MgmAttCkbDept.Text = "Department";
            this.MgmAttCkbDept.UseVisualStyleBackColor = true;
            this.MgmAttCkbDept.CheckStateChanged += new System.EventHandler(this.MgmAttCkbDept_CheckStateChanged);
            // 
            // MgmAttTxtRftag
            // 
            this.MgmAttTxtRftag.Location = new System.Drawing.Point(102, 66);
            this.MgmAttTxtRftag.Name = "MgmAttTxtRftag";
            this.MgmAttTxtRftag.Size = new System.Drawing.Size(151, 20);
            this.MgmAttTxtRftag.TabIndex = 17;
            this.MgmAttTxtRftag.Visible = false;
            this.MgmAttTxtRftag.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmAttTxtRftag_KeyPress);
            // 
            // MgmAttGroupDetailview
            // 
            this.MgmAttGroupDetailview.Controls.Add(this.MgmAttGridAttdetailvu);
            this.MgmAttGroupDetailview.Location = new System.Drawing.Point(18, 276);
            this.MgmAttGroupDetailview.Name = "MgmAttGroupDetailview";
            this.MgmAttGroupDetailview.Size = new System.Drawing.Size(905, 272);
            this.MgmAttGroupDetailview.TabIndex = 14;
            this.MgmAttGroupDetailview.TabStop = false;
            this.MgmAttGroupDetailview.Text = "Detail View";
            // 
            // MgmAttGridAttdetailvu
            // 
            this.MgmAttGridAttdetailvu.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.MgmAttGridAttdetailvu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MgmAttGridAttdetailvu.Location = new System.Drawing.Point(15, 17);
            this.MgmAttGridAttdetailvu.Name = "MgmAttGridAttdetailvu";
            this.MgmAttGridAttdetailvu.ReadOnly = true;
            this.MgmAttGridAttdetailvu.Size = new System.Drawing.Size(875, 245);
            this.MgmAttGridAttdetailvu.TabIndex = 4;
            // 
            // MgmAttGroupEmpatt
            // 
            this.MgmAttGroupEmpatt.Controls.Add(this.MgmAttGridAttendVu);
            this.MgmAttGroupEmpatt.Location = new System.Drawing.Point(18, 38);
            this.MgmAttGroupEmpatt.Name = "MgmAttGroupEmpatt";
            this.MgmAttGroupEmpatt.Size = new System.Drawing.Size(581, 215);
            this.MgmAttGroupEmpatt.TabIndex = 9;
            this.MgmAttGroupEmpatt.TabStop = false;
            this.MgmAttGroupEmpatt.Text = "Employee Attendance";
            // 
            // MgmAttGridAttendVu
            // 
            this.MgmAttGridAttendVu.AllowUserToAddRows = false;
            this.MgmAttGridAttendVu.AllowUserToDeleteRows = false;
            this.MgmAttGridAttendVu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MgmAttGridAttendVu.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Detail_View});
            this.MgmAttGridAttendVu.Location = new System.Drawing.Point(10, 26);
            this.MgmAttGridAttendVu.Name = "MgmAttGridAttendVu";
            this.MgmAttGridAttendVu.ReadOnly = true;
            this.MgmAttGridAttendVu.Size = new System.Drawing.Size(552, 179);
            this.MgmAttGridAttendVu.TabIndex = 0;
            this.MgmAttGridAttendVu.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.MgmAttGridAttendVu_CellClick);
            // 
            // Detail_View
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.White;
            this.Detail_View.DefaultCellStyle = dataGridViewCellStyle1;
            this.Detail_View.HeaderText = "Detail View";
            this.Detail_View.Name = "Detail_View";
            this.Detail_View.ReadOnly = true;
            this.Detail_View.Text = "Check detail";
            this.Detail_View.UseColumnTextForButtonValue = true;
            // 
            // TabMgmGeneral
            // 
            this.TabMgmGeneral.BackColor = System.Drawing.Color.AliceBlue;
            this.TabMgmGeneral.Controls.Add(this.label23);
            this.TabMgmGeneral.Controls.Add(this.MgmGenGroup);
            this.TabMgmGeneral.Location = new System.Drawing.Point(4, 22);
            this.TabMgmGeneral.Name = "TabMgmGeneral";
            this.TabMgmGeneral.Size = new System.Drawing.Size(937, 551);
            this.TabMgmGeneral.TabIndex = 4;
            this.TabMgmGeneral.Text = "General";
            // 
            // label23
            // 
            this.label23.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label23.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label23.Location = new System.Drawing.Point(-4, 0);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(940, 24);
            this.label23.TabIndex = 82;
            this.label23.Text = "                                                 MIA LEAVE//HOLIDAYS/UPDATE STATU" +
                "S";
            // 
            // MgmGenGroup
            // 
            this.MgmGenGroup.Controls.Add(this.MgmGenLblFormid);
            this.MgmGenGroup.Controls.Add(this.MgmGenTxtFormid);
            this.MgmGenGroup.Controls.Add(this.groupBox7);
            this.MgmGenGroup.Controls.Add(this.MgmGenGroupLeave);
            this.MgmGenGroup.Controls.Add(this.MgmGenLblDepartment);
            this.MgmGenGroup.Controls.Add(this.MgmGenTxtDepartment);
            this.MgmGenGroup.Controls.Add(this.MgmGenPicture);
            this.MgmGenGroup.Controls.Add(this.MgmGenTxtBps);
            this.MgmGenGroup.Controls.Add(this.MgmGenLblBps);
            this.MgmGenGroup.Controls.Add(this.MgmGenLblName);
            this.MgmGenGroup.Controls.Add(this.MgmGenTxtName);
            this.MgmGenGroup.Location = new System.Drawing.Point(29, 51);
            this.MgmGenGroup.Name = "MgmGenGroup";
            this.MgmGenGroup.Size = new System.Drawing.Size(879, 473);
            this.MgmGenGroup.TabIndex = 0;
            this.MgmGenGroup.TabStop = false;
            this.MgmGenGroup.Text = "General Settings";
            // 
            // MgmGenLblFormid
            // 
            this.MgmGenLblFormid.AutoSize = true;
            this.MgmGenLblFormid.Location = new System.Drawing.Point(720, 162);
            this.MgmGenLblFormid.Name = "MgmGenLblFormid";
            this.MgmGenLblFormid.Size = new System.Drawing.Size(42, 13);
            this.MgmGenLblFormid.TabIndex = 40;
            this.MgmGenLblFormid.Text = "Form Id";
            // 
            // MgmGenTxtFormid
            // 
            this.MgmGenTxtFormid.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.MgmGenTxtFormid.Location = new System.Drawing.Point(768, 155);
            this.MgmGenTxtFormid.Name = "MgmGenTxtFormid";
            this.MgmGenTxtFormid.Size = new System.Drawing.Size(80, 20);
            this.MgmGenTxtFormid.TabIndex = 39;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.MgmGenTxtSearchbyformid);
            this.groupBox7.Controls.Add(this.MgmGenTxtSearchbyrf);
            this.groupBox7.Controls.Add(this.MgmGenBtnClear);
            this.groupBox7.Controls.Add(this.label36);
            this.groupBox7.Controls.Add(this.MgmGenBtnSearch);
            this.groupBox7.Controls.Add(this.MgmGenCkbSearchbyrf);
            this.groupBox7.Controls.Add(this.MgmGenCkbSearchbyformid);
            this.groupBox7.Location = new System.Drawing.Point(559, 180);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(289, 173);
            this.groupBox7.TabIndex = 38;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Search Options";
            // 
            // MgmGenTxtSearchbyformid
            // 
            this.MgmGenTxtSearchbyformid.AcceptsReturn = true;
            this.MgmGenTxtSearchbyformid.Location = new System.Drawing.Point(133, 90);
            this.MgmGenTxtSearchbyformid.Name = "MgmGenTxtSearchbyformid";
            this.MgmGenTxtSearchbyformid.Size = new System.Drawing.Size(122, 20);
            this.MgmGenTxtSearchbyformid.TabIndex = 31;
            this.MgmGenTxtSearchbyformid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmGenTxtSearchbyformid_KeyPress);
            // 
            // MgmGenTxtSearchbyrf
            // 
            this.MgmGenTxtSearchbyrf.Location = new System.Drawing.Point(133, 65);
            this.MgmGenTxtSearchbyrf.Name = "MgmGenTxtSearchbyrf";
            this.MgmGenTxtSearchbyrf.Size = new System.Drawing.Size(122, 20);
            this.MgmGenTxtSearchbyrf.TabIndex = 30;
            this.MgmGenTxtSearchbyrf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.MgmGenTxtSearchbyrf_KeyPress);
            // 
            // MgmGenBtnClear
            // 
            this.MgmGenBtnClear.Location = new System.Drawing.Point(193, 129);
            this.MgmGenBtnClear.Name = "MgmGenBtnClear";
            this.MgmGenBtnClear.Size = new System.Drawing.Size(61, 23);
            this.MgmGenBtnClear.TabIndex = 28;
            this.MgmGenBtnClear.Text = "Clear";
            this.MgmGenBtnClear.UseVisualStyleBackColor = true;
            this.MgmGenBtnClear.Click += new System.EventHandler(this.MgmGenBtnClear_Click);
            // 
            // label36
            // 
            this.label36.AutoSize = true;
            this.label36.ForeColor = System.Drawing.Color.Red;
            this.label36.Location = new System.Drawing.Point(12, 35);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(242, 13);
            this.label36.TabIndex = 25;
            this.label36.Text = "Following Search option can be used for a record.";
            // 
            // MgmGenBtnSearch
            // 
            this.MgmGenBtnSearch.Location = new System.Drawing.Point(132, 129);
            this.MgmGenBtnSearch.Name = "MgmGenBtnSearch";
            this.MgmGenBtnSearch.Size = new System.Drawing.Size(61, 23);
            this.MgmGenBtnSearch.TabIndex = 26;
            this.MgmGenBtnSearch.Text = "Search Record";
            this.MgmGenBtnSearch.UseVisualStyleBackColor = true;
            this.MgmGenBtnSearch.Click += new System.EventHandler(this.MgmGenBtnSearch_Click);
            // 
            // MgmGenCkbSearchbyrf
            // 
            this.MgmGenCkbSearchbyrf.AutoSize = true;
            this.MgmGenCkbSearchbyrf.Location = new System.Drawing.Point(15, 67);
            this.MgmGenCkbSearchbyrf.Name = "MgmGenCkbSearchbyrf";
            this.MgmGenCkbSearchbyrf.Size = new System.Drawing.Size(113, 17);
            this.MgmGenCkbSearchbyrf.TabIndex = 1;
            this.MgmGenCkbSearchbyrf.Text = "Search by RF Tag";
            this.MgmGenCkbSearchbyrf.UseVisualStyleBackColor = true;
            this.MgmGenCkbSearchbyrf.CheckStateChanged += new System.EventHandler(this.MgmGenCkbSearchbyrf_CheckStateChanged);
            // 
            // MgmGenCkbSearchbyformid
            // 
            this.MgmGenCkbSearchbyformid.AutoSize = true;
            this.MgmGenCkbSearchbyformid.Location = new System.Drawing.Point(15, 92);
            this.MgmGenCkbSearchbyformid.Name = "MgmGenCkbSearchbyformid";
            this.MgmGenCkbSearchbyformid.Size = new System.Drawing.Size(112, 17);
            this.MgmGenCkbSearchbyformid.TabIndex = 0;
            this.MgmGenCkbSearchbyformid.Text = "Search by Form Id";
            this.MgmGenCkbSearchbyformid.UseVisualStyleBackColor = true;
            this.MgmGenCkbSearchbyformid.CheckStateChanged += new System.EventHandler(this.MgmGenCkbSearchbyformid_CheckStateChanged);
            // 
            // MgmGenGroupLeave
            // 
            this.MgmGenGroupLeave.Controls.Add(this.MgmGenGroupBoxStatus);
            this.MgmGenGroupLeave.Controls.Add(this.MgmGenBtnAllow);
            this.MgmGenGroupLeave.Controls.Add(this.MgmGenGroupBoxLeave);
            this.MgmGenGroupLeave.Controls.Add(this.MgmGenGroupboxSetholiday);
            this.MgmGenGroupLeave.Controls.Add(this.MgmGenCkbSetstatus);
            this.MgmGenGroupLeave.Controls.Add(this.MgmGenCkbSetholidays);
            this.MgmGenGroupLeave.Controls.Add(this.MgmGenCkbSetleave);
            this.MgmGenGroupLeave.Location = new System.Drawing.Point(33, 180);
            this.MgmGenGroupLeave.Name = "MgmGenGroupLeave";
            this.MgmGenGroupLeave.Size = new System.Drawing.Size(463, 234);
            this.MgmGenGroupLeave.TabIndex = 31;
            this.MgmGenGroupLeave.TabStop = false;
            this.MgmGenGroupLeave.Text = "Leave";
            // 
            // MgmGenGroupBoxStatus
            // 
            this.MgmGenGroupBoxStatus.Controls.Add(this.MgmGenLblSelectdate);
            this.MgmGenGroupBoxStatus.Controls.Add(this.label34);
            this.MgmGenGroupBoxStatus.Controls.Add(this.MgmGenLblStatus);
            this.MgmGenGroupBoxStatus.Controls.Add(this.MgmGenCmbStatus);
            this.MgmGenGroupBoxStatus.Controls.Add(this.MgmGenDtTime);
            this.MgmGenGroupBoxStatus.Location = new System.Drawing.Point(27, 67);
            this.MgmGenGroupBoxStatus.Name = "MgmGenGroupBoxStatus";
            this.MgmGenGroupBoxStatus.Size = new System.Drawing.Size(410, 80);
            this.MgmGenGroupBoxStatus.TabIndex = 39;
            this.MgmGenGroupBoxStatus.TabStop = false;
            this.MgmGenGroupBoxStatus.Text = "Set Status";
            // 
            // MgmGenLblSelectdate
            // 
            this.MgmGenLblSelectdate.AutoSize = true;
            this.MgmGenLblSelectdate.Location = new System.Drawing.Point(7, 44);
            this.MgmGenLblSelectdate.Name = "MgmGenLblSelectdate";
            this.MgmGenLblSelectdate.Size = new System.Drawing.Size(75, 13);
            this.MgmGenLblSelectdate.TabIndex = 30;
            this.MgmGenLblSelectdate.Text = "Selected Date";
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Location = new System.Drawing.Point(7, 21);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(188, 13);
            this.label34.TabIndex = 29;
            this.label34.Text = "You can re-set Status of an Employee.";
            // 
            // MgmGenLblStatus
            // 
            this.MgmGenLblStatus.AutoSize = true;
            this.MgmGenLblStatus.Location = new System.Drawing.Point(229, 48);
            this.MgmGenLblStatus.Name = "MgmGenLblStatus";
            this.MgmGenLblStatus.Size = new System.Drawing.Size(37, 13);
            this.MgmGenLblStatus.TabIndex = 28;
            this.MgmGenLblStatus.Text = "Status";
            // 
            // MgmGenCmbStatus
            // 
            this.MgmGenCmbStatus.FormattingEnabled = true;
            this.MgmGenCmbStatus.Items.AddRange(new object[] {
            "ABSENT",
            "LEAVE",
            "IN",
            "OUT"});
            this.MgmGenCmbStatus.Location = new System.Drawing.Point(272, 44);
            this.MgmGenCmbStatus.Name = "MgmGenCmbStatus";
            this.MgmGenCmbStatus.Size = new System.Drawing.Size(121, 21);
            this.MgmGenCmbStatus.TabIndex = 27;
            // 
            // MgmGenDtTime
            // 
            this.MgmGenDtTime.Location = new System.Drawing.Point(88, 41);
            this.MgmGenDtTime.Name = "MgmGenDtTime";
            this.MgmGenDtTime.Size = new System.Drawing.Size(120, 20);
            this.MgmGenDtTime.TabIndex = 26;
            // 
            // MgmGenBtnAllow
            // 
            this.MgmGenBtnAllow.Location = new System.Drawing.Point(382, 196);
            this.MgmGenBtnAllow.Name = "MgmGenBtnAllow";
            this.MgmGenBtnAllow.Size = new System.Drawing.Size(61, 23);
            this.MgmGenBtnAllow.TabIndex = 32;
            this.MgmGenBtnAllow.Text = "Allow";
            this.MgmGenBtnAllow.UseVisualStyleBackColor = true;
            this.MgmGenBtnAllow.Click += new System.EventHandler(this.MgmGenBtnAllow_Click);
            // 
            // MgmGenGroupBoxLeave
            // 
            this.MgmGenGroupBoxLeave.Controls.Add(this.label24);
            this.MgmGenGroupBoxLeave.Controls.Add(this.MgmGenTxtLeavedates);
            this.MgmGenGroupBoxLeave.Controls.Add(this.numericUpDown1);
            this.MgmGenGroupBoxLeave.Controls.Add(this.MgmGenCtlDate1);
            this.MgmGenGroupBoxLeave.Controls.Add(this.RecepLblLeavedates);
            this.MgmGenGroupBoxLeave.Controls.Add(this.label26);
            this.MgmGenGroupBoxLeave.Controls.Add(this.label27);
            this.MgmGenGroupBoxLeave.Location = new System.Drawing.Point(33, 51);
            this.MgmGenGroupBoxLeave.Name = "MgmGenGroupBoxLeave";
            this.MgmGenGroupBoxLeave.Size = new System.Drawing.Size(410, 101);
            this.MgmGenGroupBoxLeave.TabIndex = 38;
            this.MgmGenGroupBoxLeave.TabStop = false;
            this.MgmGenGroupBoxLeave.Text = "Set Leave";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(7, 21);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(181, 13);
            this.label24.TabIndex = 30;
            this.label24.Text = "You can allow leave to an Emplioyee";
            // 
            // MgmGenTxtLeavedates
            // 
            this.MgmGenTxtLeavedates.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.MgmGenTxtLeavedates.Enabled = false;
            this.MgmGenTxtLeavedates.Location = new System.Drawing.Point(52, 73);
            this.MgmGenTxtLeavedates.Name = "MgmGenTxtLeavedates";
            this.MgmGenTxtLeavedates.Size = new System.Drawing.Size(341, 20);
            this.MgmGenTxtLeavedates.TabIndex = 22;
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.AllowDrop = true;
            this.numericUpDown1.Location = new System.Drawing.Point(302, 46);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(90, 20);
            this.numericUpDown1.TabIndex = 20;
            this.numericUpDown1.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // MgmGenCtlDate1
            // 
            this.MgmGenCtlDate1.Location = new System.Drawing.Point(52, 46);
            this.MgmGenCtlDate1.Name = "MgmGenCtlDate1";
            this.MgmGenCtlDate1.Size = new System.Drawing.Size(128, 20);
            this.MgmGenCtlDate1.TabIndex = 26;
            // 
            // RecepLblLeavedates
            // 
            this.RecepLblLeavedates.AutoSize = true;
            this.RecepLblLeavedates.Location = new System.Drawing.Point(7, 77);
            this.RecepLblLeavedates.Name = "RecepLblLeavedates";
            this.RecepLblLeavedates.Size = new System.Drawing.Size(37, 13);
            this.RecepLblLeavedates.TabIndex = 28;
            this.RecepLblLeavedates.Text = "Period";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(7, 50);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(30, 13);
            this.label26.TabIndex = 24;
            this.label26.Text = "From";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(223, 50);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(77, 13);
            this.label27.TabIndex = 27;
            this.label27.Text = "Required Days";
            // 
            // MgmGenGroupboxSetholiday
            // 
            this.MgmGenGroupboxSetholiday.Controls.Add(this.label30);
            this.MgmGenGroupboxSetholiday.Controls.Add(this.MgmGenLblReason);
            this.MgmGenGroupboxSetholiday.Controls.Add(this.MgmGenCmbReason);
            this.MgmGenGroupboxSetholiday.Controls.Add(this.MgmGenDttimeHoliday);
            this.MgmGenGroupboxSetholiday.Controls.Add(this.MgmGenLblSelecteddate);
            this.MgmGenGroupboxSetholiday.Location = new System.Drawing.Point(27, 86);
            this.MgmGenGroupboxSetholiday.Name = "MgmGenGroupboxSetholiday";
            this.MgmGenGroupboxSetholiday.Size = new System.Drawing.Size(410, 80);
            this.MgmGenGroupboxSetholiday.TabIndex = 40;
            this.MgmGenGroupboxSetholiday.TabStop = false;
            this.MgmGenGroupboxSetholiday.Text = "Set Holiday";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Location = new System.Drawing.Point(7, 21);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(191, 13);
            this.label30.TabIndex = 29;
            this.label30.Text = "You can set Holiday on Particular date.";
            // 
            // MgmGenLblReason
            // 
            this.MgmGenLblReason.AutoSize = true;
            this.MgmGenLblReason.Location = new System.Drawing.Point(235, 48);
            this.MgmGenLblReason.Name = "MgmGenLblReason";
            this.MgmGenLblReason.Size = new System.Drawing.Size(44, 13);
            this.MgmGenLblReason.TabIndex = 28;
            this.MgmGenLblReason.Text = "Reason";
            // 
            // MgmGenCmbReason
            // 
            this.MgmGenCmbReason.FormattingEnabled = true;
            this.MgmGenCmbReason.Items.AddRange(new object[] {
            "OFFICIAL HOLIDAY"});
            this.MgmGenCmbReason.Location = new System.Drawing.Point(281, 44);
            this.MgmGenCmbReason.Name = "MgmGenCmbReason";
            this.MgmGenCmbReason.Size = new System.Drawing.Size(121, 21);
            this.MgmGenCmbReason.TabIndex = 27;
            // 
            // MgmGenDttimeHoliday
            // 
            this.MgmGenDttimeHoliday.Location = new System.Drawing.Point(86, 44);
            this.MgmGenDttimeHoliday.Name = "MgmGenDttimeHoliday";
            this.MgmGenDttimeHoliday.Size = new System.Drawing.Size(122, 20);
            this.MgmGenDttimeHoliday.TabIndex = 26;
            // 
            // MgmGenLblSelecteddate
            // 
            this.MgmGenLblSelecteddate.AutoSize = true;
            this.MgmGenLblSelecteddate.Location = new System.Drawing.Point(7, 48);
            this.MgmGenLblSelecteddate.Name = "MgmGenLblSelecteddate";
            this.MgmGenLblSelecteddate.Size = new System.Drawing.Size(75, 13);
            this.MgmGenLblSelecteddate.TabIndex = 24;
            this.MgmGenLblSelecteddate.Text = "Selected Date";
            // 
            // MgmGenCkbSetstatus
            // 
            this.MgmGenCkbSetstatus.AutoSize = true;
            this.MgmGenCkbSetstatus.Location = new System.Drawing.Point(132, 27);
            this.MgmGenCkbSetstatus.Name = "MgmGenCkbSetstatus";
            this.MgmGenCkbSetstatus.Size = new System.Drawing.Size(75, 17);
            this.MgmGenCkbSetstatus.TabIndex = 41;
            this.MgmGenCkbSetstatus.Text = "Set Status";
            this.MgmGenCkbSetstatus.UseVisualStyleBackColor = true;
            this.MgmGenCkbSetstatus.CheckStateChanged += new System.EventHandler(this.MgmGenCkbSetstatus_CheckStateChanged);
            // 
            // MgmGenCkbSetholidays
            // 
            this.MgmGenCkbSetholidays.AutoSize = true;
            this.MgmGenCkbSetholidays.Location = new System.Drawing.Point(233, 27);
            this.MgmGenCkbSetholidays.Name = "MgmGenCkbSetholidays";
            this.MgmGenCkbSetholidays.Size = new System.Drawing.Size(85, 17);
            this.MgmGenCkbSetholidays.TabIndex = 40;
            this.MgmGenCkbSetholidays.Text = "Set Holidays";
            this.MgmGenCkbSetholidays.UseVisualStyleBackColor = true;
            this.MgmGenCkbSetholidays.CheckStateChanged += new System.EventHandler(this.MgmGenCkbSetholidays_CheckStateChanged);
            // 
            // MgmGenCkbSetleave
            // 
            this.MgmGenCkbSetleave.AutoSize = true;
            this.MgmGenCkbSetleave.Location = new System.Drawing.Point(27, 27);
            this.MgmGenCkbSetleave.Name = "MgmGenCkbSetleave";
            this.MgmGenCkbSetleave.Size = new System.Drawing.Size(75, 17);
            this.MgmGenCkbSetleave.TabIndex = 39;
            this.MgmGenCkbSetleave.Text = "Set Leave";
            this.MgmGenCkbSetleave.UseVisualStyleBackColor = true;
            this.MgmGenCkbSetleave.CheckStateChanged += new System.EventHandler(this.MgmGenCkbSetleave_CheckStateChanged);
            // 
            // MgmGenLblDepartment
            // 
            this.MgmGenLblDepartment.AutoSize = true;
            this.MgmGenLblDepartment.Location = new System.Drawing.Point(30, 99);
            this.MgmGenLblDepartment.Name = "MgmGenLblDepartment";
            this.MgmGenLblDepartment.Size = new System.Drawing.Size(62, 13);
            this.MgmGenLblDepartment.TabIndex = 29;
            this.MgmGenLblDepartment.Text = "Department";
            // 
            // MgmGenTxtDepartment
            // 
            this.MgmGenTxtDepartment.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.MgmGenTxtDepartment.Location = new System.Drawing.Point(100, 99);
            this.MgmGenTxtDepartment.Name = "MgmGenTxtDepartment";
            this.MgmGenTxtDepartment.Size = new System.Drawing.Size(200, 20);
            this.MgmGenTxtDepartment.TabIndex = 30;
            // 
            // MgmGenPicture
            // 
            this.MgmGenPicture.BackColor = System.Drawing.Color.LightGray;
            this.MgmGenPicture.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.MgmGenPicture.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.MgmGenPicture.Location = new System.Drawing.Point(723, 19);
            this.MgmGenPicture.Name = "MgmGenPicture";
            this.MgmGenPicture.Size = new System.Drawing.Size(125, 130);
            this.MgmGenPicture.TabIndex = 24;
            this.MgmGenPicture.TabStop = false;
            // 
            // MgmGenTxtBps
            // 
            this.MgmGenTxtBps.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.MgmGenTxtBps.Location = new System.Drawing.Point(100, 73);
            this.MgmGenTxtBps.Name = "MgmGenTxtBps";
            this.MgmGenTxtBps.Size = new System.Drawing.Size(80, 20);
            this.MgmGenTxtBps.TabIndex = 28;
            // 
            // MgmGenLblBps
            // 
            this.MgmGenLblBps.AutoSize = true;
            this.MgmGenLblBps.Location = new System.Drawing.Point(30, 76);
            this.MgmGenLblBps.Name = "MgmGenLblBps";
            this.MgmGenLblBps.Size = new System.Drawing.Size(25, 13);
            this.MgmGenLblBps.TabIndex = 27;
            this.MgmGenLblBps.Text = "Bps";
            // 
            // MgmGenLblName
            // 
            this.MgmGenLblName.AutoSize = true;
            this.MgmGenLblName.Location = new System.Drawing.Point(30, 46);
            this.MgmGenLblName.Name = "MgmGenLblName";
            this.MgmGenLblName.Size = new System.Drawing.Size(35, 13);
            this.MgmGenLblName.TabIndex = 25;
            this.MgmGenLblName.Text = "Name";
            // 
            // MgmGenTxtName
            // 
            this.MgmGenTxtName.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.MgmGenTxtName.Location = new System.Drawing.Point(100, 46);
            this.MgmGenTxtName.Name = "MgmGenTxtName";
            this.MgmGenTxtName.Size = new System.Drawing.Size(200, 20);
            this.MgmGenTxtName.TabIndex = 26;
            // 
            // TabEmp
            // 
            this.TabEmp.BackColor = System.Drawing.SystemColors.Desktop;
            this.TabEmp.Controls.Add(this.button4);
            this.TabEmp.Controls.Add(this.Employees);
            this.TabEmp.Location = new System.Drawing.Point(4, 22);
            this.TabEmp.Name = "TabEmp";
            this.TabEmp.Size = new System.Drawing.Size(994, 630);
            this.TabEmp.TabIndex = 2;
            this.TabEmp.Text = "Employees";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(893, 605);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(74, 23);
            this.button4.TabIndex = 17;
            this.button4.Text = "Log out";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Employees
            // 
            this.Employees.Controls.Add(this.TabGeneral);
            this.Employees.Controls.Add(this.TabMonthlySettings);
            this.Employees.Controls.Add(this.TabSalary);
            this.Employees.Controls.Add(this.TabAttendance);
            this.Employees.Location = new System.Drawing.Point(30, 16);
            this.Employees.Name = "Employees";
            this.Employees.SelectedIndex = 0;
            this.Employees.Size = new System.Drawing.Size(937, 590);
            this.Employees.TabIndex = 6;
            // 
            // TabGeneral
            // 
            this.TabGeneral.BackColor = System.Drawing.Color.AliceBlue;
            this.TabGeneral.Controls.Add(this.label35);
            this.TabGeneral.Controls.Add(this.EmpGenTxtUpdateby);
            this.TabGeneral.Controls.Add(this.EmpGenLblUpdateby);
            this.TabGeneral.Controls.Add(this.EmpGenLblRfid);
            this.TabGeneral.Controls.Add(this.EmpGenTxtRfid);
            this.TabGeneral.Controls.Add(this.EmpGenLblHiredate);
            this.TabGeneral.Controls.Add(this.EmpGenTxtHiredate);
            this.TabGeneral.Controls.Add(this.EmpGenLblBnkbranch);
            this.TabGeneral.Controls.Add(this.EmpGenTxtBnkbranch);
            this.TabGeneral.Controls.Add(this.EmpGenLblBnkacc);
            this.TabGeneral.Controls.Add(this.EmpGenTxtBnkacc);
            this.TabGeneral.Controls.Add(this.EmpGenGroupSearch);
            this.TabGeneral.Controls.Add(this.EmpGenLblPic);
            this.TabGeneral.Controls.Add(this.EmpGenLblTempadd);
            this.TabGeneral.Controls.Add(this.EmpGenTxtTempadd);
            this.TabGeneral.Controls.Add(this.EmpGenLblPermadd);
            this.TabGeneral.Controls.Add(this.EmpGenTxtPermadd);
            this.TabGeneral.Controls.Add(this.EmpGenLblPhno);
            this.TabGeneral.Controls.Add(this.EmpGenTxtPhno);
            this.TabGeneral.Controls.Add(this.EmpGenLblEmail);
            this.TabGeneral.Controls.Add(this.EmpGenTxtEmail);
            this.TabGeneral.Controls.Add(this.EmpGenLblDomicile);
            this.TabGeneral.Controls.Add(this.EmpGenTxtDomicile);
            this.TabGeneral.Controls.Add(this.EmpGenLblReligion);
            this.TabGeneral.Controls.Add(this.EmpGenTxtReligion);
            this.TabGeneral.Controls.Add(this.EmpGenLblDob);
            this.TabGeneral.Controls.Add(this.EmpGenLblCnic);
            this.TabGeneral.Controls.Add(this.EmpGenLblFname);
            this.TabGeneral.Controls.Add(this.EmpGenLblEname);
            this.TabGeneral.Controls.Add(this.EmpGenTxtCnic);
            this.TabGeneral.Controls.Add(this.EmpGenTxtDob);
            this.TabGeneral.Controls.Add(this.EmpGenTxtFname);
            this.TabGeneral.Controls.Add(this.EmpGentxtEname);
            this.TabGeneral.Controls.Add(this.EmpGenBtnBrowse);
            this.TabGeneral.Controls.Add(this.EmpGenPicbox);
            this.TabGeneral.Controls.Add(this.EmpGenGroupEmpaccount);
            this.TabGeneral.Location = new System.Drawing.Point(4, 22);
            this.TabGeneral.Name = "TabGeneral";
            this.TabGeneral.Padding = new System.Windows.Forms.Padding(3);
            this.TabGeneral.Size = new System.Drawing.Size(929, 564);
            this.TabGeneral.TabIndex = 1;
            this.TabGeneral.Text = "General";
            // 
            // label35
            // 
            this.label35.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label35.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label35.Location = new System.Drawing.Point(-4, 0);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(940, 24);
            this.label35.TabIndex = 86;
            this.label35.Text = "                                                 MIA EMPLOYEE GENERAL SETTINGS";
            // 
            // EmpGenTxtUpdateby
            // 
            this.EmpGenTxtUpdateby.AcceptsReturn = true;
            this.EmpGenTxtUpdateby.Location = new System.Drawing.Point(760, 216);
            this.EmpGenTxtUpdateby.Name = "EmpGenTxtUpdateby";
            this.EmpGenTxtUpdateby.ReadOnly = true;
            this.EmpGenTxtUpdateby.Size = new System.Drawing.Size(116, 20);
            this.EmpGenTxtUpdateby.TabIndex = 34;
            this.EmpGenTxtUpdateby.Visible = false;
            // 
            // EmpGenLblUpdateby
            // 
            this.EmpGenLblUpdateby.AutoSize = true;
            this.EmpGenLblUpdateby.Location = new System.Drawing.Point(692, 219);
            this.EmpGenLblUpdateby.Name = "EmpGenLblUpdateby";
            this.EmpGenLblUpdateby.Size = new System.Drawing.Size(62, 13);
            this.EmpGenLblUpdateby.TabIndex = 33;
            this.EmpGenLblUpdateby.Text = "Updated by";
            this.EmpGenLblUpdateby.Visible = false;
            // 
            // EmpGenLblRfid
            // 
            this.EmpGenLblRfid.AutoSize = true;
            this.EmpGenLblRfid.Location = new System.Drawing.Point(310, 145);
            this.EmpGenLblRfid.Name = "EmpGenLblRfid";
            this.EmpGenLblRfid.Size = new System.Drawing.Size(55, 13);
            this.EmpGenLblRfid.TabIndex = 32;
            this.EmpGenLblRfid.Text = "RF Tag Id";
            // 
            // EmpGenTxtRfid
            // 
            this.EmpGenTxtRfid.Location = new System.Drawing.Point(397, 141);
            this.EmpGenTxtRfid.Name = "EmpGenTxtRfid";
            this.EmpGenTxtRfid.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtRfid.TabIndex = 31;
            // 
            // EmpGenLblHiredate
            // 
            this.EmpGenLblHiredate.AutoSize = true;
            this.EmpGenLblHiredate.Location = new System.Drawing.Point(35, 145);
            this.EmpGenLblHiredate.Name = "EmpGenLblHiredate";
            this.EmpGenLblHiredate.Size = new System.Drawing.Size(52, 13);
            this.EmpGenLblHiredate.TabIndex = 30;
            this.EmpGenLblHiredate.Text = "Hire Date";
            // 
            // EmpGenTxtHiredate
            // 
            this.EmpGenTxtHiredate.Location = new System.Drawing.Point(146, 141);
            this.EmpGenTxtHiredate.Name = "EmpGenTxtHiredate";
            this.EmpGenTxtHiredate.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtHiredate.TabIndex = 29;
            // 
            // EmpGenLblBnkbranch
            // 
            this.EmpGenLblBnkbranch.AutoSize = true;
            this.EmpGenLblBnkbranch.Location = new System.Drawing.Point(310, 203);
            this.EmpGenLblBnkbranch.Name = "EmpGenLblBnkbranch";
            this.EmpGenLblBnkbranch.Size = new System.Drawing.Size(69, 13);
            this.EmpGenLblBnkbranch.TabIndex = 28;
            this.EmpGenLblBnkbranch.Text = "Bank Branch";
            // 
            // EmpGenTxtBnkbranch
            // 
            this.EmpGenTxtBnkbranch.Location = new System.Drawing.Point(397, 199);
            this.EmpGenTxtBnkbranch.Name = "EmpGenTxtBnkbranch";
            this.EmpGenTxtBnkbranch.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtBnkbranch.TabIndex = 27;
            this.EmpGenTxtBnkbranch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtBnkbranch_KeyPress);
            // 
            // EmpGenLblBnkacc
            // 
            this.EmpGenLblBnkacc.AutoSize = true;
            this.EmpGenLblBnkacc.Location = new System.Drawing.Point(35, 203);
            this.EmpGenLblBnkacc.Name = "EmpGenLblBnkacc";
            this.EmpGenLblBnkacc.Size = new System.Drawing.Size(66, 13);
            this.EmpGenLblBnkacc.TabIndex = 26;
            this.EmpGenLblBnkacc.Text = "Bank Acc Id";
            // 
            // EmpGenTxtBnkacc
            // 
            this.EmpGenTxtBnkacc.Location = new System.Drawing.Point(146, 199);
            this.EmpGenTxtBnkacc.Name = "EmpGenTxtBnkacc";
            this.EmpGenTxtBnkacc.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtBnkacc.TabIndex = 25;
            this.EmpGenTxtBnkacc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtBnkacc_KeyPress);
            // 
            // EmpGenGroupSearch
            // 
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenBtnEdit);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenBtnCancel);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenLblSearchoption);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenCkbCnic);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenBtnSearch);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenBtnSave);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenCkbRfid);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenLblSearchrec);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenCkbFormid);
            this.EmpGenGroupSearch.Controls.Add(this.EmpGenTxtSearchrec);
            this.EmpGenGroupSearch.Location = new System.Drawing.Point(454, 320);
            this.EmpGenGroupSearch.Name = "EmpGenGroupSearch";
            this.EmpGenGroupSearch.Size = new System.Drawing.Size(422, 225);
            this.EmpGenGroupSearch.TabIndex = 1;
            this.EmpGenGroupSearch.TabStop = false;
            this.EmpGenGroupSearch.Text = "Search Options";
            // 
            // EmpGenBtnEdit
            // 
            this.EmpGenBtnEdit.Enabled = false;
            this.EmpGenBtnEdit.Location = new System.Drawing.Point(314, 143);
            this.EmpGenBtnEdit.Name = "EmpGenBtnEdit";
            this.EmpGenBtnEdit.Size = new System.Drawing.Size(74, 23);
            this.EmpGenBtnEdit.TabIndex = 27;
            this.EmpGenBtnEdit.Text = "Edit";
            this.EmpGenBtnEdit.UseVisualStyleBackColor = true;
            this.EmpGenBtnEdit.Click += new System.EventHandler(this.EmpGenBtnEdit_Click);
            // 
            // EmpGenBtnCancel
            // 
            this.EmpGenBtnCancel.Enabled = false;
            this.EmpGenBtnCancel.Location = new System.Drawing.Point(314, 169);
            this.EmpGenBtnCancel.Name = "EmpGenBtnCancel";
            this.EmpGenBtnCancel.Size = new System.Drawing.Size(74, 23);
            this.EmpGenBtnCancel.TabIndex = 28;
            this.EmpGenBtnCancel.Text = "Cancel";
            this.EmpGenBtnCancel.UseVisualStyleBackColor = true;
            this.EmpGenBtnCancel.Click += new System.EventHandler(this.EmpGenBtnCancel_Click);
            // 
            // EmpGenLblSearchoption
            // 
            this.EmpGenLblSearchoption.AutoSize = true;
            this.EmpGenLblSearchoption.Location = new System.Drawing.Point(21, 30);
            this.EmpGenLblSearchoption.Name = "EmpGenLblSearchoption";
            this.EmpGenLblSearchoption.Size = new System.Drawing.Size(242, 13);
            this.EmpGenLblSearchoption.TabIndex = 25;
            this.EmpGenLblSearchoption.Text = "Following Search option can be used for a record.";
            // 
            // EmpGenCkbCnic
            // 
            this.EmpGenCkbCnic.AutoSize = true;
            this.EmpGenCkbCnic.Location = new System.Drawing.Point(25, 114);
            this.EmpGenCkbCnic.Name = "EmpGenCkbCnic";
            this.EmpGenCkbCnic.Size = new System.Drawing.Size(102, 17);
            this.EmpGenCkbCnic.TabIndex = 2;
            this.EmpGenCkbCnic.Text = "Search by CNIC";
            this.EmpGenCkbCnic.UseVisualStyleBackColor = true;
            this.EmpGenCkbCnic.CheckStateChanged += new System.EventHandler(this.GenCkCnic_CheckStateChanged);
            // 
            // EmpGenBtnSearch
            // 
            this.EmpGenBtnSearch.Location = new System.Drawing.Point(235, 143);
            this.EmpGenBtnSearch.Name = "EmpGenBtnSearch";
            this.EmpGenBtnSearch.Size = new System.Drawing.Size(74, 23);
            this.EmpGenBtnSearch.TabIndex = 26;
            this.EmpGenBtnSearch.Text = "Search Record";
            this.EmpGenBtnSearch.UseVisualStyleBackColor = true;
            this.EmpGenBtnSearch.Click += new System.EventHandler(this.GenBtnSearchrec_Click);
            // 
            // EmpGenBtnSave
            // 
            this.EmpGenBtnSave.Enabled = false;
            this.EmpGenBtnSave.Location = new System.Drawing.Point(234, 169);
            this.EmpGenBtnSave.Name = "EmpGenBtnSave";
            this.EmpGenBtnSave.Size = new System.Drawing.Size(74, 23);
            this.EmpGenBtnSave.TabIndex = 2;
            this.EmpGenBtnSave.Text = "Save";
            this.EmpGenBtnSave.UseVisualStyleBackColor = true;
            this.EmpGenBtnSave.Click += new System.EventHandler(this.GenBtnSave_Click);
            // 
            // EmpGenCkbRfid
            // 
            this.EmpGenCkbRfid.AutoSize = true;
            this.EmpGenCkbRfid.ForeColor = System.Drawing.Color.Black;
            this.EmpGenCkbRfid.Location = new System.Drawing.Point(25, 87);
            this.EmpGenCkbRfid.Name = "EmpGenCkbRfid";
            this.EmpGenCkbRfid.Size = new System.Drawing.Size(113, 17);
            this.EmpGenCkbRfid.TabIndex = 1;
            this.EmpGenCkbRfid.Text = "Search by RF Tag";
            this.EmpGenCkbRfid.UseVisualStyleBackColor = true;
            this.EmpGenCkbRfid.CheckStateChanged += new System.EventHandler(this.GenCkRfid_CheckStateChanged);
            // 
            // EmpGenLblSearchrec
            // 
            this.EmpGenLblSearchrec.AutoSize = true;
            this.EmpGenLblSearchrec.Location = new System.Drawing.Point(147, 116);
            this.EmpGenLblSearchrec.Name = "EmpGenLblSearchrec";
            this.EmpGenLblSearchrec.Size = new System.Drawing.Size(79, 13);
            this.EmpGenLblSearchrec.TabIndex = 25;
            this.EmpGenLblSearchrec.Text = "Search Record";
            // 
            // EmpGenCkbFormid
            // 
            this.EmpGenCkbFormid.AutoSize = true;
            this.EmpGenCkbFormid.Enabled = false;
            this.EmpGenCkbFormid.Location = new System.Drawing.Point(25, 60);
            this.EmpGenCkbFormid.Name = "EmpGenCkbFormid";
            this.EmpGenCkbFormid.Size = new System.Drawing.Size(112, 17);
            this.EmpGenCkbFormid.TabIndex = 0;
            this.EmpGenCkbFormid.Text = "Search by Form Id";
            this.EmpGenCkbFormid.UseVisualStyleBackColor = true;
            this.EmpGenCkbFormid.Visible = false;
            this.EmpGenCkbFormid.CheckStateChanged += new System.EventHandler(this.GenCkformid_CheckStateChanged);
            // 
            // EmpGenTxtSearchrec
            // 
            this.EmpGenTxtSearchrec.Location = new System.Drawing.Point(235, 112);
            this.EmpGenTxtSearchrec.Name = "EmpGenTxtSearchrec";
            this.EmpGenTxtSearchrec.Size = new System.Drawing.Size(153, 20);
            this.EmpGenTxtSearchrec.TabIndex = 23;
            this.EmpGenTxtSearchrec.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtSearchrec_KeyPress);
            // 
            // EmpGenLblPic
            // 
            this.EmpGenLblPic.AutoSize = true;
            this.EmpGenLblPic.Location = new System.Drawing.Point(714, 44);
            this.EmpGenLblPic.Name = "EmpGenLblPic";
            this.EmpGenLblPic.Size = new System.Drawing.Size(40, 13);
            this.EmpGenLblPic.TabIndex = 24;
            this.EmpGenLblPic.Text = "Picture";
            // 
            // EmpGenLblTempadd
            // 
            this.EmpGenLblTempadd.AutoSize = true;
            this.EmpGenLblTempadd.Location = new System.Drawing.Point(35, 265);
            this.EmpGenLblTempadd.Name = "EmpGenLblTempadd";
            this.EmpGenLblTempadd.Size = new System.Drawing.Size(98, 13);
            this.EmpGenLblTempadd.TabIndex = 24;
            this.EmpGenLblTempadd.Text = "Temporary Address";
            // 
            // EmpGenTxtTempadd
            // 
            this.EmpGenTxtTempadd.Location = new System.Drawing.Point(146, 261);
            this.EmpGenTxtTempadd.Name = "EmpGenTxtTempadd";
            this.EmpGenTxtTempadd.Size = new System.Drawing.Size(399, 20);
            this.EmpGenTxtTempadd.TabIndex = 23;
            this.EmpGenTxtTempadd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtTempadd_KeyPress);
            // 
            // EmpGenLblPermadd
            // 
            this.EmpGenLblPermadd.AutoSize = true;
            this.EmpGenLblPermadd.Location = new System.Drawing.Point(35, 292);
            this.EmpGenLblPermadd.Name = "EmpGenLblPermadd";
            this.EmpGenLblPermadd.Size = new System.Drawing.Size(99, 13);
            this.EmpGenLblPermadd.TabIndex = 22;
            this.EmpGenLblPermadd.Text = "Permanent Address";
            // 
            // EmpGenTxtPermadd
            // 
            this.EmpGenTxtPermadd.Location = new System.Drawing.Point(146, 288);
            this.EmpGenTxtPermadd.Name = "EmpGenTxtPermadd";
            this.EmpGenTxtPermadd.Size = new System.Drawing.Size(399, 20);
            this.EmpGenTxtPermadd.TabIndex = 21;
            this.EmpGenTxtPermadd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtPermadd_KeyPress);
            // 
            // EmpGenLblPhno
            // 
            this.EmpGenLblPhno.AutoSize = true;
            this.EmpGenLblPhno.Location = new System.Drawing.Point(310, 229);
            this.EmpGenLblPhno.Name = "EmpGenLblPhno";
            this.EmpGenLblPhno.Size = new System.Drawing.Size(55, 13);
            this.EmpGenLblPhno.TabIndex = 18;
            this.EmpGenLblPhno.Text = "Phone No";
            // 
            // EmpGenTxtPhno
            // 
            this.EmpGenTxtPhno.Location = new System.Drawing.Point(397, 225);
            this.EmpGenTxtPhno.Name = "EmpGenTxtPhno";
            this.EmpGenTxtPhno.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtPhno.TabIndex = 17;
            this.EmpGenTxtPhno.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtPhno_KeyPress);
            // 
            // EmpGenLblEmail
            // 
            this.EmpGenLblEmail.AutoSize = true;
            this.EmpGenLblEmail.Location = new System.Drawing.Point(35, 229);
            this.EmpGenLblEmail.Name = "EmpGenLblEmail";
            this.EmpGenLblEmail.Size = new System.Drawing.Size(73, 13);
            this.EmpGenLblEmail.TabIndex = 16;
            this.EmpGenLblEmail.Text = "Email Address";
            // 
            // EmpGenTxtEmail
            // 
            this.EmpGenTxtEmail.Location = new System.Drawing.Point(146, 225);
            this.EmpGenTxtEmail.Name = "EmpGenTxtEmail";
            this.EmpGenTxtEmail.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtEmail.TabIndex = 15;
            this.EmpGenTxtEmail.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtEmail_KeyPress);
            // 
            // EmpGenLblDomicile
            // 
            this.EmpGenLblDomicile.AutoSize = true;
            this.EmpGenLblDomicile.Location = new System.Drawing.Point(310, 120);
            this.EmpGenLblDomicile.Name = "EmpGenLblDomicile";
            this.EmpGenLblDomicile.Size = new System.Drawing.Size(47, 13);
            this.EmpGenLblDomicile.TabIndex = 14;
            this.EmpGenLblDomicile.Text = "Domicile";
            // 
            // EmpGenTxtDomicile
            // 
            this.EmpGenTxtDomicile.Location = new System.Drawing.Point(397, 116);
            this.EmpGenTxtDomicile.Name = "EmpGenTxtDomicile";
            this.EmpGenTxtDomicile.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtDomicile.TabIndex = 13;
            this.EmpGenTxtDomicile.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtDomicile_KeyPress);
            // 
            // EmpGenLblReligion
            // 
            this.EmpGenLblReligion.AutoSize = true;
            this.EmpGenLblReligion.Location = new System.Drawing.Point(35, 120);
            this.EmpGenLblReligion.Name = "EmpGenLblReligion";
            this.EmpGenLblReligion.Size = new System.Drawing.Size(45, 13);
            this.EmpGenLblReligion.TabIndex = 12;
            this.EmpGenLblReligion.Text = "Religion";
            // 
            // EmpGenTxtReligion
            // 
            this.EmpGenTxtReligion.Location = new System.Drawing.Point(146, 116);
            this.EmpGenTxtReligion.Name = "EmpGenTxtReligion";
            this.EmpGenTxtReligion.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtReligion.TabIndex = 11;
            this.EmpGenTxtReligion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtReligion_KeyPress);
            // 
            // EmpGenLblDob
            // 
            this.EmpGenLblDob.AutoSize = true;
            this.EmpGenLblDob.Location = new System.Drawing.Point(310, 89);
            this.EmpGenLblDob.Name = "EmpGenLblDob";
            this.EmpGenLblDob.Size = new System.Drawing.Size(66, 13);
            this.EmpGenLblDob.TabIndex = 10;
            this.EmpGenLblDob.Text = "Date of Birth";
            // 
            // EmpGenLblCnic
            // 
            this.EmpGenLblCnic.AutoSize = true;
            this.EmpGenLblCnic.Location = new System.Drawing.Point(35, 89);
            this.EmpGenLblCnic.Name = "EmpGenLblCnic";
            this.EmpGenLblCnic.Size = new System.Drawing.Size(32, 13);
            this.EmpGenLblCnic.TabIndex = 9;
            this.EmpGenLblCnic.Text = "CNIC";
            // 
            // EmpGenLblFname
            // 
            this.EmpGenLblFname.AutoSize = true;
            this.EmpGenLblFname.Location = new System.Drawing.Point(310, 64);
            this.EmpGenLblFname.Name = "EmpGenLblFname";
            this.EmpGenLblFname.Size = new System.Drawing.Size(68, 13);
            this.EmpGenLblFname.TabIndex = 8;
            this.EmpGenLblFname.Text = "Father Name";
            // 
            // EmpGenLblEname
            // 
            this.EmpGenLblEname.AutoSize = true;
            this.EmpGenLblEname.Location = new System.Drawing.Point(35, 64);
            this.EmpGenLblEname.Name = "EmpGenLblEname";
            this.EmpGenLblEname.Size = new System.Drawing.Size(84, 13);
            this.EmpGenLblEname.TabIndex = 7;
            this.EmpGenLblEname.Text = "Employee Name";
            // 
            // EmpGenTxtCnic
            // 
            this.EmpGenTxtCnic.Location = new System.Drawing.Point(146, 85);
            this.EmpGenTxtCnic.Name = "EmpGenTxtCnic";
            this.EmpGenTxtCnic.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtCnic.TabIndex = 6;
            this.EmpGenTxtCnic.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtCnic_KeyPress);
            // 
            // EmpGenTxtDob
            // 
            this.EmpGenTxtDob.Location = new System.Drawing.Point(397, 85);
            this.EmpGenTxtDob.Name = "EmpGenTxtDob";
            this.EmpGenTxtDob.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtDob.TabIndex = 5;
            // 
            // EmpGenTxtFname
            // 
            this.EmpGenTxtFname.Location = new System.Drawing.Point(397, 60);
            this.EmpGenTxtFname.Name = "EmpGenTxtFname";
            this.EmpGenTxtFname.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtFname.TabIndex = 4;
            this.EmpGenTxtFname.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtFname_KeyPress);
            // 
            // EmpGentxtEname
            // 
            this.EmpGentxtEname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpGentxtEname.Location = new System.Drawing.Point(146, 60);
            this.EmpGentxtEname.Name = "EmpGentxtEname";
            this.EmpGentxtEname.Size = new System.Drawing.Size(148, 20);
            this.EmpGentxtEname.TabIndex = 3;
            this.EmpGentxtEname.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGentxtEname_KeyPress);
            // 
            // EmpGenBtnBrowse
            // 
            this.EmpGenBtnBrowse.Location = new System.Drawing.Point(760, 190);
            this.EmpGenBtnBrowse.Name = "EmpGenBtnBrowse";
            this.EmpGenBtnBrowse.Size = new System.Drawing.Size(116, 21);
            this.EmpGenBtnBrowse.TabIndex = 2;
            this.EmpGenBtnBrowse.Text = "Browse";
            this.EmpGenBtnBrowse.UseVisualStyleBackColor = true;
            this.EmpGenBtnBrowse.Click += new System.EventHandler(this.EmpGenBtnBrowse_Click);
            // 
            // EmpGenPicbox
            // 
            this.EmpGenPicbox.BackColor = System.Drawing.Color.DimGray;
            this.EmpGenPicbox.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.EmpGenPicbox.Location = new System.Drawing.Point(760, 44);
            this.EmpGenPicbox.Name = "EmpGenPicbox";
            this.EmpGenPicbox.Size = new System.Drawing.Size(116, 143);
            this.EmpGenPicbox.TabIndex = 1;
            this.EmpGenPicbox.TabStop = false;
            // 
            // EmpGenGroupEmpaccount
            // 
            this.EmpGenGroupEmpaccount.Controls.Add(this.EmpGenTxtPwd);
            this.EmpGenGroupEmpaccount.Controls.Add(this.EmpGenLblPwd);
            this.EmpGenGroupEmpaccount.Controls.Add(this.groupBox1);
            this.EmpGenGroupEmpaccount.Controls.Add(this.label2);
            this.EmpGenGroupEmpaccount.Controls.Add(this.EmpGenTxtLogname);
            this.EmpGenGroupEmpaccount.Controls.Add(this.EmpGenLblLogname);
            this.EmpGenGroupEmpaccount.Location = new System.Drawing.Point(35, 320);
            this.EmpGenGroupEmpaccount.Name = "EmpGenGroupEmpaccount";
            this.EmpGenGroupEmpaccount.Size = new System.Drawing.Size(364, 225);
            this.EmpGenGroupEmpaccount.TabIndex = 0;
            this.EmpGenGroupEmpaccount.TabStop = false;
            this.EmpGenGroupEmpaccount.Text = "Employee Account";
            // 
            // EmpGenTxtPwd
            // 
            this.EmpGenTxtPwd.Location = new System.Drawing.Point(122, 84);
            this.EmpGenTxtPwd.Name = "EmpGenTxtPwd";
            this.EmpGenTxtPwd.PasswordChar = '*';
            this.EmpGenTxtPwd.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtPwd.TabIndex = 32;
            this.EmpGenTxtPwd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpGenTxtPwd_KeyPress);
            // 
            // EmpGenLblPwd
            // 
            this.EmpGenLblPwd.AutoSize = true;
            this.EmpGenLblPwd.Location = new System.Drawing.Point(27, 88);
            this.EmpGenLblPwd.Name = "EmpGenLblPwd";
            this.EmpGenLblPwd.Size = new System.Drawing.Size(53, 13);
            this.EmpGenLblPwd.TabIndex = 33;
            this.EmpGenLblPwd.Text = "Password";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.EmpGenLblAccstatus);
            this.groupBox1.Controls.Add(this.EmpGenRdbDisable);
            this.groupBox1.Controls.Add(this.EmpGenRdbEnable);
            this.groupBox1.Location = new System.Drawing.Point(27, 109);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(267, 103);
            this.groupBox1.TabIndex = 31;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Account Status";
            // 
            // EmpGenLblAccstatus
            // 
            this.EmpGenLblAccstatus.AutoSize = true;
            this.EmpGenLblAccstatus.Location = new System.Drawing.Point(7, 28);
            this.EmpGenLblAccstatus.Name = "EmpGenLblAccstatus";
            this.EmpGenLblAccstatus.Size = new System.Drawing.Size(234, 13);
            this.EmpGenLblAccstatus.TabIndex = 28;
            this.EmpGenLblAccstatus.Text = "Temporarily an Account can be Enable/Disable.";
            // 
            // EmpGenRdbDisable
            // 
            this.EmpGenRdbDisable.AutoSize = true;
            this.EmpGenRdbDisable.Location = new System.Drawing.Point(11, 74);
            this.EmpGenRdbDisable.Name = "EmpGenRdbDisable";
            this.EmpGenRdbDisable.Size = new System.Drawing.Size(60, 17);
            this.EmpGenRdbDisable.TabIndex = 30;
            this.EmpGenRdbDisable.TabStop = true;
            this.EmpGenRdbDisable.Text = "Disable";
            this.EmpGenRdbDisable.UseVisualStyleBackColor = true;
            // 
            // EmpGenRdbEnable
            // 
            this.EmpGenRdbEnable.AutoSize = true;
            this.EmpGenRdbEnable.Location = new System.Drawing.Point(11, 52);
            this.EmpGenRdbEnable.Name = "EmpGenRdbEnable";
            this.EmpGenRdbEnable.Size = new System.Drawing.Size(58, 17);
            this.EmpGenRdbEnable.TabIndex = 29;
            this.EmpGenRdbEnable.TabStop = true;
            this.EmpGenRdbEnable.Text = "Enable";
            this.EmpGenRdbEnable.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(210, 13);
            this.label2.TabIndex = 27;
            this.label2.Text = "This login name will be useful only on Web.";
            // 
            // EmpGenTxtLogname
            // 
            this.EmpGenTxtLogname.Location = new System.Drawing.Point(122, 60);
            this.EmpGenTxtLogname.Name = "EmpGenTxtLogname";
            this.EmpGenTxtLogname.Size = new System.Drawing.Size(148, 20);
            this.EmpGenTxtLogname.TabIndex = 19;
            // 
            // EmpGenLblLogname
            // 
            this.EmpGenLblLogname.AutoSize = true;
            this.EmpGenLblLogname.Location = new System.Drawing.Point(27, 64);
            this.EmpGenLblLogname.Name = "EmpGenLblLogname";
            this.EmpGenLblLogname.Size = new System.Drawing.Size(64, 13);
            this.EmpGenLblLogname.TabIndex = 20;
            this.EmpGenLblLogname.Text = "Login Name";
            // 
            // TabMonthlySettings
            // 
            this.TabMonthlySettings.BackColor = System.Drawing.Color.AliceBlue;
            this.TabMonthlySettings.Controls.Add(this.label33);
            this.TabMonthlySettings.Controls.Add(this.EmpMonGroupSet0);
            this.TabMonthlySettings.Controls.Add(this.EmpMonGridvu);
            this.TabMonthlySettings.Controls.Add(this.EmpMonGroupsearch);
            this.TabMonthlySettings.Controls.Add(this.EmpMonGroup2);
            this.TabMonthlySettings.Location = new System.Drawing.Point(4, 22);
            this.TabMonthlySettings.Name = "TabMonthlySettings";
            this.TabMonthlySettings.Size = new System.Drawing.Size(929, 564);
            this.TabMonthlySettings.TabIndex = 3;
            this.TabMonthlySettings.Text = "Monthly Settings";
            // 
            // label33
            // 
            this.label33.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label33.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label33.Location = new System.Drawing.Point(-4, 0);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(940, 24);
            this.label33.TabIndex = 85;
            this.label33.Text = "                                                 MIA MONTHLY ALLOWANCES VIEW";
            // 
            // EmpMonGroupSet0
            // 
            this.EmpMonGroupSet0.BackColor = System.Drawing.Color.AliceBlue;
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonGroup1);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonTxtName);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonLblAllowtotal);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonLblJobid);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonTxtDedtotal);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonTxtJobid);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonTxtAllowtotal);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonBtnApply);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonLblName);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonLblDedtotal);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonTxtTotalamount);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonBtnCancel0);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonTxtBps);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonLblBps);
            this.EmpMonGroupSet0.Controls.Add(this.EmpMonLblTotalamount);
            this.EmpMonGroupSet0.Location = new System.Drawing.Point(23, 313);
            this.EmpMonGroupSet0.Name = "EmpMonGroupSet0";
            this.EmpMonGroupSet0.Size = new System.Drawing.Size(543, 228);
            this.EmpMonGroupSet0.TabIndex = 57;
            this.EmpMonGroupSet0.TabStop = false;
            this.EmpMonGroupSet0.Text = "Set Details";
            this.EmpMonGroupSet0.Visible = false;
            // 
            // EmpMonGroup1
            // 
            this.EmpMonGroup1.Controls.Add(this.label4);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbPhoneall);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbOtherded);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbLateded);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbHouserent);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbMedallowance);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbOtherall);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbConvAllowance);
            this.EmpMonGroup1.Controls.Add(this.EmpMonCkbLatesitall);
            this.EmpMonGroup1.Location = new System.Drawing.Point(14, 47);
            this.EmpMonGroup1.Name = "EmpMonGroup1";
            this.EmpMonGroup1.Size = new System.Drawing.Size(304, 167);
            this.EmpMonGroup1.TabIndex = 0;
            this.EmpMonGroup1.TabStop = false;
            this.EmpMonGroup1.Text = "Allowance/Deduction Setting";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 29);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(292, 13);
            this.label4.TabIndex = 55;
            this.label4.Text = "Following Allowance/Deduction can be set for an Employee.";
            // 
            // EmpMonCkbPhoneall
            // 
            this.EmpMonCkbPhoneall.AutoSize = true;
            this.EmpMonCkbPhoneall.Checked = true;
            this.EmpMonCkbPhoneall.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EmpMonCkbPhoneall.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbPhoneall.Location = new System.Drawing.Point(9, 136);
            this.EmpMonCkbPhoneall.Name = "EmpMonCkbPhoneall";
            this.EmpMonCkbPhoneall.Size = new System.Drawing.Size(124, 17);
            this.EmpMonCkbPhoneall.TabIndex = 50;
            this.EmpMonCkbPhoneall.Text = "Phone Allowance";
            this.EmpMonCkbPhoneall.UseVisualStyleBackColor = true;
            // 
            // EmpMonCkbOtherded
            // 
            this.EmpMonCkbOtherded.AutoSize = true;
            this.EmpMonCkbOtherded.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbOtherded.Location = new System.Drawing.Point(167, 136);
            this.EmpMonCkbOtherded.Name = "EmpMonCkbOtherded";
            this.EmpMonCkbOtherded.Size = new System.Drawing.Size(104, 17);
            this.EmpMonCkbOtherded.TabIndex = 54;
            this.EmpMonCkbOtherded.Text = "Other Deduction";
            this.EmpMonCkbOtherded.UseVisualStyleBackColor = true;
            // 
            // EmpMonCkbLateded
            // 
            this.EmpMonCkbLateded.AutoSize = true;
            this.EmpMonCkbLateded.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbLateded.Location = new System.Drawing.Point(167, 109);
            this.EmpMonCkbLateded.Name = "EmpMonCkbLateded";
            this.EmpMonCkbLateded.Size = new System.Drawing.Size(99, 17);
            this.EmpMonCkbLateded.TabIndex = 53;
            this.EmpMonCkbLateded.Text = "Late Deduction";
            this.EmpMonCkbLateded.UseVisualStyleBackColor = true;
            // 
            // EmpMonCkbHouserent
            // 
            this.EmpMonCkbHouserent.AutoSize = true;
            this.EmpMonCkbHouserent.Checked = true;
            this.EmpMonCkbHouserent.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EmpMonCkbHouserent.Cursor = System.Windows.Forms.Cursors.Default;
            this.EmpMonCkbHouserent.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbHouserent.Location = new System.Drawing.Point(9, 58);
            this.EmpMonCkbHouserent.Name = "EmpMonCkbHouserent";
            this.EmpMonCkbHouserent.Size = new System.Drawing.Size(93, 17);
            this.EmpMonCkbHouserent.TabIndex = 47;
            this.EmpMonCkbHouserent.Text = "House Rent";
            this.EmpMonCkbHouserent.UseVisualStyleBackColor = true;
            this.EmpMonCkbHouserent.CheckStateChanged += new System.EventHandler(this.EmpMsetCkbHouserent_CheckStateChanged);
            // 
            // EmpMonCkbMedallowance
            // 
            this.EmpMonCkbMedallowance.AutoSize = true;
            this.EmpMonCkbMedallowance.Checked = true;
            this.EmpMonCkbMedallowance.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EmpMonCkbMedallowance.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbMedallowance.Location = new System.Drawing.Point(9, 84);
            this.EmpMonCkbMedallowance.Name = "EmpMonCkbMedallowance";
            this.EmpMonCkbMedallowance.Size = new System.Drawing.Size(132, 17);
            this.EmpMonCkbMedallowance.TabIndex = 48;
            this.EmpMonCkbMedallowance.Text = "Medical Allowance";
            this.EmpMonCkbMedallowance.UseVisualStyleBackColor = true;
            // 
            // EmpMonCkbOtherall
            // 
            this.EmpMonCkbOtherall.AutoSize = true;
            this.EmpMonCkbOtherall.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbOtherall.Location = new System.Drawing.Point(167, 84);
            this.EmpMonCkbOtherall.Name = "EmpMonCkbOtherall";
            this.EmpMonCkbOtherall.Size = new System.Drawing.Size(104, 17);
            this.EmpMonCkbOtherall.TabIndex = 52;
            this.EmpMonCkbOtherall.Text = "Other Allowance";
            this.EmpMonCkbOtherall.UseVisualStyleBackColor = true;
            // 
            // EmpMonCkbConvAllowance
            // 
            this.EmpMonCkbConvAllowance.AutoSize = true;
            this.EmpMonCkbConvAllowance.Checked = true;
            this.EmpMonCkbConvAllowance.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EmpMonCkbConvAllowance.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbConvAllowance.Location = new System.Drawing.Point(9, 109);
            this.EmpMonCkbConvAllowance.Name = "EmpMonCkbConvAllowance";
            this.EmpMonCkbConvAllowance.Size = new System.Drawing.Size(158, 17);
            this.EmpMonCkbConvAllowance.TabIndex = 49;
            this.EmpMonCkbConvAllowance.Text = "Conveyance Allowance";
            this.EmpMonCkbConvAllowance.UseVisualStyleBackColor = true;
            this.EmpMonCkbConvAllowance.CheckedChanged += new System.EventHandler(this.EmpMsetCkbConvAllowance_CheckedChanged);
            // 
            // EmpMonCkbLatesitall
            // 
            this.EmpMonCkbLatesitall.AutoSize = true;
            this.EmpMonCkbLatesitall.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonCkbLatesitall.Location = new System.Drawing.Point(167, 58);
            this.EmpMonCkbLatesitall.Name = "EmpMonCkbLatesitall";
            this.EmpMonCkbLatesitall.Size = new System.Drawing.Size(131, 17);
            this.EmpMonCkbLatesitall.TabIndex = 51;
            this.EmpMonCkbLatesitall.Text = "Late Sitting Allowance";
            this.EmpMonCkbLatesitall.UseVisualStyleBackColor = true;
            // 
            // EmpMonTxtName
            // 
            this.EmpMonTxtName.Location = new System.Drawing.Point(55, 21);
            this.EmpMonTxtName.Name = "EmpMonTxtName";
            this.EmpMonTxtName.ReadOnly = true;
            this.EmpMonTxtName.Size = new System.Drawing.Size(145, 20);
            this.EmpMonTxtName.TabIndex = 11;
            this.EmpMonTxtName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtName_KeyPress);
            // 
            // EmpMonLblAllowtotal
            // 
            this.EmpMonLblAllowtotal.AutoSize = true;
            this.EmpMonLblAllowtotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonLblAllowtotal.Location = new System.Drawing.Point(333, 78);
            this.EmpMonLblAllowtotal.Name = "EmpMonLblAllowtotal";
            this.EmpMonLblAllowtotal.Size = new System.Drawing.Size(98, 13);
            this.EmpMonLblAllowtotal.TabIndex = 42;
            this.EmpMonLblAllowtotal.Text = "Allowance Total";
            // 
            // EmpMonLblJobid
            // 
            this.EmpMonLblJobid.AutoSize = true;
            this.EmpMonLblJobid.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonLblJobid.Location = new System.Drawing.Point(307, 25);
            this.EmpMonLblJobid.Name = "EmpMonLblJobid";
            this.EmpMonLblJobid.Size = new System.Drawing.Size(36, 13);
            this.EmpMonLblJobid.TabIndex = 38;
            this.EmpMonLblJobid.Text = "Job Id";
            // 
            // EmpMonTxtDedtotal
            // 
            this.EmpMonTxtDedtotal.Location = new System.Drawing.Point(442, 97);
            this.EmpMonTxtDedtotal.Name = "EmpMonTxtDedtotal";
            this.EmpMonTxtDedtotal.ReadOnly = true;
            this.EmpMonTxtDedtotal.Size = new System.Drawing.Size(81, 20);
            this.EmpMonTxtDedtotal.TabIndex = 43;
            this.EmpMonTxtDedtotal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtDedtotal_KeyPress);
            // 
            // EmpMonTxtJobid
            // 
            this.EmpMonTxtJobid.Location = new System.Drawing.Point(352, 21);
            this.EmpMonTxtJobid.Name = "EmpMonTxtJobid";
            this.EmpMonTxtJobid.ReadOnly = true;
            this.EmpMonTxtJobid.Size = new System.Drawing.Size(171, 20);
            this.EmpMonTxtJobid.TabIndex = 37;
            this.EmpMonTxtJobid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtJobid_KeyPress);
            // 
            // EmpMonTxtAllowtotal
            // 
            this.EmpMonTxtAllowtotal.Location = new System.Drawing.Point(442, 74);
            this.EmpMonTxtAllowtotal.Name = "EmpMonTxtAllowtotal";
            this.EmpMonTxtAllowtotal.ReadOnly = true;
            this.EmpMonTxtAllowtotal.Size = new System.Drawing.Size(81, 20);
            this.EmpMonTxtAllowtotal.TabIndex = 41;
            this.EmpMonTxtAllowtotal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtAllowtotal_KeyPress);
            // 
            // EmpMonBtnApply
            // 
            this.EmpMonBtnApply.Location = new System.Drawing.Point(369, 181);
            this.EmpMonBtnApply.Name = "EmpMonBtnApply";
            this.EmpMonBtnApply.Size = new System.Drawing.Size(74, 23);
            this.EmpMonBtnApply.TabIndex = 3;
            this.EmpMonBtnApply.Text = "Apply";
            this.EmpMonBtnApply.UseVisualStyleBackColor = true;
            this.EmpMonBtnApply.Click += new System.EventHandler(this.EmpMonBtnApply_Click);
            // 
            // EmpMonLblName
            // 
            this.EmpMonLblName.AutoSize = true;
            this.EmpMonLblName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonLblName.Location = new System.Drawing.Point(14, 25);
            this.EmpMonLblName.Name = "EmpMonLblName";
            this.EmpMonLblName.Size = new System.Drawing.Size(35, 13);
            this.EmpMonLblName.TabIndex = 12;
            this.EmpMonLblName.Text = "Name";
            // 
            // EmpMonLblDedtotal
            // 
            this.EmpMonLblDedtotal.AutoSize = true;
            this.EmpMonLblDedtotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonLblDedtotal.Location = new System.Drawing.Point(333, 101);
            this.EmpMonLblDedtotal.Name = "EmpMonLblDedtotal";
            this.EmpMonLblDedtotal.Size = new System.Drawing.Size(98, 13);
            this.EmpMonLblDedtotal.TabIndex = 44;
            this.EmpMonLblDedtotal.Text = "Deduction Total";
            // 
            // EmpMonTxtTotalamount
            // 
            this.EmpMonTxtTotalamount.Location = new System.Drawing.Point(442, 131);
            this.EmpMonTxtTotalamount.Name = "EmpMonTxtTotalamount";
            this.EmpMonTxtTotalamount.ReadOnly = true;
            this.EmpMonTxtTotalamount.Size = new System.Drawing.Size(81, 20);
            this.EmpMonTxtTotalamount.TabIndex = 45;
            this.EmpMonTxtTotalamount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtTotalamount_KeyPress);
            // 
            // EmpMonBtnCancel0
            // 
            this.EmpMonBtnCancel0.Location = new System.Drawing.Point(449, 181);
            this.EmpMonBtnCancel0.Name = "EmpMonBtnCancel0";
            this.EmpMonBtnCancel0.Size = new System.Drawing.Size(74, 23);
            this.EmpMonBtnCancel0.TabIndex = 29;
            this.EmpMonBtnCancel0.Text = "Cancel";
            this.EmpMonBtnCancel0.UseVisualStyleBackColor = true;
            this.EmpMonBtnCancel0.Click += new System.EventHandler(this.EmpMSetBtnCancel1_Click);
            // 
            // EmpMonTxtBps
            // 
            this.EmpMonTxtBps.Location = new System.Drawing.Point(237, 21);
            this.EmpMonTxtBps.Name = "EmpMonTxtBps";
            this.EmpMonTxtBps.ReadOnly = true;
            this.EmpMonTxtBps.Size = new System.Drawing.Size(44, 20);
            this.EmpMonTxtBps.TabIndex = 9;
            this.EmpMonTxtBps.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtBps_KeyPress);
            // 
            // EmpMonLblBps
            // 
            this.EmpMonLblBps.AutoSize = true;
            this.EmpMonLblBps.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonLblBps.Location = new System.Drawing.Point(203, 25);
            this.EmpMonLblBps.Name = "EmpMonLblBps";
            this.EmpMonLblBps.Size = new System.Drawing.Size(28, 13);
            this.EmpMonLblBps.TabIndex = 10;
            this.EmpMonLblBps.Text = "BPS";
            // 
            // EmpMonLblTotalamount
            // 
            this.EmpMonLblTotalamount.AutoSize = true;
            this.EmpMonLblTotalamount.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpMonLblTotalamount.Location = new System.Drawing.Point(333, 134);
            this.EmpMonLblTotalamount.Name = "EmpMonLblTotalamount";
            this.EmpMonLblTotalamount.Size = new System.Drawing.Size(82, 13);
            this.EmpMonLblTotalamount.TabIndex = 46;
            this.EmpMonLblTotalamount.Text = "Total Amount";
            // 
            // EmpMonGridvu
            // 
            this.EmpMonGridvu.AllowUserToAddRows = false;
            this.EmpMonGridvu.AllowUserToDeleteRows = false;
            this.EmpMonGridvu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EmpMonGridvu.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Edit});
            this.EmpMonGridvu.Location = new System.Drawing.Point(23, 35);
            this.EmpMonGridvu.Name = "EmpMonGridvu";
            this.EmpMonGridvu.ReadOnly = true;
            this.EmpMonGridvu.Size = new System.Drawing.Size(883, 270);
            this.EmpMonGridvu.TabIndex = 4;
            this.EmpMonGridvu.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EmpMSetGridvu_CellClick);
            // 
            // Edit
            // 
            this.Edit.HeaderText = "Edit Record";
            this.Edit.Name = "Edit";
            this.Edit.ReadOnly = true;
            this.Edit.Text = "Edit";
            this.Edit.UseColumnTextForButtonValue = true;
            // 
            // EmpMonGroupsearch
            // 
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonTxtFormid);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonTxtRfid);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonCkbDepartment);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonCbmDepartment);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonBtnClear);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonLblSearchoption);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonBtnSearch);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonCkbByrfid);
            this.EmpMonGroupsearch.Controls.Add(this.EmpMonCkbByformid);
            this.EmpMonGroupsearch.Location = new System.Drawing.Point(623, 323);
            this.EmpMonGroupsearch.Name = "EmpMonGroupsearch";
            this.EmpMonGroupsearch.Size = new System.Drawing.Size(283, 173);
            this.EmpMonGroupsearch.TabIndex = 2;
            this.EmpMonGroupsearch.TabStop = false;
            this.EmpMonGroupsearch.Text = "Search Options";
            // 
            // EmpMonTxtFormid
            // 
            this.EmpMonTxtFormid.AcceptsReturn = true;
            this.EmpMonTxtFormid.Location = new System.Drawing.Point(133, 110);
            this.EmpMonTxtFormid.Name = "EmpMonTxtFormid";
            this.EmpMonTxtFormid.Size = new System.Drawing.Size(122, 20);
            this.EmpMonTxtFormid.TabIndex = 31;
            this.EmpMonTxtFormid.Visible = false;
            this.EmpMonTxtFormid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtFormid_KeyPress);
            // 
            // EmpMonTxtRfid
            // 
            this.EmpMonTxtRfid.Location = new System.Drawing.Point(133, 85);
            this.EmpMonTxtRfid.Name = "EmpMonTxtRfid";
            this.EmpMonTxtRfid.Size = new System.Drawing.Size(122, 20);
            this.EmpMonTxtRfid.TabIndex = 30;
            this.EmpMonTxtRfid.Visible = false;
            this.EmpMonTxtRfid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonTxtRfid_KeyPress);
            // 
            // EmpMonCkbDepartment
            // 
            this.EmpMonCkbDepartment.AutoSize = true;
            this.EmpMonCkbDepartment.Checked = true;
            this.EmpMonCkbDepartment.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EmpMonCkbDepartment.Location = new System.Drawing.Point(15, 60);
            this.EmpMonCkbDepartment.Name = "EmpMonCkbDepartment";
            this.EmpMonCkbDepartment.Size = new System.Drawing.Size(81, 17);
            this.EmpMonCkbDepartment.TabIndex = 29;
            this.EmpMonCkbDepartment.Text = "Department";
            this.EmpMonCkbDepartment.UseVisualStyleBackColor = true;
            this.EmpMonCkbDepartment.CheckStateChanged += new System.EventHandler(this.EmpMSetCkbDepartment_CheckStateChanged);
            // 
            // EmpMonCbmDepartment
            // 
            this.EmpMonCbmDepartment.FormattingEnabled = true;
            this.EmpMonCbmDepartment.Location = new System.Drawing.Point(133, 58);
            this.EmpMonCbmDepartment.Name = "EmpMonCbmDepartment";
            this.EmpMonCbmDepartment.Size = new System.Drawing.Size(122, 21);
            this.EmpMonCbmDepartment.TabIndex = 7;
            this.EmpMonCbmDepartment.MouseDown += new System.Windows.Forms.MouseEventHandler(this.EmpMSetCbmDepartment_MouseDown);
            this.EmpMonCbmDepartment.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpMonCbmDepartment_KeyPress);
            // 
            // EmpMonBtnClear
            // 
            this.EmpMonBtnClear.Location = new System.Drawing.Point(194, 138);
            this.EmpMonBtnClear.Name = "EmpMonBtnClear";
            this.EmpMonBtnClear.Size = new System.Drawing.Size(61, 23);
            this.EmpMonBtnClear.TabIndex = 28;
            this.EmpMonBtnClear.Text = "Clear";
            this.EmpMonBtnClear.UseVisualStyleBackColor = true;
            this.EmpMonBtnClear.Click += new System.EventHandler(this.EmpMSetBtnCancel_Click);
            // 
            // EmpMonLblSearchoption
            // 
            this.EmpMonLblSearchoption.AutoSize = true;
            this.EmpMonLblSearchoption.Location = new System.Drawing.Point(12, 27);
            this.EmpMonLblSearchoption.Name = "EmpMonLblSearchoption";
            this.EmpMonLblSearchoption.Size = new System.Drawing.Size(242, 13);
            this.EmpMonLblSearchoption.TabIndex = 25;
            this.EmpMonLblSearchoption.Text = "Following Search option can be used for a record.";
            // 
            // EmpMonBtnSearch
            // 
            this.EmpMonBtnSearch.Location = new System.Drawing.Point(133, 138);
            this.EmpMonBtnSearch.Name = "EmpMonBtnSearch";
            this.EmpMonBtnSearch.Size = new System.Drawing.Size(61, 23);
            this.EmpMonBtnSearch.TabIndex = 26;
            this.EmpMonBtnSearch.Text = "Search Record";
            this.EmpMonBtnSearch.UseVisualStyleBackColor = true;
            this.EmpMonBtnSearch.Click += new System.EventHandler(this.EmpMSetBtnSearch_Click);
            // 
            // EmpMonCkbByrfid
            // 
            this.EmpMonCkbByrfid.AutoSize = true;
            this.EmpMonCkbByrfid.ForeColor = System.Drawing.Color.Black;
            this.EmpMonCkbByrfid.Location = new System.Drawing.Point(15, 87);
            this.EmpMonCkbByrfid.Name = "EmpMonCkbByrfid";
            this.EmpMonCkbByrfid.Size = new System.Drawing.Size(113, 17);
            this.EmpMonCkbByrfid.TabIndex = 1;
            this.EmpMonCkbByrfid.Text = "Search by RF Tag";
            this.EmpMonCkbByrfid.UseVisualStyleBackColor = true;
            this.EmpMonCkbByrfid.CheckStateChanged += new System.EventHandler(this.EmpMSetCkbByrfid_CheckStateChanged);
            // 
            // EmpMonCkbByformid
            // 
            this.EmpMonCkbByformid.AutoSize = true;
            this.EmpMonCkbByformid.Location = new System.Drawing.Point(15, 112);
            this.EmpMonCkbByformid.Name = "EmpMonCkbByformid";
            this.EmpMonCkbByformid.Size = new System.Drawing.Size(112, 17);
            this.EmpMonCkbByformid.TabIndex = 0;
            this.EmpMonCkbByformid.Text = "Search by Form Id";
            this.EmpMonCkbByformid.UseVisualStyleBackColor = true;
            this.EmpMonCkbByformid.CheckStateChanged += new System.EventHandler(this.EmpMSetCkbByformid_CheckStateChanged);
            // 
            // EmpMonGroup2
            // 
            this.EmpMonGroup2.BackColor = System.Drawing.Color.AliceBlue;
            this.EmpMonGroup2.Controls.Add(this.label28);
            this.EmpMonGroup2.Controls.Add(this.label22);
            this.EmpMonGroup2.Controls.Add(this.pictureBox2);
            this.EmpMonGroup2.Controls.Add(this.EmpMonBtnOk);
            this.EmpMonGroup2.Location = new System.Drawing.Point(23, 316);
            this.EmpMonGroup2.Name = "EmpMonGroup2";
            this.EmpMonGroup2.Size = new System.Drawing.Size(549, 148);
            this.EmpMonGroup2.TabIndex = 58;
            this.EmpMonGroup2.TabStop = false;
            this.EmpMonGroup2.Text = "Restricted to Set Details";
            this.EmpMonGroup2.Visible = false;
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label28.ForeColor = System.Drawing.Color.Red;
            this.label28.Location = new System.Drawing.Point(125, 27);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(54, 13);
            this.label28.TabIndex = 60;
            this.label28.Text = "SORRY!";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.ForeColor = System.Drawing.Color.Red;
            this.label22.Location = new System.Drawing.Point(125, 54);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(374, 13);
            this.label22.TabIndex = 32;
            this.label22.Text = "Before 25th of this Month, You can not Set Monthly Details for the Employees.";
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox2.BackgroundImage")));
            this.pictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox2.Location = new System.Drawing.Point(10, 28);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(109, 107);
            this.pictureBox2.TabIndex = 58;
            this.pictureBox2.TabStop = false;
            // 
            // EmpMonBtnOk
            // 
            this.EmpMonBtnOk.Location = new System.Drawing.Point(421, 86);
            this.EmpMonBtnOk.Name = "EmpMonBtnOk";
            this.EmpMonBtnOk.Size = new System.Drawing.Size(74, 23);
            this.EmpMonBtnOk.TabIndex = 29;
            this.EmpMonBtnOk.Text = "Ok";
            this.EmpMonBtnOk.UseVisualStyleBackColor = true;
            this.EmpMonBtnOk.Click += new System.EventHandler(this.EmpMonBtnCancel1_Click);
            // 
            // TabSalary
            // 
            this.TabSalary.BackColor = System.Drawing.Color.AliceBlue;
            this.TabSalary.Controls.Add(this.label32);
            this.TabSalary.Controls.Add(this.EmpSalGroupSalary);
            this.TabSalary.Location = new System.Drawing.Point(4, 22);
            this.TabSalary.Name = "TabSalary";
            this.TabSalary.Size = new System.Drawing.Size(929, 564);
            this.TabSalary.TabIndex = 4;
            this.TabSalary.Text = "Salary";
            // 
            // label32
            // 
            this.label32.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label32.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label32.Location = new System.Drawing.Point(-4, 0);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(940, 24);
            this.label32.TabIndex = 84;
            this.label32.Text = "                                                 MIA MONTHLY SALARY VIEW";
            // 
            // EmpSalGroupSalary
            // 
            this.EmpSalGroupSalary.Controls.Add(this.EmpSalGroupAlloDeductdetail);
            this.EmpSalGroupSalary.Controls.Add(this.EmpSalGroupSalarydetail);
            this.EmpSalGroupSalary.Controls.Add(this.EmpSalGroupdaysdetail);
            this.EmpSalGroupSalary.Controls.Add(this.EmpSalPicbox);
            this.EmpSalGroupSalary.Controls.Add(this.EmpSalTxtName);
            this.EmpSalGroupSalary.Controls.Add(this.groupBox2);
            this.EmpSalGroupSalary.Controls.Add(this.EmpSalGridvu);
            this.EmpSalGroupSalary.Location = new System.Drawing.Point(18, 43);
            this.EmpSalGroupSalary.Name = "EmpSalGroupSalary";
            this.EmpSalGroupSalary.Size = new System.Drawing.Size(891, 503);
            this.EmpSalGroupSalary.TabIndex = 3;
            this.EmpSalGroupSalary.TabStop = false;
            this.EmpSalGroupSalary.Text = "Salary";
            // 
            // EmpSalGroupAlloDeductdetail
            // 
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblHouseallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblOtherdeduc);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblAbsentdeduc);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblLatededuc);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtOtherdeduc);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtAdvadeduc);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtLatededuc);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblOtherallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblLatesittingallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblPhoneallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblConvallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalLblMedicalallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtOtherallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtConveyanceallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtLatesittingallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtMedicalallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtPhoneallow);
            this.EmpSalGroupAlloDeductdetail.Controls.Add(this.EmpSalTxtHouseallow);
            this.EmpSalGroupAlloDeductdetail.Location = new System.Drawing.Point(13, 396);
            this.EmpSalGroupAlloDeductdetail.Name = "EmpSalGroupAlloDeductdetail";
            this.EmpSalGroupAlloDeductdetail.Size = new System.Drawing.Size(572, 97);
            this.EmpSalGroupAlloDeductdetail.TabIndex = 74;
            this.EmpSalGroupAlloDeductdetail.TabStop = false;
            this.EmpSalGroupAlloDeductdetail.Text = "Allowance/Deduction Details";
            this.EmpSalGroupAlloDeductdetail.Visible = false;
            // 
            // EmpSalLblHouseallow
            // 
            this.EmpSalLblHouseallow.AutoSize = true;
            this.EmpSalLblHouseallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblHouseallow.Location = new System.Drawing.Point(9, 30);
            this.EmpSalLblHouseallow.Name = "EmpSalLblHouseallow";
            this.EmpSalLblHouseallow.Size = new System.Drawing.Size(105, 13);
            this.EmpSalLblHouseallow.TabIndex = 74;
            this.EmpSalLblHouseallow.Text = "House Allowance";
            // 
            // EmpSalLblOtherdeduc
            // 
            this.EmpSalLblOtherdeduc.AutoSize = true;
            this.EmpSalLblOtherdeduc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblOtherdeduc.Location = new System.Drawing.Point(366, 37);
            this.EmpSalLblOtherdeduc.Name = "EmpSalLblOtherdeduc";
            this.EmpSalLblOtherdeduc.Size = new System.Drawing.Size(100, 13);
            this.EmpSalLblOtherdeduc.TabIndex = 90;
            this.EmpSalLblOtherdeduc.Text = "Other Deduction";
            // 
            // EmpSalLblAbsentdeduc
            // 
            this.EmpSalLblAbsentdeduc.AutoSize = true;
            this.EmpSalLblAbsentdeduc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblAbsentdeduc.Location = new System.Drawing.Point(176, 37);
            this.EmpSalLblAbsentdeduc.Name = "EmpSalLblAbsentdeduc";
            this.EmpSalLblAbsentdeduc.Size = new System.Drawing.Size(108, 13);
            this.EmpSalLblAbsentdeduc.TabIndex = 89;
            this.EmpSalLblAbsentdeduc.Text = "Absent Deduction";
            // 
            // EmpSalLblLatededuc
            // 
            this.EmpSalLblLatededuc.AutoSize = true;
            this.EmpSalLblLatededuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblLatededuc.Location = new System.Drawing.Point(9, 37);
            this.EmpSalLblLatededuc.Name = "EmpSalLblLatededuc";
            this.EmpSalLblLatededuc.Size = new System.Drawing.Size(94, 13);
            this.EmpSalLblLatededuc.TabIndex = 85;
            this.EmpSalLblLatededuc.Text = "Late Deduction";
            // 
            // EmpSalTxtOtherdeduc
            // 
            this.EmpSalTxtOtherdeduc.Location = new System.Drawing.Point(509, 33);
            this.EmpSalTxtOtherdeduc.Name = "EmpSalTxtOtherdeduc";
            this.EmpSalTxtOtherdeduc.ReadOnly = true;
            this.EmpSalTxtOtherdeduc.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtOtherdeduc.TabIndex = 88;
            this.EmpSalTxtOtherdeduc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtOtherdeduc_KeyPress);
            // 
            // EmpSalTxtAdvadeduc
            // 
            this.EmpSalTxtAdvadeduc.Location = new System.Drawing.Point(307, 33);
            this.EmpSalTxtAdvadeduc.Name = "EmpSalTxtAdvadeduc";
            this.EmpSalTxtAdvadeduc.ReadOnly = true;
            this.EmpSalTxtAdvadeduc.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtAdvadeduc.TabIndex = 87;
            this.EmpSalTxtAdvadeduc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtAdvadeduc_KeyPress);
            // 
            // EmpSalTxtLatededuc
            // 
            this.EmpSalTxtLatededuc.Location = new System.Drawing.Point(115, 33);
            this.EmpSalTxtLatededuc.Name = "EmpSalTxtLatededuc";
            this.EmpSalTxtLatededuc.ReadOnly = true;
            this.EmpSalTxtLatededuc.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtLatededuc.TabIndex = 86;
            this.EmpSalTxtLatededuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtLatededuc_KeyPress);
            // 
            // EmpSalLblOtherallow
            // 
            this.EmpSalLblOtherallow.AutoSize = true;
            this.EmpSalLblOtherallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblOtherallow.Location = new System.Drawing.Point(366, 64);
            this.EmpSalLblOtherallow.Name = "EmpSalLblOtherallow";
            this.EmpSalLblOtherallow.Size = new System.Drawing.Size(100, 13);
            this.EmpSalLblOtherallow.TabIndex = 84;
            this.EmpSalLblOtherallow.Text = "Other Allowance";
            // 
            // EmpSalLblLatesittingallow
            // 
            this.EmpSalLblLatesittingallow.AutoSize = true;
            this.EmpSalLblLatesittingallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblLatesittingallow.Location = new System.Drawing.Point(176, 64);
            this.EmpSalLblLatesittingallow.Name = "EmpSalLblLatesittingallow";
            this.EmpSalLblLatesittingallow.Size = new System.Drawing.Size(128, 13);
            this.EmpSalLblLatesittingallow.TabIndex = 83;
            this.EmpSalLblLatesittingallow.Text = "Latesitting Allowance";
            // 
            // EmpSalLblPhoneallow
            // 
            this.EmpSalLblPhoneallow.AutoSize = true;
            this.EmpSalLblPhoneallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblPhoneallow.Location = new System.Drawing.Point(9, 64);
            this.EmpSalLblPhoneallow.Name = "EmpSalLblPhoneallow";
            this.EmpSalLblPhoneallow.Size = new System.Drawing.Size(105, 13);
            this.EmpSalLblPhoneallow.TabIndex = 82;
            this.EmpSalLblPhoneallow.Text = "Phone Allowance";
            // 
            // EmpSalLblConvallow
            // 
            this.EmpSalLblConvallow.AutoSize = true;
            this.EmpSalLblConvallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblConvallow.Location = new System.Drawing.Point(366, 30);
            this.EmpSalLblConvallow.Name = "EmpSalLblConvallow";
            this.EmpSalLblConvallow.Size = new System.Drawing.Size(143, 13);
            this.EmpSalLblConvallow.TabIndex = 81;
            this.EmpSalLblConvallow.Text = "Conveyance Allowance ";
            // 
            // EmpSalLblMedicalallow
            // 
            this.EmpSalLblMedicalallow.AutoSize = true;
            this.EmpSalLblMedicalallow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblMedicalallow.Location = new System.Drawing.Point(176, 30);
            this.EmpSalLblMedicalallow.Name = "EmpSalLblMedicalallow";
            this.EmpSalLblMedicalallow.Size = new System.Drawing.Size(113, 13);
            this.EmpSalLblMedicalallow.TabIndex = 80;
            this.EmpSalLblMedicalallow.Text = "Medical Allowance";
            // 
            // EmpSalTxtOtherallow
            // 
            this.EmpSalTxtOtherallow.Location = new System.Drawing.Point(509, 60);
            this.EmpSalTxtOtherallow.Name = "EmpSalTxtOtherallow";
            this.EmpSalTxtOtherallow.ReadOnly = true;
            this.EmpSalTxtOtherallow.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtOtherallow.TabIndex = 79;
            this.EmpSalTxtOtherallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtOtherallow_KeyPress);
            // 
            // EmpSalTxtConveyanceallow
            // 
            this.EmpSalTxtConveyanceallow.Location = new System.Drawing.Point(509, 26);
            this.EmpSalTxtConveyanceallow.Name = "EmpSalTxtConveyanceallow";
            this.EmpSalTxtConveyanceallow.ReadOnly = true;
            this.EmpSalTxtConveyanceallow.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtConveyanceallow.TabIndex = 78;
            this.EmpSalTxtConveyanceallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtConveyanceallow_KeyPress);
            // 
            // EmpSalTxtLatesittingallow
            // 
            this.EmpSalTxtLatesittingallow.Location = new System.Drawing.Point(307, 60);
            this.EmpSalTxtLatesittingallow.Name = "EmpSalTxtLatesittingallow";
            this.EmpSalTxtLatesittingallow.ReadOnly = true;
            this.EmpSalTxtLatesittingallow.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtLatesittingallow.TabIndex = 77;
            this.EmpSalTxtLatesittingallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtLatesittingallow_KeyPress);
            // 
            // EmpSalTxtMedicalallow
            // 
            this.EmpSalTxtMedicalallow.Location = new System.Drawing.Point(307, 26);
            this.EmpSalTxtMedicalallow.Name = "EmpSalTxtMedicalallow";
            this.EmpSalTxtMedicalallow.ReadOnly = true;
            this.EmpSalTxtMedicalallow.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtMedicalallow.TabIndex = 76;
            this.EmpSalTxtMedicalallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtMedicalallow_KeyPress);
            // 
            // EmpSalTxtPhoneallow
            // 
            this.EmpSalTxtPhoneallow.Location = new System.Drawing.Point(115, 60);
            this.EmpSalTxtPhoneallow.Name = "EmpSalTxtPhoneallow";
            this.EmpSalTxtPhoneallow.ReadOnly = true;
            this.EmpSalTxtPhoneallow.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtPhoneallow.TabIndex = 75;
            this.EmpSalTxtPhoneallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtPhoneallow_KeyPress);
            // 
            // EmpSalTxtHouseallow
            // 
            this.EmpSalTxtHouseallow.Location = new System.Drawing.Point(115, 26);
            this.EmpSalTxtHouseallow.Name = "EmpSalTxtHouseallow";
            this.EmpSalTxtHouseallow.ReadOnly = true;
            this.EmpSalTxtHouseallow.Size = new System.Drawing.Size(55, 20);
            this.EmpSalTxtHouseallow.TabIndex = 74;
            this.EmpSalTxtHouseallow.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtHouseallow_KeyPress);
            // 
            // EmpSalGroupSalarydetail
            // 
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalCkbAllowamount);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalLblBasicpay);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalTxtBasicpay);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalTxtAllowamount);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalLblNetsalary);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalTxtDeducamount);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalCkbDeducamount);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalLblFollow0);
            this.EmpSalGroupSalarydetail.Controls.Add(this.EmpSalTxtNetsalary);
            this.EmpSalGroupSalarydetail.Location = new System.Drawing.Point(13, 206);
            this.EmpSalGroupSalarydetail.Name = "EmpSalGroupSalarydetail";
            this.EmpSalGroupSalarydetail.Size = new System.Drawing.Size(247, 183);
            this.EmpSalGroupSalarydetail.TabIndex = 70;
            this.EmpSalGroupSalarydetail.TabStop = false;
            this.EmpSalGroupSalarydetail.Text = "Allowance/Deduction Detail";
            this.EmpSalGroupSalarydetail.Visible = false;
            // 
            // EmpSalCkbAllowamount
            // 
            this.EmpSalCkbAllowamount.AutoSize = true;
            this.EmpSalCkbAllowamount.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalCkbAllowamount.Location = new System.Drawing.Point(18, 58);
            this.EmpSalCkbAllowamount.Name = "EmpSalCkbAllowamount";
            this.EmpSalCkbAllowamount.Size = new System.Drawing.Size(136, 17);
            this.EmpSalCkbAllowamount.TabIndex = 56;
            this.EmpSalCkbAllowamount.Text = "Allowances Amount";
            this.EmpSalCkbAllowamount.UseVisualStyleBackColor = true;
            this.EmpSalCkbAllowamount.CheckStateChanged += new System.EventHandler(this.EmpSalCkbAllowamount_CheckStateChanged);
            // 
            // EmpSalLblBasicpay
            // 
            this.EmpSalLblBasicpay.AutoSize = true;
            this.EmpSalLblBasicpay.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblBasicpay.Location = new System.Drawing.Point(18, 119);
            this.EmpSalLblBasicpay.Name = "EmpSalLblBasicpay";
            this.EmpSalLblBasicpay.Size = new System.Drawing.Size(70, 13);
            this.EmpSalLblBasicpay.TabIndex = 61;
            this.EmpSalLblBasicpay.Text = "Basiec Pay";
            // 
            // EmpSalTxtBasicpay
            // 
            this.EmpSalTxtBasicpay.Location = new System.Drawing.Point(142, 116);
            this.EmpSalTxtBasicpay.Name = "EmpSalTxtBasicpay";
            this.EmpSalTxtBasicpay.ReadOnly = true;
            this.EmpSalTxtBasicpay.Size = new System.Drawing.Size(80, 20);
            this.EmpSalTxtBasicpay.TabIndex = 56;
            this.EmpSalTxtBasicpay.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtBasicpay_KeyPress);
            // 
            // EmpSalTxtAllowamount
            // 
            this.EmpSalTxtAllowamount.Location = new System.Drawing.Point(165, 56);
            this.EmpSalTxtAllowamount.Name = "EmpSalTxtAllowamount";
            this.EmpSalTxtAllowamount.ReadOnly = true;
            this.EmpSalTxtAllowamount.Size = new System.Drawing.Size(57, 20);
            this.EmpSalTxtAllowamount.TabIndex = 65;
            this.EmpSalTxtAllowamount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtAllowamount_KeyPress);
            // 
            // EmpSalLblNetsalary
            // 
            this.EmpSalLblNetsalary.AutoSize = true;
            this.EmpSalLblNetsalary.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblNetsalary.Location = new System.Drawing.Point(19, 154);
            this.EmpSalLblNetsalary.Name = "EmpSalLblNetsalary";
            this.EmpSalLblNetsalary.Size = new System.Drawing.Size(66, 13);
            this.EmpSalLblNetsalary.TabIndex = 72;
            this.EmpSalLblNetsalary.Text = "Net Salary";
            // 
            // EmpSalTxtDeducamount
            // 
            this.EmpSalTxtDeducamount.Location = new System.Drawing.Point(165, 83);
            this.EmpSalTxtDeducamount.Name = "EmpSalTxtDeducamount";
            this.EmpSalTxtDeducamount.ReadOnly = true;
            this.EmpSalTxtDeducamount.Size = new System.Drawing.Size(57, 20);
            this.EmpSalTxtDeducamount.TabIndex = 63;
            this.EmpSalTxtDeducamount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtDeducamount_KeyPress);
            // 
            // EmpSalCkbDeducamount
            // 
            this.EmpSalCkbDeducamount.AutoSize = true;
            this.EmpSalCkbDeducamount.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalCkbDeducamount.Location = new System.Drawing.Point(18, 85);
            this.EmpSalCkbDeducamount.Name = "EmpSalCkbDeducamount";
            this.EmpSalCkbDeducamount.Size = new System.Drawing.Size(136, 17);
            this.EmpSalCkbDeducamount.TabIndex = 73;
            this.EmpSalCkbDeducamount.Text = "Deductions Amount";
            this.EmpSalCkbDeducamount.UseVisualStyleBackColor = true;
            this.EmpSalCkbDeducamount.CheckStateChanged += new System.EventHandler(this.EmpSalCkbDeducamount_CheckStateChanged);
            // 
            // EmpSalLblFollow0
            // 
            this.EmpSalLblFollow0.AutoSize = true;
            this.EmpSalLblFollow0.Location = new System.Drawing.Point(14, 25);
            this.EmpSalLblFollow0.Name = "EmpSalLblFollow0";
            this.EmpSalLblFollow0.Size = new System.Drawing.Size(124, 13);
            this.EmpSalLblFollow0.TabIndex = 55;
            this.EmpSalLblFollow0.Text = "Following is Salary detail.";
            // 
            // EmpSalTxtNetsalary
            // 
            this.EmpSalTxtNetsalary.Location = new System.Drawing.Point(101, 150);
            this.EmpSalTxtNetsalary.Name = "EmpSalTxtNetsalary";
            this.EmpSalTxtNetsalary.ReadOnly = true;
            this.EmpSalTxtNetsalary.Size = new System.Drawing.Size(121, 20);
            this.EmpSalTxtNetsalary.TabIndex = 67;
            this.EmpSalTxtNetsalary.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtNetsalary_KeyPress);
            // 
            // EmpSalGroupdaysdetail
            // 
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalTxtTotaldays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalLblTotaldays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalLblFollow1);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalLblPresentdays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalTxtPresentdays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalTxtAbsentdays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalLblAbsentdays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalLblLeavedays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalTxtLeavedays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalTxtHolidays);
            this.EmpSalGroupdaysdetail.Controls.Add(this.EmpSalLblHolidays);
            this.EmpSalGroupdaysdetail.Location = new System.Drawing.Point(275, 206);
            this.EmpSalGroupdaysdetail.Name = "EmpSalGroupdaysdetail";
            this.EmpSalGroupdaysdetail.Size = new System.Drawing.Size(310, 183);
            this.EmpSalGroupdaysdetail.TabIndex = 47;
            this.EmpSalGroupdaysdetail.TabStop = false;
            this.EmpSalGroupdaysdetail.Text = "Attendance Detail";
            this.EmpSalGroupdaysdetail.Visible = false;
            // 
            // EmpSalTxtTotaldays
            // 
            this.EmpSalTxtTotaldays.Location = new System.Drawing.Point(108, 148);
            this.EmpSalTxtTotaldays.Name = "EmpSalTxtTotaldays";
            this.EmpSalTxtTotaldays.ReadOnly = true;
            this.EmpSalTxtTotaldays.Size = new System.Drawing.Size(46, 20);
            this.EmpSalTxtTotaldays.TabIndex = 67;
            this.EmpSalTxtTotaldays.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtTotaldays_KeyPress);
            // 
            // EmpSalLblTotaldays
            // 
            this.EmpSalLblTotaldays.AutoSize = true;
            this.EmpSalLblTotaldays.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblTotaldays.Location = new System.Drawing.Point(23, 152);
            this.EmpSalLblTotaldays.Name = "EmpSalLblTotaldays";
            this.EmpSalLblTotaldays.Size = new System.Drawing.Size(66, 13);
            this.EmpSalLblTotaldays.TabIndex = 68;
            this.EmpSalLblTotaldays.Text = "Total days";
            // 
            // EmpSalLblFollow1
            // 
            this.EmpSalLblFollow1.AutoSize = true;
            this.EmpSalLblFollow1.Location = new System.Drawing.Point(18, 25);
            this.EmpSalLblFollow1.Name = "EmpSalLblFollow1";
            this.EmpSalLblFollow1.Size = new System.Drawing.Size(190, 13);
            this.EmpSalLblFollow1.TabIndex = 55;
            this.EmpSalLblFollow1.Text = "Following is Monthly Attendance detail.";
            // 
            // EmpSalLblPresentdays
            // 
            this.EmpSalLblPresentdays.AutoSize = true;
            this.EmpSalLblPresentdays.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblPresentdays.Location = new System.Drawing.Point(23, 67);
            this.EmpSalLblPresentdays.Name = "EmpSalLblPresentdays";
            this.EmpSalLblPresentdays.Size = new System.Drawing.Size(80, 13);
            this.EmpSalLblPresentdays.TabIndex = 64;
            this.EmpSalLblPresentdays.Text = "Present days";
            // 
            // EmpSalTxtPresentdays
            // 
            this.EmpSalTxtPresentdays.Location = new System.Drawing.Point(108, 63);
            this.EmpSalTxtPresentdays.Name = "EmpSalTxtPresentdays";
            this.EmpSalTxtPresentdays.ReadOnly = true;
            this.EmpSalTxtPresentdays.Size = new System.Drawing.Size(46, 20);
            this.EmpSalTxtPresentdays.TabIndex = 63;
            this.EmpSalTxtPresentdays.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtPresentdays_KeyPress);
            // 
            // EmpSalTxtAbsentdays
            // 
            this.EmpSalTxtAbsentdays.Location = new System.Drawing.Point(108, 98);
            this.EmpSalTxtAbsentdays.Name = "EmpSalTxtAbsentdays";
            this.EmpSalTxtAbsentdays.ReadOnly = true;
            this.EmpSalTxtAbsentdays.Size = new System.Drawing.Size(46, 20);
            this.EmpSalTxtAbsentdays.TabIndex = 65;
            this.EmpSalTxtAbsentdays.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtAbsentdays_KeyPress);
            // 
            // EmpSalLblAbsentdays
            // 
            this.EmpSalLblAbsentdays.AutoSize = true;
            this.EmpSalLblAbsentdays.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblAbsentdays.Location = new System.Drawing.Point(23, 102);
            this.EmpSalLblAbsentdays.Name = "EmpSalLblAbsentdays";
            this.EmpSalLblAbsentdays.Size = new System.Drawing.Size(76, 13);
            this.EmpSalLblAbsentdays.TabIndex = 66;
            this.EmpSalLblAbsentdays.Text = "Absent days";
            // 
            // EmpSalLblLeavedays
            // 
            this.EmpSalLblLeavedays.AutoSize = true;
            this.EmpSalLblLeavedays.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblLeavedays.Location = new System.Drawing.Point(166, 67);
            this.EmpSalLblLeavedays.Name = "EmpSalLblLeavedays";
            this.EmpSalLblLeavedays.Size = new System.Drawing.Size(72, 13);
            this.EmpSalLblLeavedays.TabIndex = 57;
            this.EmpSalLblLeavedays.Text = "Leave days";
            // 
            // EmpSalTxtLeavedays
            // 
            this.EmpSalTxtLeavedays.Location = new System.Drawing.Point(246, 63);
            this.EmpSalTxtLeavedays.Name = "EmpSalTxtLeavedays";
            this.EmpSalTxtLeavedays.ReadOnly = true;
            this.EmpSalTxtLeavedays.Size = new System.Drawing.Size(46, 20);
            this.EmpSalTxtLeavedays.TabIndex = 56;
            this.EmpSalTxtLeavedays.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtLeavedays_KeyPress);
            // 
            // EmpSalTxtHolidays
            // 
            this.EmpSalTxtHolidays.Location = new System.Drawing.Point(246, 98);
            this.EmpSalTxtHolidays.Name = "EmpSalTxtHolidays";
            this.EmpSalTxtHolidays.ReadOnly = true;
            this.EmpSalTxtHolidays.Size = new System.Drawing.Size(46, 20);
            this.EmpSalTxtHolidays.TabIndex = 58;
            this.EmpSalTxtHolidays.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtHolidays_KeyPress);
            // 
            // EmpSalLblHolidays
            // 
            this.EmpSalLblHolidays.AutoSize = true;
            this.EmpSalLblHolidays.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmpSalLblHolidays.Location = new System.Drawing.Point(166, 102);
            this.EmpSalLblHolidays.Name = "EmpSalLblHolidays";
            this.EmpSalLblHolidays.Size = new System.Drawing.Size(55, 13);
            this.EmpSalLblHolidays.TabIndex = 59;
            this.EmpSalLblHolidays.Text = "Holidays";
            // 
            // EmpSalPicbox
            // 
            this.EmpSalPicbox.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.EmpSalPicbox.Location = new System.Drawing.Point(754, 17);
            this.EmpSalPicbox.Name = "EmpSalPicbox";
            this.EmpSalPicbox.Size = new System.Drawing.Size(124, 151);
            this.EmpSalPicbox.TabIndex = 32;
            this.EmpSalPicbox.TabStop = false;
            // 
            // EmpSalTxtName
            // 
            this.EmpSalTxtName.Location = new System.Drawing.Point(754, 172);
            this.EmpSalTxtName.Name = "EmpSalTxtName";
            this.EmpSalTxtName.ReadOnly = true;
            this.EmpSalTxtName.Size = new System.Drawing.Size(124, 20);
            this.EmpSalTxtName.TabIndex = 51;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.EmpSalDatetime);
            this.groupBox2.Controls.Add(this.EmpSalCkbRftag);
            this.groupBox2.Controls.Add(this.EmpSalCkbFormid);
            this.groupBox2.Controls.Add(this.EmpSalTxtFormid);
            this.groupBox2.Controls.Add(this.EmpSalBtnSearch);
            this.groupBox2.Controls.Add(this.EmpSalTxtRftag);
            this.groupBox2.Controls.Add(this.EmpSalLblFollow2);
            this.groupBox2.Controls.Add(this.EmpSalCkbDept);
            this.groupBox2.Controls.Add(this.EmpSalBtnCancel0);
            this.groupBox2.Controls.Add(this.EmpSalCombDept);
            this.groupBox2.Location = new System.Drawing.Point(604, 206);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(281, 243);
            this.groupBox2.TabIndex = 33;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Search Options";
            // 
            // EmpSalDatetime
            // 
            this.EmpSalDatetime.Location = new System.Drawing.Point(130, 62);
            this.EmpSalDatetime.Name = "EmpSalDatetime";
            this.EmpSalDatetime.Size = new System.Drawing.Size(127, 20);
            this.EmpSalDatetime.TabIndex = 35;
            // 
            // EmpSalCkbRftag
            // 
            this.EmpSalCkbRftag.AutoSize = true;
            this.EmpSalCkbRftag.ForeColor = System.Drawing.Color.Black;
            this.EmpSalCkbRftag.Location = new System.Drawing.Point(16, 141);
            this.EmpSalCkbRftag.Name = "EmpSalCkbRftag";
            this.EmpSalCkbRftag.Size = new System.Drawing.Size(113, 17);
            this.EmpSalCkbRftag.TabIndex = 1;
            this.EmpSalCkbRftag.Text = "Search by RF Tag";
            this.EmpSalCkbRftag.UseVisualStyleBackColor = true;
            this.EmpSalCkbRftag.CheckStateChanged += new System.EventHandler(this.EmpSalCkbRftag_CheckStateChanged);
            // 
            // EmpSalCkbFormid
            // 
            this.EmpSalCkbFormid.AutoSize = true;
            this.EmpSalCkbFormid.Location = new System.Drawing.Point(16, 170);
            this.EmpSalCkbFormid.Name = "EmpSalCkbFormid";
            this.EmpSalCkbFormid.Size = new System.Drawing.Size(112, 17);
            this.EmpSalCkbFormid.TabIndex = 0;
            this.EmpSalCkbFormid.Text = "Search by Form Id";
            this.EmpSalCkbFormid.UseVisualStyleBackColor = true;
            this.EmpSalCkbFormid.CheckStateChanged += new System.EventHandler(this.EmpSalCkbFormid_CheckStateChanged);
            // 
            // EmpSalTxtFormid
            // 
            this.EmpSalTxtFormid.AcceptsReturn = true;
            this.EmpSalTxtFormid.Enabled = false;
            this.EmpSalTxtFormid.Location = new System.Drawing.Point(130, 168);
            this.EmpSalTxtFormid.Name = "EmpSalTxtFormid";
            this.EmpSalTxtFormid.Size = new System.Drawing.Size(126, 20);
            this.EmpSalTxtFormid.TabIndex = 31;
            this.EmpSalTxtFormid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtFormid_KeyPress);
            // 
            // EmpSalBtnSearch
            // 
            this.EmpSalBtnSearch.Location = new System.Drawing.Point(128, 201);
            this.EmpSalBtnSearch.Name = "EmpSalBtnSearch";
            this.EmpSalBtnSearch.Size = new System.Drawing.Size(61, 23);
            this.EmpSalBtnSearch.TabIndex = 26;
            this.EmpSalBtnSearch.Text = "Search Record";
            this.EmpSalBtnSearch.UseVisualStyleBackColor = true;
            this.EmpSalBtnSearch.Click += new System.EventHandler(this.EmpSalBtnSearch_Click);
            // 
            // EmpSalTxtRftag
            // 
            this.EmpSalTxtRftag.Enabled = false;
            this.EmpSalTxtRftag.Location = new System.Drawing.Point(130, 139);
            this.EmpSalTxtRftag.Name = "EmpSalTxtRftag";
            this.EmpSalTxtRftag.Size = new System.Drawing.Size(127, 20);
            this.EmpSalTxtRftag.TabIndex = 30;
            this.EmpSalTxtRftag.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpSalTxtRftag_KeyPress);
            // 
            // EmpSalLblFollow2
            // 
            this.EmpSalLblFollow2.AutoSize = true;
            this.EmpSalLblFollow2.Location = new System.Drawing.Point(13, 66);
            this.EmpSalLblFollow2.Name = "EmpSalLblFollow2";
            this.EmpSalLblFollow2.Size = new System.Drawing.Size(73, 13);
            this.EmpSalLblFollow2.TabIndex = 25;
            this.EmpSalLblFollow2.Text = "Select Month:";
            // 
            // EmpSalCkbDept
            // 
            this.EmpSalCkbDept.AutoSize = true;
            this.EmpSalCkbDept.Checked = true;
            this.EmpSalCkbDept.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EmpSalCkbDept.Location = new System.Drawing.Point(16, 109);
            this.EmpSalCkbDept.Name = "EmpSalCkbDept";
            this.EmpSalCkbDept.Size = new System.Drawing.Size(81, 17);
            this.EmpSalCkbDept.TabIndex = 29;
            this.EmpSalCkbDept.Text = "Department";
            this.EmpSalCkbDept.UseVisualStyleBackColor = true;
            this.EmpSalCkbDept.CheckStateChanged += new System.EventHandler(this.EmpSalCkbDept_CheckStateChanged);
            // 
            // EmpSalBtnCancel0
            // 
            this.EmpSalBtnCancel0.Location = new System.Drawing.Point(196, 201);
            this.EmpSalBtnCancel0.Name = "EmpSalBtnCancel0";
            this.EmpSalBtnCancel0.Size = new System.Drawing.Size(61, 23);
            this.EmpSalBtnCancel0.TabIndex = 28;
            this.EmpSalBtnCancel0.Text = "Cancel";
            this.EmpSalBtnCancel0.UseVisualStyleBackColor = true;
            this.EmpSalBtnCancel0.Click += new System.EventHandler(this.EmpSalBtnCancel0_Click);
            // 
            // EmpSalCombDept
            // 
            this.EmpSalCombDept.FormattingEnabled = true;
            this.EmpSalCombDept.Location = new System.Drawing.Point(130, 107);
            this.EmpSalCombDept.Name = "EmpSalCombDept";
            this.EmpSalCombDept.Size = new System.Drawing.Size(127, 21);
            this.EmpSalCombDept.TabIndex = 7;
            this.EmpSalCombDept.MouseDown += new System.Windows.Forms.MouseEventHandler(this.EmpSalCombDept_MouseDown);
            // 
            // EmpSalGridvu
            // 
            this.EmpSalGridvu.AllowUserToAddRows = false;
            this.EmpSalGridvu.AllowUserToDeleteRows = false;
            this.EmpSalGridvu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EmpSalGridvu.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ViewSalaryDetail});
            this.EmpSalGridvu.Location = new System.Drawing.Point(13, 17);
            this.EmpSalGridvu.Name = "EmpSalGridvu";
            this.EmpSalGridvu.ReadOnly = true;
            this.EmpSalGridvu.Size = new System.Drawing.Size(718, 151);
            this.EmpSalGridvu.TabIndex = 4;
            this.EmpSalGridvu.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EmpSalGridvu_CellClick);
            this.EmpSalGridvu.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.EmpSalGridvu_CellEnter);
            // 
            // ViewSalaryDetail
            // 
            this.ViewSalaryDetail.HeaderText = "View Detail";
            this.ViewSalaryDetail.Name = "ViewSalaryDetail";
            this.ViewSalaryDetail.ReadOnly = true;
            this.ViewSalaryDetail.Text = "Detail";
            this.ViewSalaryDetail.UseColumnTextForButtonValue = true;
            // 
            // TabAttendance
            // 
            this.TabAttendance.BackColor = System.Drawing.Color.AliceBlue;
            this.TabAttendance.Controls.Add(this.label31);
            this.TabAttendance.Controls.Add(this.EmpAttGroup0);
            this.TabAttendance.Controls.Add(this.EmpAttGroup1);
            this.TabAttendance.Location = new System.Drawing.Point(4, 22);
            this.TabAttendance.Name = "TabAttendance";
            this.TabAttendance.Size = new System.Drawing.Size(929, 564);
            this.TabAttendance.TabIndex = 5;
            this.TabAttendance.Text = "Attendance";
            // 
            // label31
            // 
            this.label31.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label31.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label31.Location = new System.Drawing.Point(-4, 0);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(940, 24);
            this.label31.TabIndex = 83;
            this.label31.Text = "                                                 MIA MONTHLY ATTENDANCE VIEW";
            // 
            // EmpAttGroup0
            // 
            this.EmpAttGroup0.Controls.Add(this.EmpAttBtnClear);
            this.EmpAttGroup0.Controls.Add(this.EmpAttTxtSearchbyformid);
            this.EmpAttGroup0.Controls.Add(this.EmpAttTxtSearchbyrfid);
            this.EmpAttGroup0.Controls.Add(this.EmpAttCkbSearchbyrfid);
            this.EmpAttGroup0.Controls.Add(this.EmpAttCkbSearchbyformid);
            this.EmpAttGroup0.Controls.Add(this.EmpAttPicBox);
            this.EmpAttGroup0.Controls.Add(this.EmpAttGridvu);
            this.EmpAttGroup0.Controls.Add(this.EmAttBtnSearch);
            this.EmpAttGroup0.Controls.Add(this.EmpAttCtlDatetime);
            this.EmpAttGroup0.Location = new System.Drawing.Point(15, 32);
            this.EmpAttGroup0.Name = "EmpAttGroup0";
            this.EmpAttGroup0.Size = new System.Drawing.Size(899, 338);
            this.EmpAttGroup0.TabIndex = 5;
            this.EmpAttGroup0.TabStop = false;
            this.EmpAttGroup0.Text = "Monthly Attendance";
            // 
            // EmpAttBtnClear
            // 
            this.EmpAttBtnClear.Location = new System.Drawing.Point(814, 287);
            this.EmpAttBtnClear.Name = "EmpAttBtnClear";
            this.EmpAttBtnClear.Size = new System.Drawing.Size(74, 23);
            this.EmpAttBtnClear.TabIndex = 36;
            this.EmpAttBtnClear.Text = "Clear";
            this.EmpAttBtnClear.UseVisualStyleBackColor = true;
            this.EmpAttBtnClear.Click += new System.EventHandler(this.EmpAttBtnClear_Click);
            // 
            // EmpAttTxtSearchbyformid
            // 
            this.EmpAttTxtSearchbyformid.AcceptsReturn = true;
            this.EmpAttTxtSearchbyformid.Location = new System.Drawing.Point(825, 254);
            this.EmpAttTxtSearchbyformid.Name = "EmpAttTxtSearchbyformid";
            this.EmpAttTxtSearchbyformid.Size = new System.Drawing.Size(63, 20);
            this.EmpAttTxtSearchbyformid.TabIndex = 35;
            this.EmpAttTxtSearchbyformid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpAttTxtSearchbyformid_KeyPress);
            // 
            // EmpAttTxtSearchbyrfid
            // 
            this.EmpAttTxtSearchbyrfid.Enabled = false;
            this.EmpAttTxtSearchbyrfid.Location = new System.Drawing.Point(825, 227);
            this.EmpAttTxtSearchbyrfid.Name = "EmpAttTxtSearchbyrfid";
            this.EmpAttTxtSearchbyrfid.Size = new System.Drawing.Size(63, 20);
            this.EmpAttTxtSearchbyrfid.TabIndex = 34;
            this.EmpAttTxtSearchbyrfid.Visible = false;
            this.EmpAttTxtSearchbyrfid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EmpAttTxtSearchbyrfid_KeyPress);
            // 
            // EmpAttCkbSearchbyrfid
            // 
            this.EmpAttCkbSearchbyrfid.AutoSize = true;
            this.EmpAttCkbSearchbyrfid.Enabled = false;
            this.EmpAttCkbSearchbyrfid.ForeColor = System.Drawing.Color.Black;
            this.EmpAttCkbSearchbyrfid.Location = new System.Drawing.Point(712, 229);
            this.EmpAttCkbSearchbyrfid.Name = "EmpAttCkbSearchbyrfid";
            this.EmpAttCkbSearchbyrfid.Size = new System.Drawing.Size(113, 17);
            this.EmpAttCkbSearchbyrfid.TabIndex = 33;
            this.EmpAttCkbSearchbyrfid.Text = "Search by RF Tag";
            this.EmpAttCkbSearchbyrfid.UseVisualStyleBackColor = true;
            this.EmpAttCkbSearchbyrfid.Visible = false;
            this.EmpAttCkbSearchbyrfid.CheckStateChanged += new System.EventHandler(this.EmpAttCkbSearchbyrfid_CheckStateChanged);
            // 
            // EmpAttCkbSearchbyformid
            // 
            this.EmpAttCkbSearchbyformid.AutoSize = true;
            this.EmpAttCkbSearchbyformid.Location = new System.Drawing.Point(712, 256);
            this.EmpAttCkbSearchbyformid.Name = "EmpAttCkbSearchbyformid";
            this.EmpAttCkbSearchbyformid.Size = new System.Drawing.Size(112, 17);
            this.EmpAttCkbSearchbyformid.TabIndex = 32;
            this.EmpAttCkbSearchbyformid.Text = "Search by Form Id";
            this.EmpAttCkbSearchbyformid.UseVisualStyleBackColor = true;
            this.EmpAttCkbSearchbyformid.CheckStateChanged += new System.EventHandler(this.EmpAttLblSearchbyformid_CheckStateChanged);
            // 
            // EmpAttPicBox
            // 
            this.EmpAttPicBox.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.EmpAttPicBox.Location = new System.Drawing.Point(770, 32);
            this.EmpAttPicBox.Name = "EmpAttPicBox";
            this.EmpAttPicBox.Size = new System.Drawing.Size(118, 151);
            this.EmpAttPicBox.TabIndex = 3;
            this.EmpAttPicBox.TabStop = false;
            // 
            // EmpAttGridvu
            // 
            this.EmpAttGridvu.AllowUserToAddRows = false;
            this.EmpAttGridvu.AllowUserToDeleteRows = false;
            this.EmpAttGridvu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EmpAttGridvu.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CheckDetaill});
            this.EmpAttGridvu.Location = new System.Drawing.Point(10, 32);
            this.EmpAttGridvu.Name = "EmpAttGridvu";
            this.EmpAttGridvu.ReadOnly = true;
            this.EmpAttGridvu.Size = new System.Drawing.Size(696, 293);
            this.EmpAttGridvu.TabIndex = 0;
            this.EmpAttGridvu.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EmpAttGridvu_CellClick);
            // 
            // CheckDetaill
            // 
            this.CheckDetaill.HeaderText = "Detail View";
            this.CheckDetaill.Name = "CheckDetaill";
            this.CheckDetaill.ReadOnly = true;
            this.CheckDetaill.Text = "Detail";
            this.CheckDetaill.UseColumnTextForButtonValue = true;
            // 
            // EmAttBtnSearch
            // 
            this.EmAttBtnSearch.Location = new System.Drawing.Point(734, 288);
            this.EmAttBtnSearch.Name = "EmAttBtnSearch";
            this.EmAttBtnSearch.Size = new System.Drawing.Size(74, 23);
            this.EmAttBtnSearch.TabIndex = 1;
            this.EmAttBtnSearch.Text = "Search";
            this.EmAttBtnSearch.UseVisualStyleBackColor = true;
            this.EmAttBtnSearch.Click += new System.EventHandler(this.EmAttBtnSearch_Click_1);
            // 
            // EmpAttCtlDatetime
            // 
            this.EmpAttCtlDatetime.Location = new System.Drawing.Point(735, 201);
            this.EmpAttCtlDatetime.Name = "EmpAttCtlDatetime";
            this.EmpAttCtlDatetime.Size = new System.Drawing.Size(153, 20);
            this.EmpAttCtlDatetime.TabIndex = 2;
            // 
            // EmpAttGroup1
            // 
            this.EmpAttGroup1.Controls.Add(this.EmpAttGridvudetail);
            this.EmpAttGroup1.Location = new System.Drawing.Point(39, 376);
            this.EmpAttGroup1.Name = "EmpAttGroup1";
            this.EmpAttGroup1.Size = new System.Drawing.Size(736, 185);
            this.EmpAttGroup1.TabIndex = 4;
            this.EmpAttGroup1.TabStop = false;
            this.EmpAttGroup1.Text = "Detailed View";
            this.EmpAttGroup1.Visible = false;
            // 
            // EmpAttGridvudetail
            // 
            this.EmpAttGridvudetail.AllowUserToAddRows = false;
            this.EmpAttGridvudetail.AllowUserToDeleteRows = false;
            this.EmpAttGridvudetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EmpAttGridvudetail.Location = new System.Drawing.Point(10, 19);
            this.EmpAttGridvudetail.Name = "EmpAttGridvudetail";
            this.EmpAttGridvudetail.ReadOnly = true;
            this.EmpAttGridvudetail.Size = new System.Drawing.Size(715, 160);
            this.EmpAttGridvudetail.TabIndex = 3;
            // 
            // TabScanner
            // 
            this.TabScanner.Location = new System.Drawing.Point(4, 22);
            this.TabScanner.Name = "TabScanner";
            this.TabScanner.Size = new System.Drawing.Size(994, 630);
            this.TabScanner.TabIndex = 3;
            this.TabScanner.Text = "Scanner";
            this.TabScanner.UseVisualStyleBackColor = true;
            // 
            // TabDEO
            // 
            this.TabDEO.Location = new System.Drawing.Point(4, 22);
            this.TabDEO.Name = "TabDEO";
            this.TabDEO.Size = new System.Drawing.Size(994, 630);
            this.TabDEO.TabIndex = 4;
            this.TabDEO.Text = "Data Entry";
            this.TabDEO.UseVisualStyleBackColor = true;
            // 
            // TabHr
            // 
            this.TabHr.Location = new System.Drawing.Point(4, 22);
            this.TabHr.Name = "TabHr";
            this.TabHr.Size = new System.Drawing.Size(994, 630);
            this.TabHr.TabIndex = 5;
            this.TabHr.Text = "Hr";
            this.TabHr.UseVisualStyleBackColor = true;
            // 
            // TabReception
            // 
            this.TabReception.Location = new System.Drawing.Point(4, 22);
            this.TabReception.Name = "TabReception";
            this.TabReception.Size = new System.Drawing.Size(994, 630);
            this.TabReception.TabIndex = 6;
            this.TabReception.Text = "Reception";
            this.TabReception.UseVisualStyleBackColor = true;
            // 
            // TabReportViewer
            // 
            this.TabReportViewer.BackColor = System.Drawing.SystemColors.Desktop;
            this.TabReportViewer.Controls.Add(this.tabControl1);
            this.TabReportViewer.Location = new System.Drawing.Point(4, 22);
            this.TabReportViewer.Name = "TabReportViewer";
            this.TabReportViewer.Padding = new System.Windows.Forms.Padding(3);
            this.TabReportViewer.Size = new System.Drawing.Size(994, 630);
            this.TabReportViewer.TabIndex = 7;
            this.TabReportViewer.Text = "Report Viewer";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.TabReport_Employee);
            this.tabControl1.Location = new System.Drawing.Point(29, 17);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(949, 583);
            this.tabControl1.TabIndex = 0;
            // 
            // TabReport_Employee
            // 
            this.TabReport_Employee.Controls.Add(this.label37);
            this.TabReport_Employee.Controls.Add(this.CrysReport_Employee);
            this.TabReport_Employee.Location = new System.Drawing.Point(4, 22);
            this.TabReport_Employee.Name = "TabReport_Employee";
            this.TabReport_Employee.Padding = new System.Windows.Forms.Padding(3);
            this.TabReport_Employee.Size = new System.Drawing.Size(941, 557);
            this.TabReport_Employee.TabIndex = 0;
            this.TabReport_Employee.Text = "Employee";
            this.TabReport_Employee.UseVisualStyleBackColor = true;
            // 
            // label37
            // 
            this.label37.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label37.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label37.Location = new System.Drawing.Point(0, 3);
            this.label37.Name = "label37";
            this.label37.Size = new System.Drawing.Size(945, 24);
            this.label37.TabIndex = 87;
            this.label37.Text = "                                                 MIA EMPLOYEE REPORTS VIEWR";
            // 
            // CrysReport_Employee
            // 
            this.CrysReport_Employee.ActiveViewIndex = 0;
            this.CrysReport_Employee.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.CrysReport_Employee.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.CrysReport_Employee.DisplayGroupTree = false;
            this.CrysReport_Employee.DisplayStatusBar = false;
            this.CrysReport_Employee.DisplayToolbar = false;
            this.CrysReport_Employee.Location = new System.Drawing.Point(10, 43);
            this.CrysReport_Employee.Name = "CrysReport_Employee";
            this.CrysReport_Employee.ReportSource = this.CryReport_Employee1;
            this.CrysReport_Employee.Size = new System.Drawing.Size(921, 505);
            this.CrysReport_Employee.TabIndex = 1;
            // 
            // ChoosRow
            // 
            this.ChoosRow.HeaderText = "Choose Row";
            this.ChoosRow.Name = "ChoosRow";
            this.ChoosRow.ReadOnly = true;
            // 
            // LblNullogin
            // 
            this.LblNullogin.AutoSize = true;
            this.LblNullogin.BackColor = System.Drawing.SystemColors.Control;
            this.LblNullogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblNullogin.ForeColor = System.Drawing.SystemColors.Desktop;
            this.LblNullogin.Location = new System.Drawing.Point(258, 693);
            this.LblNullogin.Name = "LblNullogin";
            this.LblNullogin.Size = new System.Drawing.Size(54, 13);
            this.LblNullogin.TabIndex = 2;
            this.LblNullogin.Text = "Null Login";
            // 
            // LblLoginstatus
            // 
            this.LblLoginstatus.AllowDrop = true;
            this.LblLoginstatus.AutoSize = true;
            this.LblLoginstatus.BackColor = System.Drawing.SystemColors.Control;
            this.LblLoginstatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblLoginstatus.Location = new System.Drawing.Point(173, 693);
            this.LblLoginstatus.Name = "LblLoginstatus";
            this.LblLoginstatus.Size = new System.Drawing.Size(69, 13);
            this.LblLoginstatus.TabIndex = 3;
            this.LblLoginstatus.Text = "Login Status:";
            // 
            // LblLoginame
            // 
            this.LblLoginame.AllowDrop = true;
            this.LblLoginame.AutoSize = true;
            this.LblLoginame.BackColor = System.Drawing.SystemColors.Control;
            this.LblLoginame.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblLoginame.Location = new System.Drawing.Point(15, 693);
            this.LblLoginame.Name = "LblLoginame";
            this.LblLoginame.Size = new System.Drawing.Size(67, 13);
            this.LblLoginame.TabIndex = 4;
            this.LblLoginame.Text = "Login Name:";
            // 
            // LblNullname
            // 
            this.LblNullname.AutoSize = true;
            this.LblNullname.BackColor = System.Drawing.SystemColors.Control;
            this.LblNullname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblNullname.ForeColor = System.Drawing.SystemColors.Desktop;
            this.LblNullname.Location = new System.Drawing.Point(99, 693);
            this.LblNullname.Name = "LblNullname";
            this.LblNullname.Size = new System.Drawing.Size(56, 13);
            this.LblNullname.TabIndex = 5;
            this.LblNullname.Text = "Null Name";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(309, 107);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Date of Birth";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(35, 112);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(32, 13);
            this.label6.TabIndex = 9;
            this.label6.Text = "CNIC";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(309, 71);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 13);
            this.label7.TabIndex = 8;
            this.label7.Text = "Father Name";
            // 
            // textBox15
            // 
            this.textBox15.Location = new System.Drawing.Point(146, 105);
            this.textBox15.Name = "textBox15";
            this.textBox15.Size = new System.Drawing.Size(148, 20);
            this.textBox15.TabIndex = 6;
            // 
            // textBox16
            // 
            this.textBox16.Location = new System.Drawing.Point(397, 103);
            this.textBox16.Name = "textBox16";
            this.textBox16.Size = new System.Drawing.Size(148, 20);
            this.textBox16.TabIndex = 5;
            // 
            // textBox17
            // 
            this.textBox17.Location = new System.Drawing.Point(397, 67);
            this.textBox17.Name = "textBox17";
            this.textBox17.Size = new System.Drawing.Size(148, 20);
            this.textBox17.TabIndex = 4;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(708, 51);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(40, 13);
            this.label8.TabIndex = 24;
            this.label8.Text = "Picture";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(754, 231);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(151, 21);
            this.button1.TabIndex = 2;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackColor = System.Drawing.Color.DimGray;
            this.pictureBox3.Location = new System.Drawing.Point(754, 51);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(148, 174);
            this.pictureBox3.TabIndex = 1;
            this.pictureBox3.TabStop = false;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(35, 77);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(84, 13);
            this.label9.TabIndex = 7;
            this.label9.Text = "Employee Name";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.checkBox1);
            this.groupBox3.Controls.Add(this.button2);
            this.groupBox3.Controls.Add(this.button3);
            this.groupBox3.Controls.Add(this.checkBox2);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.checkBox3);
            this.groupBox3.Controls.Add(this.textBox18);
            this.groupBox3.Location = new System.Drawing.Point(470, 342);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(432, 233);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Search Options";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(34, 44);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(242, 13);
            this.label10.TabIndex = 25;
            this.label10.Text = "Following Search option can be used for a record.";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(38, 135);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(102, 17);
            this.checkBox1.TabIndex = 2;
            this.checkBox1.Text = "Search by CNIC";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(323, 206);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(74, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Edit";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(248, 206);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(74, 23);
            this.button3.TabIndex = 26;
            this.button3.Text = "Search Record";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(38, 108);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(113, 17);
            this.checkBox2.TabIndex = 1;
            this.checkBox2.Text = "Search by RF Tag";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(163, 157);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(79, 13);
            this.label11.TabIndex = 25;
            this.label11.Text = "Search Record";
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(38, 81);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(112, 17);
            this.checkBox3.TabIndex = 0;
            this.checkBox3.Text = "Search by Form Id";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // textBox18
            // 
            this.textBox18.Location = new System.Drawing.Point(248, 153);
            this.textBox18.Name = "textBox18";
            this.textBox18.Size = new System.Drawing.Size(152, 20);
            this.textBox18.TabIndex = 23;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(36, 252);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(98, 13);
            this.label12.TabIndex = 24;
            this.label12.Text = "Temporary Address";
            // 
            // textBox19
            // 
            this.textBox19.Location = new System.Drawing.Point(146, 248);
            this.textBox19.Name = "textBox19";
            this.textBox19.Size = new System.Drawing.Size(450, 20);
            this.textBox19.TabIndex = 23;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(36, 288);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(99, 13);
            this.label13.TabIndex = 22;
            this.label13.Text = "Permanent Address";
            // 
            // textBox20
            // 
            this.textBox20.Location = new System.Drawing.Point(146, 284);
            this.textBox20.Name = "textBox20";
            this.textBox20.Size = new System.Drawing.Size(449, 20);
            this.textBox20.TabIndex = 21;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(307, 202);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(55, 13);
            this.label14.TabIndex = 18;
            this.label14.Text = "Phone No";
            // 
            // textBox21
            // 
            this.textBox21.Location = new System.Drawing.Point(397, 198);
            this.textBox21.Name = "textBox21";
            this.textBox21.Size = new System.Drawing.Size(148, 20);
            this.textBox21.TabIndex = 17;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(35, 205);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(73, 13);
            this.label15.TabIndex = 16;
            this.label15.Text = "Email Address";
            // 
            // textBox22
            // 
            this.textBox22.Location = new System.Drawing.Point(146, 198);
            this.textBox22.Name = "textBox22";
            this.textBox22.Size = new System.Drawing.Size(148, 20);
            this.textBox22.TabIndex = 15;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(307, 168);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(47, 13);
            this.label16.TabIndex = 14;
            this.label16.Text = "Domicile";
            // 
            // textBox23
            // 
            this.textBox23.Location = new System.Drawing.Point(397, 164);
            this.textBox23.Name = "textBox23";
            this.textBox23.Size = new System.Drawing.Size(148, 20);
            this.textBox23.TabIndex = 13;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(35, 171);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(45, 13);
            this.label17.TabIndex = 12;
            this.label17.Text = "Religion";
            // 
            // textBox24
            // 
            this.textBox24.Location = new System.Drawing.Point(146, 164);
            this.textBox24.Name = "textBox24";
            this.textBox24.Size = new System.Drawing.Size(148, 20);
            this.textBox24.TabIndex = 11;
            // 
            // textBox25
            // 
            this.textBox25.Location = new System.Drawing.Point(146, 70);
            this.textBox25.Name = "textBox25";
            this.textBox25.Size = new System.Drawing.Size(148, 20);
            this.textBox25.TabIndex = 3;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.groupBox5);
            this.groupBox4.Controls.Add(this.label19);
            this.groupBox4.Controls.Add(this.textBox26);
            this.groupBox4.Controls.Add(this.label20);
            this.groupBox4.Location = new System.Drawing.Point(38, 339);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(402, 236);
            this.groupBox4.TabIndex = 0;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Employee Account";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label18);
            this.groupBox5.Controls.Add(this.radioButton3);
            this.groupBox5.Controls.Add(this.radioButton4);
            this.groupBox5.Location = new System.Drawing.Point(27, 117);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(307, 120);
            this.groupBox5.TabIndex = 31;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Account Status";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(38, 40);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(234, 13);
            this.label18.TabIndex = 28;
            this.label18.Text = "Temporarily an Account can be Enable/Disable.";
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(38, 94);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(55, 17);
            this.radioButton3.TabIndex = 30;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "Diable";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Location = new System.Drawing.Point(38, 69);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(58, 17);
            this.radioButton4.TabIndex = 29;
            this.radioButton4.TabStop = true;
            this.radioButton4.Text = "Enable";
            this.radioButton4.UseVisualStyleBackColor = true;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(27, 51);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(210, 13);
            this.label19.TabIndex = 27;
            this.label19.Text = "This login name will be useful only on Web.";
            // 
            // textBox26
            // 
            this.textBox26.Location = new System.Drawing.Point(122, 76);
            this.textBox26.Name = "textBox26";
            this.textBox26.Size = new System.Drawing.Size(148, 20);
            this.textBox26.TabIndex = 19;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(27, 80);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(64, 13);
            this.label20.TabIndex = 20;
            this.label20.Text = "Login Name";
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
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.toolStripSeparator,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.employeeToolStripMenuItem,
            this.managementToolStripMenuItem});
            this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image")));
            this.openToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem.Size = new System.Drawing.Size(151, 22);
            this.openToolStripMenuItem.Text = "&Open";
            // 
            // employeeToolStripMenuItem
            // 
            this.employeeToolStripMenuItem.Name = "employeeToolStripMenuItem";
            this.employeeToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
            this.employeeToolStripMenuItem.Text = "Employee";
            // 
            // managementToolStripMenuItem
            // 
            this.managementToolStripMenuItem.Name = "managementToolStripMenuItem";
            this.managementToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
            this.managementToolStripMenuItem.Text = "Management";
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(148, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(151, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.scanningProcessToolStripMenuItem,
            this.receptonToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.toolsToolStripMenuItem.Text = "&Tools";
            // 
            // scanningProcessToolStripMenuItem
            // 
            this.scanningProcessToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.scanningToolStripMenuItem,
            this.dataEntryToolStripMenuItem,
            this.hRVerifyingToolStripMenuItem});
            this.scanningProcessToolStripMenuItem.Name = "scanningProcessToolStripMenuItem";
            this.scanningProcessToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.scanningProcessToolStripMenuItem.Text = "Data Entry Process";
            // 
            // scanningToolStripMenuItem
            // 
            this.scanningToolStripMenuItem.Name = "scanningToolStripMenuItem";
            this.scanningToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.scanningToolStripMenuItem.Text = "Scanning";
            this.scanningToolStripMenuItem.Click += new System.EventHandler(this.scanningToolStripMenuItem_Click);
            // 
            // dataEntryToolStripMenuItem
            // 
            this.dataEntryToolStripMenuItem.Name = "dataEntryToolStripMenuItem";
            this.dataEntryToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.dataEntryToolStripMenuItem.Text = "Data Entry";
            this.dataEntryToolStripMenuItem.Click += new System.EventHandler(this.dataEntryToolStripMenuItem_Click);
            // 
            // hRVerifyingToolStripMenuItem
            // 
            this.hRVerifyingToolStripMenuItem.Name = "hRVerifyingToolStripMenuItem";
            this.hRVerifyingToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.hRVerifyingToolStripMenuItem.Text = "HR Verifying";
            this.hRVerifyingToolStripMenuItem.Click += new System.EventHandler(this.hRVerifyingToolStripMenuItem_Click);
            // 
            // receptonToolStripMenuItem
            // 
            this.receptonToolStripMenuItem.Name = "receptonToolStripMenuItem";
            this.receptonToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.receptonToolStripMenuItem.Text = "Recepton";
            this.receptonToolStripMenuItem.Click += new System.EventHandler(this.receptonToolStripMenuItem_Click);
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
            this.contentsToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.contentsToolStripMenuItem.Text = "&Contents";
            // 
            // indexToolStripMenuItem
            // 
            this.indexToolStripMenuItem.Name = "indexToolStripMenuItem";
            this.indexToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.indexToolStripMenuItem.Text = "&Index";
            // 
            // searchToolStripMenuItem
            // 
            this.searchToolStripMenuItem.Name = "searchToolStripMenuItem";
            this.searchToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.searchToolStripMenuItem.Text = "&Search";
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(126, 6);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.aboutToolStripMenuItem.Text = "&About...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Black;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Location = new System.Drawing.Point(0, 24);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1028, 2);
            this.pictureBox1.TabIndex = 53;
            this.pictureBox1.TabStop = false;
            // 
            // Admin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(1028, 746);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.LblNullogin);
            this.Controls.Add(this.TabAdmin);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.LblNullname);
            this.Controls.Add(this.LblLoginstatus);
            this.Controls.Add(this.LblLoginame);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Admin";
            this.Text = "M.I.A Administration";
            this.Activated += new System.EventHandler(this.Admin_Activated);
            this.Load += new System.EventHandler(this.Admin_Load);
            this.TabAdmin.ResumeLayout(false);
            this.TabMgm.ResumeLayout(false);
            this.Management.ResumeLayout(false);
            this.TabDepartments.ResumeLayout(false);
            this.MgmDeptGroupDetailvu.ResumeLayout(false);
            this.MgmDeptGroupDetailvu.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmDeptGridvu)).EndInit();
            this.MgmDeptGroupoption.ResumeLayout(false);
            this.MgmDeptGroupoption.PerformLayout();
            this.MgmDeptGroupadddeptscale.ResumeLayout(false);
            this.MgmDeptGroupadddeptscale.PerformLayout();
            this.TabJob.ResumeLayout(false);
            this.MgmJobGroupSelectedemp.ResumeLayout(false);
            this.MgmJobGroupSelectedemp.PerformLayout();
            this.MgmJobGroupeEmpdetail.ResumeLayout(false);
            this.MgmJobGroupeEmpdetail.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmJobPicbox)).EndInit();
            this.MgmJobGroupRestricted.ResumeLayout(false);
            this.MgmJobGroupRestricted.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            this.MgmJobGroupRestrictedpermote.ResumeLayout(false);
            this.MgmJobGroupRestrictedpermote.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox6)).EndInit();
            this.MgmJobPanelPermotion.ResumeLayout(false);
            this.MgmJobPanelPermotion.PerformLayout();
            this.MgmJobGroupSearch.ResumeLayout(false);
            this.MgmJobGroupSearch.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmJobGridvu)).EndInit();
            this.TabAttend.ResumeLayout(false);
            this.MgmAttGroupSearch.ResumeLayout(false);
            this.MgmAttGroupSearch.PerformLayout();
            this.MgmAttGroupDetailview.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.MgmAttGridAttdetailvu)).EndInit();
            this.MgmAttGroupEmpatt.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.MgmAttGridAttendVu)).EndInit();
            this.TabMgmGeneral.ResumeLayout(false);
            this.MgmGenGroup.ResumeLayout(false);
            this.MgmGenGroup.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.MgmGenGroupLeave.ResumeLayout(false);
            this.MgmGenGroupLeave.PerformLayout();
            this.MgmGenGroupBoxStatus.ResumeLayout(false);
            this.MgmGenGroupBoxStatus.PerformLayout();
            this.MgmGenGroupBoxLeave.ResumeLayout(false);
            this.MgmGenGroupBoxLeave.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.MgmGenGroupboxSetholiday.ResumeLayout(false);
            this.MgmGenGroupboxSetholiday.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MgmGenPicture)).EndInit();
            this.TabEmp.ResumeLayout(false);
            this.Employees.ResumeLayout(false);
            this.TabGeneral.ResumeLayout(false);
            this.TabGeneral.PerformLayout();
            this.EmpGenGroupSearch.ResumeLayout(false);
            this.EmpGenGroupSearch.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpGenPicbox)).EndInit();
            this.EmpGenGroupEmpaccount.ResumeLayout(false);
            this.EmpGenGroupEmpaccount.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.TabMonthlySettings.ResumeLayout(false);
            this.EmpMonGroupSet0.ResumeLayout(false);
            this.EmpMonGroupSet0.PerformLayout();
            this.EmpMonGroup1.ResumeLayout(false);
            this.EmpMonGroup1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpMonGridvu)).EndInit();
            this.EmpMonGroupsearch.ResumeLayout(false);
            this.EmpMonGroupsearch.PerformLayout();
            this.EmpMonGroup2.ResumeLayout(false);
            this.EmpMonGroup2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.TabSalary.ResumeLayout(false);
            this.EmpSalGroupSalary.ResumeLayout(false);
            this.EmpSalGroupSalary.PerformLayout();
            this.EmpSalGroupAlloDeductdetail.ResumeLayout(false);
            this.EmpSalGroupAlloDeductdetail.PerformLayout();
            this.EmpSalGroupSalarydetail.ResumeLayout(false);
            this.EmpSalGroupSalarydetail.PerformLayout();
            this.EmpSalGroupdaysdetail.ResumeLayout(false);
            this.EmpSalGroupdaysdetail.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpSalPicbox)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpSalGridvu)).EndInit();
            this.TabAttendance.ResumeLayout(false);
            this.EmpAttGroup0.ResumeLayout(false);
            this.EmpAttGroup0.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EmpAttPicBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.EmpAttGridvu)).EndInit();
            this.EmpAttGroup1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.EmpAttGridvudetail)).EndInit();
            this.TabReportViewer.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.TabReport_Employee.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnLogout;
        private System.Windows.Forms.TabControl TabAdmin;
        private System.Windows.Forms.TabPage TabMgm;
        private System.Windows.Forms.TabPage TabEmp;
        private Label LblNullogin;
        private Label LblLoginstatus;
        private Label LblLoginame;
        private Label LblNullname;        
        private DataGridViewCheckBoxColumn ChoosRow;
        private Label label5;
        private Label label6;
        private Label label7;
        private TextBox textBox15;
        private TextBox textBox16;
        private TextBox textBox17;
        private Label label8;
        private Button button1;
        private PictureBox pictureBox3;
        private Label label9;
        private GroupBox groupBox3;
        private Label label10;
        private CheckBox checkBox1;
        private Button button2;
        private Button button3;
        private CheckBox checkBox2;
        private Label label11;
        private CheckBox checkBox3;
        private TextBox textBox18;
        private Label label12;
        private TextBox textBox19;
        private Label label13;
        private TextBox textBox20;
        private Label label14;
        private TextBox textBox21;
        private Label label15;
        private TextBox textBox22;
        private Label label16;
        private TextBox textBox23;
        private Label label17;
        private TextBox textBox24;
        private TextBox textBox25;
        private GroupBox groupBox4;
        private GroupBox groupBox5;
        private Label label18;
        private RadioButton radioButton3;
        private RadioButton radioButton4;
        private Label label19;
        private TextBox textBox26;
        private Label label20;
        private TabControl Employees;
        private TabPage TabGeneral;
        private GroupBox EmpGenGroupSearch;
        private Label EmpGenLblSearchoption;
        private CheckBox EmpGenCkbCnic;
        private Button EmpGenBtnSave;
        private Button EmpGenBtnSearch;
        private CheckBox EmpGenCkbRfid;
        private Label EmpGenLblSearchrec;
        private CheckBox EmpGenCkbFormid;
        private TextBox EmpGenTxtSearchrec;
        private Label EmpGenLblPic;
        private Label EmpGenLblTempadd;
        private TextBox EmpGenTxtTempadd;
        private Label EmpGenLblPermadd;
        private TextBox EmpGenTxtPermadd;
        private Label EmpGenLblPhno;
        private TextBox EmpGenTxtPhno;
        private Label EmpGenLblEmail;
        private TextBox EmpGenTxtEmail;
        private Label EmpGenLblDomicile;
        private TextBox EmpGenTxtDomicile;
        private Label EmpGenLblReligion;
        private TextBox EmpGenTxtReligion;
        private Label EmpGenLblDob;
        private Label EmpGenLblCnic;
        private Label EmpGenLblFname;
        private Label EmpGenLblEname;
        private TextBox EmpGenTxtCnic;
        private TextBox EmpGenTxtDob;
        private TextBox EmpGenTxtFname;
        private TextBox EmpGentxtEname;
        private Button EmpGenBtnBrowse;
        private PictureBox EmpGenPicbox;
        private GroupBox EmpGenGroupEmpaccount;
        private GroupBox groupBox1;
        private Label EmpGenLblAccstatus;
        private RadioButton EmpGenRdbDisable;
        private RadioButton EmpGenRdbEnable;
        private Label label2;
        private TextBox EmpGenTxtLogname;
        private Label EmpGenLblLogname;        
        private TabPage TabMonthlySettings;
        private TextBox EmpGenTxtPwd;
        private Label EmpGenLblPwd;
        private Label EmpGenLblBnkbranch;
        private TextBox EmpGenTxtBnkbranch;
        private Label EmpGenLblBnkacc;
        private TextBox EmpGenTxtBnkacc;
        private Label EmpGenLblHiredate;
        private TextBox EmpGenTxtHiredate;
        private Label EmpGenLblRfid;
        private TextBox EmpGenTxtRfid;
        private Button button4;
        private Button EmpGenBtnEdit;
        private Button EmpGenBtnCancel;
        private GroupBox EmpMonGroupsearch;
        private Button EmpMonBtnClear;
        private Label EmpMonLblSearchoption;
        private Button EmpMonBtnSearch;
        private CheckBox EmpMonCkbByrfid;
        private CheckBox EmpMonCkbByformid;
        private TextBox EmpGenTxtUpdateby;
        private Label EmpGenLblUpdateby;        
        private DataGridView EmpMonGridvu;
        private ComboBox EmpMonCbmDepartment;
        private CheckBox EmpMonCkbDepartment;
        private TabPage TabSalary;
        private TabPage TabAttendance;
        private TextBox EmpMonTxtFormid;
        private TextBox EmpMonTxtRfid;
        private Button EmpMonBtnCancel0;
        private Button EmpMonBtnApply;
        private Label EmpMonLblBps;
        private TextBox EmpMonTxtBps;
        private Label EmpMonLblName;
        private Label EmpMonLblJobid;
        private TextBox EmpMonTxtJobid;
        private Label EmpMonLblTotalamount;
        private TextBox EmpMonTxtTotalamount;
        private Label EmpMonLblDedtotal;
        private TextBox EmpMonTxtDedtotal;
        private Label EmpMonLblAllowtotal;
        private TextBox EmpMonTxtAllowtotal;
        private DataGridViewButtonColumn Edit;
        private CheckBox EmpMonCkbHouserent;
        private TextBox EmpMonTxtName;
        private CheckBox EmpMonCkbOtherded;
        private CheckBox EmpMonCkbLateded;
        private CheckBox EmpMonCkbOtherall;
        private CheckBox EmpMonCkbLatesitall;
        private CheckBox EmpMonCkbPhoneall;
        private CheckBox EmpMonCkbConvAllowance;
        private CheckBox EmpMonCkbMedallowance;
        private GroupBox EmpMonGroup1;
        private GroupBox EmpMonGroupSet0;
        private Label label4;
        private TabPage TabScanner;
        private TabPage TabDEO;
        private TabPage TabHr;
        private TabPage TabReception;
        private Button EmAttBtnSearch;
        private DataGridView EmpAttGridvu;
        private DateTimePicker EmpAttCtlDatetime;
        private DataGridView EmpAttGridvudetail;
        private GroupBox EmpAttGroup1;
        private GroupBox EmpAttGroup0;
        private PictureBox EmpAttPicBox;
        private TextBox EmpAttTxtSearchbyformid;
        private TextBox EmpAttTxtSearchbyrfid;
        private CheckBox EmpAttCkbSearchbyrfid;
        private CheckBox EmpAttCkbSearchbyformid;
        private Button EmpAttBtnClear;
        private GroupBox EmpMonGroup2;
        private Button EmpMonBtnOk;
        private Label label22;
        
        private Label label28;
        private PictureBox pictureBox2;
        private GroupBox EmpSalGroupSalary;
        private TextBox EmpSalTxtFormid;
        private TextBox EmpSalTxtRftag;
        private CheckBox EmpSalCkbDept;
        private ComboBox EmpSalCombDept;
        private Button EmpSalBtnCancel0;
        private Label EmpSalLblFollow2;
        private Button EmpSalBtnSearch;
        private CheckBox EmpSalCkbRftag;
        private CheckBox EmpSalCkbFormid;
        private GroupBox groupBox2;
        private PictureBox EmpSalPicbox;
        private DataGridView EmpSalGridvu;
        private DateTimePicker EmpSalDatetime;
        private GroupBox EmpSalGroupdaysdetail;
        private Label EmpSalLblFollow1;
        private TextBox EmpSalTxtName;
        private Label EmpSalLblLeavedays;
        private TextBox EmpSalTxtHolidays;
        private TextBox EmpSalTxtLeavedays;
        private Label EmpSalLblHolidays;
        private Label EmpSalLblBasicpay;
        private Label EmpSalLblPresentdays;
        private TextBox EmpSalTxtAbsentdays;
        private TextBox EmpSalTxtPresentdays;
        private Label EmpSalLblAbsentdays;
        private Label EmpSalLblNetsalary;
        private CheckBox EmpSalCkbDeducamount;
        private CheckBox EmpSalCkbAllowamount;
        private TextBox EmpSalTxtTotaldays;
        private Label EmpSalLblTotaldays;
        private TextBox EmpSalTxtNetsalary;
        private Label EmpSalLblFollow0;
        private TextBox EmpSalTxtDeducamount;
        private TextBox EmpSalTxtAllowamount;
        private TextBox EmpSalTxtBasicpay;
        private GroupBox EmpSalGroupSalarydetail;
        private GroupBox EmpSalGroupAlloDeductdetail;
        private TextBox EmpSalTxtOtherallow;
        private TextBox EmpSalTxtConveyanceallow;
        private TextBox EmpSalTxtLatesittingallow;
        private TextBox EmpSalTxtMedicalallow;
        private TextBox EmpSalTxtPhoneallow;
        private TextBox EmpSalTxtHouseallow;
        private DataGridViewButtonColumn ViewSalaryDetail;
        private Label EmpSalLblConvallow;
        private Label EmpSalLblMedicalallow;
        private Label EmpSalLblHouseallow;
        private Label EmpSalLblPhoneallow;
        private Label EmpSalLblOtherallow;
        private Label EmpSalLblLatesittingallow;
        private Label EmpSalLblOtherdeduc;
        private Label EmpSalLblAbsentdeduc;
        private Label EmpSalLblLatededuc;
        private TextBox EmpSalTxtOtherdeduc;
        private TextBox EmpSalTxtAdvadeduc;
        private TextBox EmpSalTxtLatededuc;
        private TabControl Management;
        private TabPage TabDepartments;
        private GroupBox MgmDeptGroupDetailvu;
        private DataGridView MgmDeptGridvu;
        private Label MgmDeptLblDeptscaleview;
        private GroupBox MgmDeptGroupoption;
        private Label MgmDeptLblDeptScalechoose;
        private RadioButton MgmDeptRbJobscales;
        private RadioButton MgmDeptRbDept;
        private Button MgmDeptBtnAddrow;
        private Button MgmDeptBtnEdit;
        private GroupBox MgmDeptGroupadddeptscale;
        private Label MgmDeptLblOtherdeduc;
        private TextBox MgmDeptTxtOtherdeduc;
        private Label MgmDeptLblLatededuc;
        private TextBox MgmDeptTxtLatededuc;
        private Label MgmDeptLblLatesitallow;
        private TextBox MgmDeptTxtLatesitallow;
        private Button MgmDeptBtnDelete;
        private Button MgmDeptBtnCancel;
        private Label MgmDeptLblOtherallow;
        private TextBox MgmDeptTxtOtherallow;
        private Label MgmDeptLblPhoneallow;
        private TextBox MgmDeptTxtPhoneallow;
        private Label MgmDeptLblConvallow;
        private TextBox MgmDeptTxtConvallow;
        private Label MgmDeptLblMedicalallow;
        private TextBox MgmDeptTxtMedicalallow;
        private Label MgmDeptLblHouserent;
        private TextBox MgmDeptTxtHouserent;
        private Label MgmDeptLblPerIncr;
        private Label MgmDeptLblDeptScaleadd;
        private TextBox MgmDeptTxtPerIncr;
        private Label MgmDeptLblDid;
        private Label MgmDeptLblMaxbpay;
        private Label MgmDeptLblDname;
        private Label MgmDeptLblMgrid;
        private TextBox MgmDeptTxtMaxbpay;
        private Button MgmDeptBtnSave;
        private TextBox MgmDeptTxtMgrid;
        private Label MgmDeptLblMinbpay;
        private TextBox MgmDeptTxtDname;
        private TextBox MgmDeptTxtDid;
        private TextBox MgmDeptTxtMinbpay;
        private Label MgmDeptLblBps;
        private TextBox MgmDeptTxtBps;
        private Label MgmDeptLblJobid;
        private TextBox MgmDeptTxtJobid;
        private Button MgmDeptBtnRefresh;
        private TabPage TabJob;
        private GroupBox MgmJobGroupSelectedemp;
        private GroupBox MgmJobGroupRestricted;
        private Label label3;
        private Label label21;
        private PictureBox pictureBox5;
        private GroupBox MgmJobGroupeEmpdetail;
        private ComboBox MgmJobCmb1Dept;
        private Label MgmJobLblJoindate;
        private Label MgmJobLblJobid;
        private DateTimePicker MgmJobCtlHiredate;
        private Label MgmJobLbl1Dept;
        private Label MgmJobLblRftag;
        private TextBox MgmJobTxtName;
        private ComboBox MgmJobCmbEmptype;
        private Label MgmJobLblName;
        private Label MgmJobLblEmptype;
        private TextBox MgmJobTxtJobid;
        private TextBox MgmJobTxtRftag;
        private ComboBox MgmJobCmb0Jobid;
        private TextBox MgmJobTxtPermotedby;
        private Label MgmJobLblPermotedby;
        private Button MgmJobBtnSaverec;
        private Panel MgmJobPanelPermotion;
        private ComboBox MgmJobCmb3Dept;
        private Label MgmJobLblComments;
        private Label MgmJobLbl1Jobid;
        private Label MgmJobLbl2Dept;
        private TextBox MgmJobTxtComments;
        private ComboBox MgmJobCmb1Bps;
        private ComboBox MgmJobCmb1Jobid;
        private Label MgmJobLblBps1;
        private Button MgmJobBtnCancel;
        private TextBox MgmJobTxtFormid1;
        private Label MgmJobLblFormid;
        private PictureBox MgmJobPicbox;
        private Label MgmJobLblPic;
        private CheckBox MgmJobCkbPermotion;
        private GroupBox MgmJobGroupRestrictedpermote;
        private Label label25;
        private Label label29;
        private PictureBox pictureBox6;
        private GroupBox MgmJobGroupSearch;
        private Label MgmJobLblCaution;
        private TextBox MgmJobTxtformid0;
        private ComboBox MgmJobCmb0Dept;
        private CheckBox MgmJobCkbFormid;
        private Button MgmJobBtnSearch;
        private CheckBox MgmJobCkbdDeptwise;
        private CheckBox MgmJobCkbJobisnull;
        private Button MgmJobBtnEdit;
        private DataGridView MgmJobGridvu;
        private TabPage TabAttend;
        private GroupBox MgmAttGroupSearch;
        private CheckBox MgmAttCkbFormid;
        private Label MgmAttLblDatetime;
        private DateTimePicker MgmAttDTP;
        private Button MgmAttBtnSearch;
        private ComboBox MgmAttCmbDept;
        private Button MgmAttBtnClear;
        private TextBox MgmAttTxtFormid;
        private CheckBox MgmAttCkbRftag;
        private CheckBox MgmAttCkbDept;
        private TextBox MgmAttTxtRftag;
        private GroupBox MgmAttGroupDetailview;
        private DataGridView MgmAttGridAttdetailvu;
        private GroupBox MgmAttGroupEmpatt;
        private DataGridView MgmAttGridAttendVu;
        private DataGridViewButtonColumn Detail_View;
        private TabPage TabMgmGeneral;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem openToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripMenuItem toolsToolStripMenuItem;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripMenuItem contentsToolStripMenuItem;
        private ToolStripMenuItem indexToolStripMenuItem;
        private ToolStripMenuItem searchToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator5;
        private ToolStripMenuItem aboutToolStripMenuItem;
        private ToolStripMenuItem employeeToolStripMenuItem;
        private ToolStripMenuItem managementToolStripMenuItem;
        private ToolStripMenuItem scanningProcessToolStripMenuItem;
        private ToolStripMenuItem scanningToolStripMenuItem;
        private ToolStripMenuItem dataEntryToolStripMenuItem;
        private ToolStripMenuItem hRVerifyingToolStripMenuItem;
        private ToolStripMenuItem receptonToolStripMenuItem;
        private PictureBox pictureBox1;
        private GroupBox MgmGenGroup;
        private GroupBox MgmGenGroupLeave;
        private GroupBox MgmGenGroupBoxLeave;
        private Label label24;
        private TextBox MgmGenTxtLeavedates;
        private NumericUpDown numericUpDown1;
        private DateTimePicker MgmGenCtlDate1;
        private Label RecepLblLeavedates;
        private Label label26;
        private Label label27;
        private GroupBox MgmGenGroupboxSetholiday;
        private Label label30;
        private Label MgmGenLblReason;
        private ComboBox MgmGenCmbReason;
        private DateTimePicker MgmGenDttimeHoliday;
        private Label MgmGenLblSelecteddate;
        private GroupBox MgmGenGroupBoxStatus;
        private Label MgmGenLblSelectdate;
        private Label label34;
        private Label MgmGenLblStatus;
        private ComboBox MgmGenCmbStatus;
        private DateTimePicker MgmGenDtTime;
        private CheckBox MgmGenCkbSetstatus;
        private CheckBox MgmGenCkbSetholidays;
        private CheckBox MgmGenCkbSetleave;
        private Label MgmGenLblDepartment;
        private TextBox MgmGenTxtDepartment;
        private PictureBox MgmGenPicture;
        private TextBox MgmGenTxtBps;
        private Label MgmGenLblBps;
        private Label MgmGenLblName;
        private TextBox MgmGenTxtName;
        private GroupBox groupBox7;
        private TextBox MgmGenTxtSearchbyformid;
        private TextBox MgmGenTxtSearchbyrf;
        private Button MgmGenBtnClear;
        private Label label36;
        private Button MgmGenBtnSearch;
        private CheckBox MgmGenCkbSearchbyrf;
        private CheckBox MgmGenCkbSearchbyformid;
        private Button MgmGenBtnAllow;
        private Label MgmGenLblFormid;
        private TextBox MgmGenTxtFormid;
        private DataGridViewButtonColumn CheckDetaill;
        private Label richTextBox7;
        private Label LblJob_update;
        private Label label1;
        private Label label23;
        private Label label31;
        private Label label33;
        private Label label32;
        private Label label35;
        private TabPage TabReportViewer;
        private TabControl tabControl1;
        private TabPage TabReport_Employee;
        private CrystalDecisions.Windows.Forms.CrystalReportViewer CrysReport_Employee;
        private OA_PS_RFID.CryReport_Employee CryReport_Employee1;
        private Label label37;
        
    }
}