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
    partial class DeleteForms
    {
        // Required designer variable.
        //
        //
        //Data members only.

        string strpath = "C:\\Deleted_Formsid\\";
        static string StrOrcl9i = "Data Source=(DESCRIPTION="
             + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + Safeinfo.machinename + ")(PORT=1521)))"
             + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" + Safeinfo.servicename + ")));"
             + "User Id=pk_ocp;Password=me;";
        OracleConnection con = new OracleConnection(StrOrcl9i);
        DataSet mydataset;
        int Total_Rows_4Delete;
        

        /*****************************************************************************************/
        /**********************************Now my Functions.**************************************/

        int show_delformidfun()    //Its shows info where Form_id=-1
        {
            try
            {                
                con.Open();
                OracleDataAdapter myadapter= new OracleDataAdapter("Select form_id,formstatus_id,deo_by from Employee where formstatus_id=-1",con);
                mydataset = new DataSet("Employee");
                myadapter.Fill(mydataset, "Employee");
                GridDelVu.DataSource = mydataset.Tables[0];
                Total_Rows_4Delete=mydataset.Tables[0].Rows.Count;
                if (mydataset.Tables[0].Rows.Count >= 1)
                {//If data available for delete then enable both buttons.
                   
                    //GridDelVu.Columns[0].ReadOnly = false;  //Now user can check/Uncheck.
                    //GridDelVu.Columns[1].ReadOnly = true;   //Now user can't check/Uncheck.
                    //GridDelVu.Columns[2].ReadOnly = false;  //Now user can check/Uncheck.                    
                    //GridDelVu.Columns[3].ReadOnly = true;   //Now user can't check/Uncheck.
                    return (1);
                }
                else
                {//If data not available for delete then Disable both buttons.
                    
                    return (0);
                }


            }//Try end
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return (-1);
            }
            finally
            {
                //MessageBox.Show("I'm going to close Connection.");
                con.Close();
            }
        }
        void deleterowsfun()        //It will delete all rows at status=-1.
        { //It will delete all rows at status=-1.
            try
            {                            
                con.Open();
                string dele_query = "Delete from Employee where formstatus_id=-1";
                OracleCommand cmd = new OracleCommand(dele_query, con);
                MessageBox.Show("Effected Rows are:"+cmd.ExecuteNonQuery());

            }//Try end
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                //MessageBox.Show("I'm going to close Connection.");
                con.Close();
            }        
        }
        void delete_resetfun(bool delete_or_reset,int disform_id)
        {
            if (delete_or_reset)
            {//Its mean you want to re-set forms.
                DialogResult result;
                result = MessageBox.Show("This form is being Re-Set back to DEO's.\nAre you sure?", "Transaction Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Directory.CreateDirectory(strpath);             //Creating directory.
                        string f_name = "Resetted_Formids.txt";                   //File Name.
                        FileStream fs = File.Create(strpath + f_name);  //Creating Files.
                        StreamWriter sw = new StreamWriter(fs);
                        sw.WriteLine("Here your Re-setted Form_ids.");
                        sw.WriteLine("---------------------------.");

                        //--
                        {//Write all forms_id and (Re-set date+time) in a file.
                            sw.WriteLine("Forms id :" + disform_id + " Re-Seted on: " + DateTime.Today.Date.ToString() + ", Re-set By:" + Deo_Module.Safeinfo.Empname);
                        }
                        sw.Close();
                        fs.Close();
                        //--

                        string dele_query = "Update Employee set formstatus_id=0 Where Form_id=" + disform_id;
                        Update_Database myobj = new Update_Database();
                        int total_effrows = myobj.Dml_Updatefun(dele_query);//Now Reset formstatus_id=0;
                        if (total_effrows >= 1)
                        {
                            MessageBox.Show("Form has been Re-setted Successfully.","Re-Set Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else MessageBox.Show("Sorry!\nForm Can not be Re-set back to DEO.\nPlease try later.", "Re-Set Status", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                }
            }// Re-set form, IF condition OFF.
            else
            {
                //Its mean you want to Delete this forms.
                DialogResult result;
                result = MessageBox.Show("It will delete this form Permanently from the Database.\nAre you sure?", "Delete Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Directory.CreateDirectory(strpath);             //Creating directory.
                        string f_name = "Deleted_Formids.txt";                   //File Name.
                        FileStream fs = File.Create(strpath + f_name);  //Creating Files.
                        StreamWriter sw = new StreamWriter(fs);
                        sw.WriteLine("Here your deldeted Form_ids.");
                        sw.WriteLine("---------------------------.");

                       
                        {//Write all forms_id and deletion date+time in a file.
                            sw.WriteLine("Forms id :" + disform_id + " Deleted on: " + DateTime.Today.Date.ToString() + ", Deleted By:" + Deo_Module.Safeinfo.Empname);
                        }
                        sw.Close();
                        fs.Close();

                        //System.IO.FileInfo fi = new System.IO.FileInfo(strpath + f_name);
                        //fi.Encrypt();

                        string dele_query = "Delete from Employee where formstatus_id=-1 And Form_id=" + disform_id;
                        Update_Database myobj = new Update_Database();
                        int total_effrows = myobj.Dml_Updatefun(dele_query);//Now delete this row from database.
                        if (total_effrows >= 1)
                        {
                            MessageBox.Show("Record deleted Successfully.","Effected Row Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else MessageBox.Show("Sorry!\nRecord can not deleted from Database.","Effected Row Status", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        
                    }catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                }//Dialog confirmation.
            }// Deletion form, Ese condition OFF.
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DeleteForms));
            this.GridDelVu = new System.Windows.Forms.DataGridView();
            this.Action = new System.Windows.Forms.DataGridViewButtonColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewButtonColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnRefresh = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.DeleteFormsRichTxtDelforms = new System.Windows.Forms.RichTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.GridDelVu)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // GridDelVu
            // 
            this.GridDelVu.AllowUserToAddRows = false;
            this.GridDelVu.AllowUserToDeleteRows = false;
            this.GridDelVu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.GridDelVu.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Action,
            this.Column1});
            this.GridDelVu.Location = new System.Drawing.Point(138, 55);
            this.GridDelVu.Name = "GridDelVu";
            this.GridDelVu.ReadOnly = true;
            this.GridDelVu.Size = new System.Drawing.Size(593, 249);
            this.GridDelVu.TabIndex = 0;
            this.GridDelVu.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.GridDelVu_CellClick);            
            // 
            // Action
            // 
            this.Action.HeaderText = "Re-Store";
            this.Action.Name = "Action";
            this.Action.ReadOnly = true;
            this.Action.Text = "Re-Set";
            this.Action.UseColumnTextForButtonValue = true;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Remove";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Text = "Delete";
            this.Column1.UseColumnTextForButtonValue = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(124, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(220, 16);
            this.label1.TabIndex = 3;
            this.label1.Text = "Delete Invalid Scanned Forms.";
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(663, 324);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(67, 21);
            this.BtnCancel.TabIndex = 3;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnRefresh
            // 
            this.BtnRefresh.Location = new System.Drawing.Point(39, 190);
            this.BtnRefresh.Name = "BtnRefresh";
            this.BtnRefresh.Size = new System.Drawing.Size(67, 21);
            this.BtnRefresh.TabIndex = 4;
            this.BtnRefresh.Text = "Refresh";
            this.BtnRefresh.UseVisualStyleBackColor = true;
            this.BtnRefresh.Click += new System.EventHandler(this.BtnRefresh_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pictureBox1.Location = new System.Drawing.Point(12, 55);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(114, 129);
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // DeleteFormsRichTxtDelforms
            // 
            this.DeleteFormsRichTxtDelforms.BackColor = System.Drawing.Color.WhiteSmoke;
            this.DeleteFormsRichTxtDelforms.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DeleteFormsRichTxtDelforms.Location = new System.Drawing.Point(-1, 0);
            this.DeleteFormsRichTxtDelforms.Name = "DeleteFormsRichTxtDelforms";
            this.DeleteFormsRichTxtDelforms.ReadOnly = true;
            this.DeleteFormsRichTxtDelforms.Size = new System.Drawing.Size(827, 29);
            this.DeleteFormsRichTxtDelforms.TabIndex = 56;
            this.DeleteFormsRichTxtDelforms.Text = "                                                      DELETED FORMS";
            // 
            // DeleteForms
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(742, 404);
            this.Controls.Add(this.DeleteFormsRichTxtDelforms);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.BtnRefresh);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.GridDelVu);
            this.HelpButton = true;
            this.Name = "DeleteForms";
            this.RightToLeftLayout = true;
            this.Text = "Delete Forms";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DeleteForms_FormClosing);
            this.Load += new System.EventHandler(this.DeleteForms_Load);
            ((System.ComponentModel.ISupportInitialize)(this.GridDelVu)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView GridDelVu;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnCancel;
        private Button BtnRefresh;
        private PictureBox pictureBox1;
        private RichTextBox DeleteFormsRichTxtDelforms;
        private DataGridViewButtonColumn Action;
        private DataGridViewButtonColumn Column1;



       
    }
}