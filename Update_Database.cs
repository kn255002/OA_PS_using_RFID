using System;
using System.Data;
using System.Data.OracleClient;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;

namespace Deo_Module
{   
    class Update_Database
    {
        /// <summary>
        /// This class Main purpose to HANDEL DML/DDL/DRL Requests. And its working find.


        private static string StrOrcl9i = "Data Source=(DESCRIPTION="
             + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=12345)(PORT=1521)))"
             + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=XXXXX)));"
             + " User Id=Pk_ocp ;Password=me;";

        OracleCommandBuilder Ocmdbldr;           //For Auto querying.                
        public OracleConnection cls_con ;
        public OracleDataReader cls_dr;         //I made it bcz Now at every form by making object you can access it, can use it.
        public DataSet cls_dataset;             //I made it bcz Now at every form by making object you can access it, can use it.
        public OracleDataAdapter cls_odadptr;   //I made it bcz Now at every form by making object you can access it, can use it.
        //public object cls_scalerobj;
        public Update_Database() 
        {//Zero Argument Constructor.
            cls_dataset = null; 
            cls_odadptr = null;
            string Temp = StrOrcl9i.Replace("XXXXX", "Orcl9i");
            Temp = Temp.Replace("12345", "Pavilion-211178");
            cls_con = new OracleConnection(Temp);

            Safeinfo.machinename = "Pavilion-211178".ToUpper();
            Safeinfo.servicename = "Orcl9i".ToUpper();

        }

        public Update_Database(string serviceName,string machineName)
        {//Two Arguments Constructor.
            if (machineName.Length == 0)
            {
                Safeinfo.machinename = machineName = "Pavilion-211178".ToUpper();                
            }
            if (serviceName.Length == 0)
            {
                Safeinfo.servicename=serviceName = "Orcl9i".ToUpper();                
            }

            string Temp = StrOrcl9i.Replace("XXXXX",serviceName);
            Temp = Temp.Replace("12345", machineName);
            cls_con = new OracleConnection(Temp);
        }

        /*****************************************************************************************************************/
        /***************************************Give Below Definition of All Methods.*************************************/
        
        
        
        
        
        // Now Below Defination of all Methods starts.       
        public void Show_Datafun(string show_result_query, int vclose)
        { //Now this Method must run two time. 
            //01 For Fetching data and keeping connectiong OPEN vclose=0. 
            //And the last one for to CLOSE connection when vclose=1(Do not fetch data);
            try
            {
                switch (vclose)
                {
                    case 0:
                        {
                            if (cls_con.State.ToString().ToUpper()=="OPEN") { cls_con.Close(); }
                            cls_con.Open();
                            OracleCommand cmd = new OracleCommand(show_result_query, cls_con);
                            cls_dr = cmd.ExecuteReader();
                            break;
                        }
                    default:
                        {
                            cls_con.Close();
                            //if (Safeinfo.status != "ADMINISTRATOR")
                            //{
                            //    cls_con.Dispose();
                            //}
                            break;
                        }
                }//End switch.
            }
            catch (RowNotInTableException ex) 
            {
                MessageBox.Show(ex.Message,"Record Not Found",MessageBoxButtons.OK,MessageBoxIcon.Error); 
            }
            catch (OracleException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
                //con.Close();
            }                  
        }
        public void Dml_UpdateAdapterfun(string exe_query)  //This is Fun overloading concept.
        {
            try
            {
                cls_con.Open();
                cls_odadptr = new OracleDataAdapter(exe_query, cls_con);
                cls_dataset = new DataSet();
                cls_odadptr.Fill(cls_dataset);
                Ocmdbldr = new OracleCommandBuilder(cls_odadptr);
            }

            catch (RowNotInTableException ex)
            {
                MessageBox.Show(ex.Message, "Record Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error); 
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
                cls_con.Close();
                //if (Safeinfo.status != "ADMINISTRATOR")
                //{
                //    cls_con.Dispose();
                //}
            }
        }
        public int  Dml_Updatefun(string exe_query)  //This is Fun overloading concept.
        {  //You just pass Query as string. It will update database and return Number of Effected rows.             
            try
            {
                cls_con.Open();
                OracleCommand cmd = new OracleCommand(exe_query, cls_con);
                int effected_rows = cmd.ExecuteNonQuery();
                return (effected_rows);
            }
            catch (OracleException Oex) 
            {
                MessageBox.Show(Oex.Message,"Oracle Error Found",MessageBoxButtons.OK,MessageBoxIcon.Error);                
                return (0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
                return (0);
            }
            finally
            {
                // MessageBox.Show("I'm going to close connection.");
                cls_con.Close();
                //if (Safeinfo.status != "ADMINISTRATOR")
                //{
                //    cls_con.Dispose();
                //}
            }
        }
        public void Grid_Loaderfun(DataGridView mygridvu, string run_query)
        {
            try
            {
                cls_con.Open();
                cls_odadptr = new OracleDataAdapter(run_query, cls_con);
                cls_dataset = new DataSet();
                cls_odadptr.Fill(cls_dataset);                        
                mygridvu.DataSource = cls_dataset.Tables[0];                
            }
            catch (RowNotInTableException ex)
            {
                MessageBox.Show(ex.Message, "Record Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error); 
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
                cls_con.Close();
                //if (Safeinfo.status != "ADMINISTRATOR")
                //{
                //    cls_con.Dispose();
                //}
            }     
        
        }        
        public void Exec_Scalerfun(string query)
        {
            try
            {
                cls_con.Open();
                OracleCommand cmd = new OracleCommand(query, cls_con);
                cmd.ExecuteScalar();
                
            }
            catch (RowNotInTableException ex) 
            {
                MessageBox.Show(ex.Message, "Record Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                cls_con.Close();
                //if (Safeinfo.status != "ADMINISTRATOR")
                //{
                //    cls_con.Dispose();
                //}
            }
        }//End Exec_Scaler().
        public int Record_Foundfun(string exe_query)
        {
            try
            {
                cls_con.Open();
                cls_odadptr = new OracleDataAdapter(exe_query, cls_con);
                cls_dataset = new DataSet();
                cls_odadptr.Fill(cls_dataset);
                if (cls_dataset.Tables[0].Rows.Count>= 1)
                { //Recor found against this query.
                    return (1);
                }
                else
                { //Recor NOT Found against this query.
                    return (0);
                }
            }

            catch (RowNotInTableException)
            {
                return (0);
            }
            catch (OracleException Oex)
            {
                MessageBox.Show(Oex.Message, "Oracle Error Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Exception Caught, Khurram :\n" + ex.ToString());
                return (0);
            }
            finally
            {
                // MessageBox.Show("I'm going to close connection.");
                cls_con.Close();
                //if (Safeinfo.status != "ADMINISTRATOR")
                //{
                //    cls_con.Dispose();
                //}
            }
        
        }

        
        




    }//Class ends.
}//namespace ends.
