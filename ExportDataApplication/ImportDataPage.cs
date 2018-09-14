using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDataApplication
{
    public class ImportDataPage
    {
        public void InsertData()
        {
            string fileName = @"D:\DataBase Script\ScreenValuesFields14-09-2018.xlsx";
            string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";         
            OleDbConnection oconn = new OleDbConnection(ConnectionString);                                                                                                                                                       // connectionstring to connect to the Excel Sheet
            try
            {
                //After connecting to the Excel sheet here we are selecting the data 
                //using select statement from the Excel sheet
                OleDbCommand ocmd = new OleDbCommand("select * from [Sheet1$]", oconn);
                oconn.Open();  //Here [Sheet1$] is the name of the sheet 
                               //in the Excel file where the data is present
                OleDbDataReader odr = ocmd.ExecuteReader();
                string Id = "";
                string FieldName = "";
                string ScreenTypeValueId = "";
                string PMTypeId = "";
                while (odr.Read())
                {
                    Id = valid(odr, 0);//Here we are calling the valid method
                    FieldName = valid(odr, 1);
                    ScreenTypeValueId = valid(odr, 2);
                    PMTypeId = valid(odr, 3);
                    //Here using this method we are inserting the data into the database
                    insertdataintosql(Id, FieldName, ScreenTypeValueId, PMTypeId);
                }
                oconn.Close();
            }
            catch (DataException ex)
            {
                Console.WriteLine("error occured while storing the data");
            }
            finally
            {
                Console.WriteLine("Data Stored Successfully");
            }
        }
        protected string valid(OleDbDataReader myreader, int stval)//if any columns are 
                                                                   //found null then they are replaced by zero
        {
            object val = myreader[stval];
            if (val != DBNull.Value)
                return val.ToString();
            else
                return Convert.ToString(0);
        }

        public void insertdataintosql(string Id, string FieldName,
             string ScreenValueTypeId, string PMTypeId)
        {//inserting data into the Sql Server
            SqlConnection conn = new SqlConnection("Data Source=.\\sqlexpress;AttachDbFileName =| DataDirectory | exceltosql.mdf; Trusted_Connection = yes");
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = "insert into emp(FieldName,ScreenTypeValueId,PMTypeId,Unit,ValueType,IsDeleted,DateCreated,Default,Position)values(@fname, @lname, @mobnum, @city, @state, @zip)";
            cmd.Parameters.Add("@FieldName", SqlDbType.NVarChar).Value = FieldName;
            cmd.Parameters.Add("@ScreenValueTypeId", SqlDbType.Int).Value = Convert.ToInt32(ScreenValueTypeId);
            cmd.Parameters.Add("@PMTypeId", SqlDbType.Int).Value = Convert.ToInt32(PMTypeId);
            cmd.Parameters.Add("@Unit", SqlDbType.NVarChar).Value = "Anonymous";
            cmd.Parameters.Add("@ValueType", SqlDbType.NVarChar).Value = "Single";
            cmd.Parameters.Add("@IsDeleted", SqlDbType.Bit).Value =false;
            cmd.Parameters.Add("@DateCreated", SqlDbType.DateTime).Value = DateTime.Now;
            cmd.CommandType = CommandType.Text;
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

    }

}
