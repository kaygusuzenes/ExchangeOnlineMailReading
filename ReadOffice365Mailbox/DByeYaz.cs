using Microsoft.Exchange.WebServices.Data;
using System;
using System.Data;
using System.Data.SqlClient;

namespace ReadOffice365Mailbox
{
    class DByeYaz
    {
        SqlConnection conn = new SqlConnection("Data Source=localhost;Initial Catalog=Backup;Integrated Security=true");

        public DByeYaz()
        {
            Console.WriteLine("Connecting to SQL Server");
            try
            {
                //conn = new SqlConnection("Data Source=localhost;Initial Catalog=Backup;Integrated Security=true");
                conn.Open();
                Console.WriteLine("Connected");
            }
            catch (SqlException e)
            {
                throw (e);
            }
        }

        public void Kaydet(EmailMessage email)
        {
            email.Load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));


            SqlCommand cmd = new SqlCommand("Insert into Emails (message_id,mailto,subject,body,received_time,mail_from) values (@message_id,@mailto,@subject,@body,@received_time,@mail_from)", conn);

            cmd.Parameters.AddWithValue("@message_id", email.InternetMessageId.ToString());
            cmd.Parameters.AddWithValue("@mailto", "backup@yesis.net");
            cmd.Parameters.AddWithValue("@subject", email.Subject.ToString());
            cmd.Parameters.AddWithValue("@body", email.TextBody.ToString());
            cmd.Parameters.AddWithValue("@received_time", Convert.ToDateTime(email.DateTimeReceived)); //email.DateTimeReceived.ToUniversalTime()));

            cmd.Parameters.AddWithValue("@mail_from", email.From.Address.ToString());

            try
            {
                if (email.Subject.Contains("failed") || email.Subject.Contains("failure")|| email.Subject.Contains("başarısız"))
                {
                    DataTable dtMails = GetDataTable("select * from Emails");

                    if (dtMails.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtMails.Rows.Count; i++)
                        {
                            using (cmd)
                            {
                                DataRow dr = dtMails.Rows[i];

                                //   var tarih = Convert.ToDateTime(dr["received_time"]);
                                DataTable dtIsAdded = GetDataTable("select * from Emails where subject= '" + dr["subject"] + "' and DATEDIFF(DAY,received_time,GETDATE()) between 0 and 7");
                                if (dtIsAdded.Rows.Count == 0)
                                {
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                    }
                }
            }

            catch (Exception ex)
            {

                Console.WriteLine(ex);
            }
        }


        public DataTable GetDataTable(string sql)
        {
            SqlConnection baglan = this.conn;
            SqlDataAdapter adapter = new SqlDataAdapter(sql, baglan);
            DataTable dt = new DataTable();

            try
            {
                adapter.Fill(dt);
            }
            catch (SqlException ex)
            {
                throw new Exception(ex.Message + "(" + sql + ")");
            }

            adapter.Dispose();
            baglan.Close();
            baglan.Dispose();
            return dt;
        }


    }
}