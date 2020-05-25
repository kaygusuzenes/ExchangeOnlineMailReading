using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;


public class EmailLoad
{   //create connection class for db connection
    //connection con=new connection();
    public EmailLoad()
    {
        try
        {
           // con.connectDb();
        }
        catch (SqlException e)
        {
            throw (e);
        }
    }

    public void Save(EmailMessage email, string folderName)
    {
        email.Load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));

        DateTime mailDate = Convert.ToDateTime(email.DateTimeReceived);
        DateTime date = Convert.ToDateTime(DateTime.Now.ToString());
        double day = (date - mailDate).TotalDays;

        //do something with this email data.

    }
}

