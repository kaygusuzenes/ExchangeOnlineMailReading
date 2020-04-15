using System;
using Microsoft.Exchange.WebServices.Data;
using System.Data.SqlClient;

namespace ReadOffice365Mailbox
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService _service;

            try
            {
                Console.WriteLine("Exchange online servislerine bağlanılıyor.");

                _service = new ExchangeService
                {
                    Credentials = new WebCredentials("backup@yesis.net", "Alp7848**1")
                };
            }
            catch
            {
                Console.WriteLine("Exchange bağlantısı başarısız.");
                return;
            }
            _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            try
            {
                DByeYaz db = new DByeYaz();

                Console.WriteLine("Mailler okunuyor...");
                // okuyacağım mail sayısı
                foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.Inbox, new ItemView(100)))
                {
                    db.Kaydet(email);

                }

                Console.WriteLine("Kapanıyor.");
            }
            catch (Exception e)
            {
                Console.WriteLine("Bir hata oluştu. \n:" + e.Message);
            }
        }
    }
}