using System;
using Microsoft.Exchange.WebServices.Data;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace ReadOffice365Mailbox
{
    class Program
    {
        static void Main(string[] args)
        {
            EmailReading();
            //you can show message maybe.
        }
    }
    //get all folders in exchange online mailbox
    public void GetAllFolders(ExchangeService service, List<Folder> completeListOfFolderIds)
    {
        FolderView folderView = new FolderView(100);
        FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.Inbox, folderView);
        foreach (Folder folder in findFolderResults)
        {
            completeListOfFolderIds.Add(folder);
            FindAllSubFolders(service, folder.Id, completeListOfFolderIds);
        }
    }
    //find all subfolders 
    public void FindAllSubFolders(ExchangeService service, FolderId parentFolderId, List<Folder> completeListOfFolderIds)
    {
        //alt klasörleri bul
        FolderView folderView = new FolderView(100);
        FindFoldersResults foundFolders = service.FindFolders(parentFolderId, folderView);


        completeListOfFolderIds.AddRange(foundFolders);
        foreach (Folder folder in foundFolders)
        {
            FindAllSubFolders(service, folder.Id, completeListOfFolderIds);
        }
    }

    public void EmailReading()
    {
        ExchangeService _service;
        try
        {
            _service = new ExchangeService
            {
                //change credentials
                Credentials = new WebCredentials("mail@mail.com", "MailPassword"),
                Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")
            };
        }
        catch
        {
            return;
        }

        List<Folder> completeListOfFolderIds = new List<Folder>();
        GetAllFolders(_service, completeListOfFolderIds);
        EmailLoad db = new EmailLoad();

        foreach (Folder folder in completeListOfFolderIds)
        {
            string folderName = folder.DisplayName;
            //last 3 email in the folder
            ItemView itemView = new ItemView(3);
            FindItemsResults<Item> searchResults = _service.FindItems(folder.Id, itemView);
            try
            {
                foreach (EmailMessage email in searchResults)
                {
                    db.Save(email, folderName);
                }
            }
            catch (Exception e)
            {
                //do something
            }
        }
    }
}
