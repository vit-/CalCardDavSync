using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Net;
using System.Threading;
using System.Runtime.Serialization.Formatters.Binary;

using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using WebDav.Client;



namespace CalCardDavSync
{
    public partial class ThisAddIn
    {
        int threadSleepTime = 150000;
        bool threadExecuteSyncContacts = true;
        bool threadExecuteSyncCalendar = true;
        Thread threadSyncContacts;
        Thread threadSyncCalendar;

        string login = "test";
        string password = "test";
        string urlContacts = "";
        string urlCalendar = "";
        string storeName = "";
        
        WebDavSession session;

        Outlook.Folder olFolderContacts;
        Outlook.Folder olFolderCalendar;

        string tmpFilnameContacts;
        string tmpFilnameCalendar;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitVars();
            threadSyncContacts = new Thread(new ThreadStart(SyncContacts));
            //threadSyncCalendar = new Thread(new ThreadStart(SyncCalendar));

            threadSyncContacts.Start();
            //threadSyncCalendar.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            threadExecuteSyncContacts = false;
            threadExecuteSyncCalendar = false;
            
            threadSyncContacts.Join();
            threadSyncCalendar.Join();
        }

        private void InitVars()
        {
            Outlook.Store store = null;
            // There is no good method for searching store.
            foreach (Outlook.Store s in Application.Session.Stores)
            {
                if (s.DisplayName == storeName) 
                { 
                    store = s;
                    break;
                }
            }
            olFolderContacts = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts) as Outlook.Folder;
            olFolderCalendar = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;

            olFolderContacts.UserDefinedProperties.Add("remoteID", Outlook.OlUserPropertyType.olText);
            olFolderContacts.UserDefinedProperties.Add("modifyDate", Outlook.OlUserPropertyType.olDateTime);

            tmpFilnameContacts = Path.GetTempFileName().Replace(".tmp", ".vcf");
            tmpFilnameCalendar = Path.GetTempFileName().Replace(".tmp", ".ics");

            session = new WebDavSession();
            session.Credentials = new NetworkCredential(login, password);
        }

        private void IterateItems<resourceType>(IFolder folder, Outlook.Folder olFolder, string tmpFilename)
        {
            List<string> processedIds = new List<string>();
            IHierarchyItem[] remoteItems = folder.GetChildren();
            foreach (IHierarchyItem remoteItem in remoteItems)
            {
                if (remoteItem.ItemType != ItemType.Resource) continue;
                processedIds.Add(remoteItem.DisplayName);
                IResource resource = folder.GetResource(remoteItem.DisplayName);

                Outlook.ContactItem existItem = (Outlook.ContactItem) olFolder.Items.Find(String.Format("[remoteID] = '{0}'", remoteItem.DisplayName));
                if (existItem != null)
                {
                    // if contact is not modified then skip.
                    // TODO fix dates. Time is rounded in UserProperty. WTF?
                    if (remoteItem.LastModified != existItem.UserProperties.Find("modifyDate").Value)
                    {
                        existItem.Delete();
                    }
                    else
                    {
                        continue;
                    }
                }

                // TODO replace this code with in-memory procedure
                using (Stream stream = resource.GetReadStream())
                {
                    using (Stream f = File.OpenWrite(tmpFilename))
                    {
                        CopyStream(stream, f);
                    }
                }

                // TODO fix bug with encodings. UTF8 is recognized as CP1252. WTF?
                Outlook.ContactItem item = (Outlook.ContactItem) Application.Session.OpenSharedItem(tmpFilename);
                File.Delete(tmpFilename);
                
                item.UserProperties.Add("remoteID", Outlook.OlUserPropertyType.olText).Value = remoteItem.DisplayName;
                item.UserProperties.Add("modifyDate", Outlook.OlUserPropertyType.olDateTime).Value = remoteItem.LastModified;
                item.Move(olFolder);
                item.Save();
            }
            string filter = String.Empty;
            foreach (string id in processedIds)
            {
                filter += String.Format("[remoteID] <> '{0}' And ", id);
            }
            Outlook.Items itemsForDelete = olFolder.Items;
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 5);
                itemsForDelete = itemsForDelete.Restrict(filter);
            }
            foreach (Outlook.ContactItem item in itemsForDelete)
            {
                item.Delete();
            }
        }

        private void SyncContacts()
        {
            while (threadExecuteSyncContacts)
            {
                IterateItems<Outlook.ContactItem>(session.OpenFolder(urlContacts), olFolderContacts, tmpFilnameContacts);
                Thread.Sleep(threadSleepTime);
            }
        }

        /// <summary>
        /// Copies the contents of input to output. Doesn't close either stream.
        /// </summary>
        public static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
