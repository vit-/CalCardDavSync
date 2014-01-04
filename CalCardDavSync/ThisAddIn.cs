﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Net;
using System.Threading;

using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using WebDav.Client;



namespace CalCardDavSync
{
    public partial class ThisAddIn
    {
        int threadSleepTime = 50000;
        bool threadExecuteSyncContacts = true;
        bool threadExecuteSyncCalendar = true;
        Thread threadSyncContacts;
        Thread threadSyncCalendar;

        string login = "test";
        string password = "test";
        string urlContacts = "";
        string urlCalendar = "";
        
        WebDavSession session;
        IFolder folderContacts;
        IFolder folderCalendar;

        Outlook.Folder olFolderContacts;
        Outlook.Folder olFolderCalendar;

        string tmpFilnameContacts;
        string tmpFilnameCalendar;
        
        Dictionary<string, string> trackContacts = new Dictionary<string, string>();
        Dictionary<string, string> trackCalendar = new Dictionary<string, string>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitVars();
            threadSyncContacts = new Thread(new ThreadStart(SyncContacts));
            threadSyncCalendar = new Thread(new ThreadStart(SyncCalendar));

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
            Outlook.Store store = Application.Session.Stores[1]; // TODO this is bad
            olFolderContacts = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts) as Outlook.Folder;
            olFolderCalendar = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
            tmpFilnameContacts = Path.GetTempFileName().Replace(".tmp", ".vcf");
            tmpFilnameCalendar = Path.GetTempFileName().Replace(".tmp", ".ics");

            session = new WebDavSession();
            session.Credentials = new NetworkCredential(login, password);
            folderContacts = session.OpenFolder(urlContacts);
            folderCalendar = session.OpenFolder(urlCalendar);
        }

        private void SyncContacts()
        {
            while (threadExecuteSyncContacts)
            {
                IHierarchyItem[] remoteContacts = folderContacts.GetChildren();
                foreach (IHierarchyItem remoteContact in remoteContacts)
                {
                    if (trackContacts.ContainsKey(remoteContact.DisplayName)) continue;
                    if (remoteContact.ItemType == ItemType.Resource)
                    {
                        IResource resource = folderContacts.GetResource(remoteContact.DisplayName);

                        // TODO replace this code with in-memory procedure
                        Stream stream = resource.GetReadStream();
                        using (Stream file = File.OpenWrite(tmpFilnameContacts))
                        {
                            CopyStream(stream, file);
                        }
                        stream.Close();

                        Outlook.ContactItem contact = Application.Session.OpenSharedItem(tmpFilnameContacts) as Outlook.ContactItem;

                        Outlook.ContactItem existContact = olFolderContacts.Items.Find(String.Format("[FirstName]='{0}' and [LastName]='{1}'", 
                            contact.FirstName, contact.LastName));
                        if (existContact != null) existContact.Delete();

                        contact.Move(olFolderContacts);
                        contact.Save();

                        trackContacts.Add(remoteContact.DisplayName, contact.EntryID);
                    }
                }
                Thread.Sleep(threadSleepTime);
            }
        }

        private void SyncCalendar()
        {
            while (threadExecuteSyncCalendar)
            {
                IHierarchyItem[] remoteEvents = folderCalendar.GetChildren();
                foreach (IHierarchyItem remoteEvent in remoteEvents)
                {
                    if (trackCalendar.ContainsKey(remoteEvent.DisplayName)) continue;
                    if (remoteEvent.ItemType == ItemType.Resource)
                    {
                        IResource resource = folderCalendar.GetResource(remoteEvent.DisplayName);

                        // TODO replace this code with in-memory procedure
                        Stream stream = resource.GetReadStream();
                        using (Stream file = File.OpenWrite(tmpFilnameCalendar))
                        {
                            CopyStream(stream, file);
                        }
                        stream.Close();

                        Outlook.AppointmentItem appointment = Application.Session.OpenSharedItem(tmpFilnameCalendar) as Outlook.AppointmentItem;
                        appointment.Move(olFolderCalendar);
                        appointment.Save();

                        trackCalendar.Add(remoteEvent.DisplayName, appointment.EntryID);
                    }
                }
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