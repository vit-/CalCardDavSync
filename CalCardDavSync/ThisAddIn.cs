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
    public class DataTrack
    {
        public class ItemData
        {
            public string RemoteID {get; private set;}
            public string LocalID { get; private set; }
            public DateTime ModifyDate {get; private set;}
            public int ID {get; private set;}

            public ItemData(int id, string remoteID, string localID, DateTime modifyDate)
            {
                ID = id;
                RemoteID = remoteID;
                LocalID = localID;
                ModifyDate = modifyDate;
            }
        }
        private int counter = 0;
        private Dictionary<string, List<string>> remoteIDs = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> localIDs = new Dictionary<string, List<string>>();
        private Dictionary<int, DateTime> modifyDates = new Dictionary<int, DateTime>();

        public void Add(string remoteID, string localID, DateTime modifyDate)
        {
            remoteIDs.Add(remoteID, new List<string>() {localID, counter.ToString()});
            localIDs.Add(localID, new List<string>() {remoteID, counter.ToString()});
            modifyDates.Add(counter, modifyDate);
            counter++;
        }

        public void RemoveByRemoteID(string remoteID)
        {
            ItemData item = GetByRemoteID(remoteID);
            Remove(item);
        }

        public void RemoveByLocalID(string localID)
        {
            ItemData item = GetByLocalID(localID);
            Remove(item);
        }

        public void Remove(ItemData item)
        {
            remoteIDs.Remove(item.RemoteID);
            localIDs.Remove(item.LocalID);
            modifyDates.Remove(item.ID);
        }

        public ItemData GetByRemoteID(string remoteID)
        {
            if (!remoteIDs.ContainsKey(remoteID)) return null;
            var value = remoteIDs[remoteID];
            string localID = value[0];
            int id = Convert.ToInt32(value[1]);
            return new ItemData(id, remoteID, localID, modifyDates[id]);
        }

        public ItemData GetByLocalID(string localID)
        {
            if (!localIDs.ContainsKey(localID)) return null;
            var value = localIDs[localID];
            string remoteID = value[0];
            int id = Convert.ToInt32(value[1]);
            return new ItemData(id, remoteID, localID, modifyDates[id]);
        }
    }

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

        string syncStatusFilenameContacts;
        string syncStatusFilenameCalendar;
        Dictionary<string, string> trackContacts;
        Dictionary<string, string> trackCalendar;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitVars();
            LoadStats();
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

        private Dictionary<string, string> LoadStats(string filename)
        {
            Dictionary<string, string> result;
            BinaryFormatter formatter = new BinaryFormatter();
            try
            {
                using (Stream f = File.OpenRead(filename))
                {
                    result = formatter.Deserialize(f) as Dictionary<string, string>;
                }
            }
            catch { 
                return new Dictionary<string, string>();
            }
            return result;
        }

        private void LoadStats()
        {
            string appDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            appDir = Path.Combine(appDir, "CalCardDavSync");
            if (!Directory.Exists(appDir)) Directory.CreateDirectory(appDir);

            syncStatusFilenameContacts = Path.Combine(appDir, "synccontacts.dat");
            syncStatusFilenameCalendar = Path.Combine(appDir, "synccalendar.dat");

            trackContacts = LoadStats(syncStatusFilenameContacts);
            trackCalendar = LoadStats(syncStatusFilenameCalendar);
        }

        private void SaveStats(string filename, Dictionary<string, string> dict){
            BinaryFormatter formatter = new BinaryFormatter();
            using (Stream f = File.OpenWrite(filename))
            {
                formatter.Serialize(f, dict);
            }
        }

        private void SaveStats()
        {
            SaveStats(syncStatusFilenameContacts, trackContacts);
            SaveStats(syncStatusFilenameCalendar, trackCalendar);
        }

        private void IterateItems<resourceType>(IFolder folder, Outlook.Folder olFolder, Dictionary<string, string> trackDict, string tmpFilename)
        {
            IHierarchyItem[] remoteItems = folder.GetChildren();
            foreach (IHierarchyItem remoteItem in remoteItems)
            {
                if (remoteItem.ItemType != ItemType.Resource) continue;
                IResource resource = folder.GetResource(remoteItem.DisplayName);

                resourceType existItem = default(resourceType);
                //if (trackDict.ContainsKey(remoteItem.DisplayName))
                //    existItem = olFolder.Items.Find(String.Format("[EntryID]='{0}'", trackDict[remoteItem.DisplayName]));
                
                // sync decision here
                if (existItem != null) continue;
                
                // TODO replace this code with in-memory procedure
                using (Stream stream = resource.GetReadStream())
                {
                    using (Stream f = File.OpenWrite(tmpFilename))
                    {
                        CopyStream(stream, f);
                    }
                }

                Outlook.ContactItem item = (Outlook.ContactItem)Application.Session.OpenSharedItem(tmpFilename);
                item.Move(olFolder);
                item.Save();
                if (!trackDict.ContainsKey(remoteItem.DisplayName))
                    trackDict.Add(remoteItem.DisplayName, item.EntryID);
            }
        }

        private void SyncContacts()
        {
            while (threadExecuteSyncContacts)
            {
                IterateItems<Outlook.ContactItem>(folderContacts, olFolderContacts, trackContacts, tmpFilnameContacts);
                SaveStats(syncStatusFilenameContacts, trackContacts);
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
                SaveStats(syncStatusFilenameCalendar, trackCalendar);
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
