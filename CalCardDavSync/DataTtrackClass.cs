using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalCardDavSync
{
    public class DataTrack
    {
        public class ItemData
        {
            public string RemoteID { get; private set; }
            public string LocalID { get; private set; }
            public DateTime ModifyDate { get; private set; }
            public int ID { get; private set; }

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
            remoteIDs.Add(remoteID, new List<string>() { localID, counter.ToString() });
            localIDs.Add(localID, new List<string>() { remoteID, counter.ToString() });
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
}
