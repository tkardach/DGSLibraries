using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace SharePointLibrary
{
    public class SharePoint
    {
        private string _url;

        public SharePoint(string url)
        {
            _url = url;
        }

        #region Methods
        // SharePoint Website URL
        public string Url { set { _url = value; } get { return _url; } }

        /// <summary>
        /// GetListCollectionFromSharePointSite will return the ListCollection of the current SharePoint site.
        /// </summary>
        /// <returns>ListCollection containing all lists on the given SharePoint website.</returns>
        public ListCollection GetListCollectionFromSharePointSite()
        {
            return GetListCollectionFromSharePointSite(Url);
        }

        /// <summary>
        /// GetListCollectionFromSharePointSite will take a URL and return a ListCollection of all Lists 
        /// associated with that site.
        /// </summary>
        /// <param name="url">Url of the SharePoint website being queried.</param>
        /// <returns>ListCollection containing all lists on the given SharePoint website.</returns>
        public static ListCollection GetListCollectionFromSharePointSite(string url)
        {
            ClientContext context = new ClientContext(url);
            // The SharePoint web at the URL.
            Web web = context.Web;
            // Load the query into the context
            context.Load(web.Lists, lists => lists.Include(list => list.Title,
                                                           list => list.Id));
            try
            {
                // Execute query. 
                context.ExecuteQuery();
                return web.Lists;
            }
            catch (Microsoft.SharePoint.Client.ClientRequestException ex)
            {
                Console.Write(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// GetListItemCollectionByTitle takes a SharePoint URL and List Title and returns the ListItemCollection 
        /// associated with that information.
        /// </summary>
        /// <param name="listTitle">Title of the list being queried on the SharePoint website.</param>
        /// <returns>ListItemCollection of all items in the specified list on the SharePoint website.</returns>
        public ListItemCollection GetListItemCollectionByTitle(string listTitle)
        {
            return GetListItemCollectionByTitle(Url, listTitle);
        }

        /// <summary>
        /// GetListItemCollectionByTitle takes a SharePoint URL and List Title and returns the ListItemCollection 
        /// associated with that information.
        /// </summary>
        /// <param name="url">Url of the SharePoint website being queried.</param>
        /// <param name="listTitle">Title of the list being queried on the SharePoint website.</param>
        /// <returns>ListItemCollection of all items in the specified list on the SharePoint website.</returns>
        public static ListItemCollection GetListItemCollectionByTitle(string url, string listTitle)
        {
            ListItemCollection items = null;

            // Create a context using SharePoint URL, initialize the web using this URL.
            ClientContext context = new ClientContext(url);
            Web web = context.Web;

            try
            {
                // Get the desired list from the web page.
                List lists = context.Web.Lists.GetByTitle(listTitle);

                // Create a query to grab all items.
                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                items = lists.GetItems(query);

                // Retrieve all of the items returned from the query.
                context.Load(items);
                context.ExecuteQuery();
            }
            catch (Microsoft.SharePoint.Client.ClientRequestException ex)
            {
                Console.WriteLine("Error Connecting to URL: " + ex.Message);
            }

            return items;
        }

        /// <summary>
        /// Returns a list of all the field titles for the specified SharePoint List.
        /// </summary>
        /// <param name="listTitle">Title of the SharePoint list being queried.</param>
        /// <returns>A string list of all the field names.</returns>
        public List<string> GetListFields(string listTitle)
        {
            return GetListFields(Url, listTitle);
        }

        /// <summary>
        /// Returns a list of all the field titles for the specified SharePoint List.
        /// </summary>
        /// <param name="url">Url of the SharePoint site.</param>
        /// <param name="listTitle">Title of the SharePoint list being queried.</param>
        /// <returns>A string list of all the field names.</returns>
        public static List<string> GetListFields(string url, string listTitle)
        {
            List<string> fieldNames = new List<string>();
            // Create the context for the SharePoint site
            ClientContext context = new ClientContext(url);
            try
            {
                // Get the SharePoint List being queried
                List list = context.Web.Lists.GetByTitle(listTitle);
                // Load and execute the query
                context.Load(list.Fields);
                context.ExecuteQuery();
                // Itterate through each item in the list
                foreach (Field field in list.Fields)
                {
                    // Add the name of the current field to the fields list
                    fieldNames.Add(field.InternalName);
                }
            }
            catch (Microsoft.SharePoint.Client.ClientRequestException ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            return fieldNames;
        }
        #endregion
    }

    public class DGSSharePoint : SharePoint
    {
        private const string _dgsSharePointUrl = "/*DGS SharePoint URL for public repository*/";

        /// <summary>
        /// Constructs a SharePoint object using the DGS SharePoint url.
        /// </summary>
        public DGSSharePoint() : base(_dgsSharePointUrl) { }

        public static string DGSSharePointUrl { get { return _dgsSharePointUrl; } }
        #region Methods
        // Returns the ListItemCollection for the specified DGS SharePoint list
        protected static ListItemCollection GetListItems(string listName)
        {
            return GetListItemCollectionByTitle(DGSSharePointUrl, listName);
        }

        // Returns the ListCollection for the DGS website
        protected static ListCollection GetListCollection()
        {
            return GetListCollectionFromSharePointSite(DGSSharePointUrl);
        }
        #endregion
    }

    /// <summary>
    /// Representation of the DGS Server Inventory.
    /// </summary>
    public class DGSServerInventory : DGSSharePoint
    {
        // The values for these fields have been removed for privacy. 
        private const string _serverInventoryList = "/*Removed*/";  // Name of the SharePoint List
        private const string _supportStaffField = "/*Removed*/";    // FieldName for Support Staff
        private const string _serverDescriptionField = "/*Removed*/"; // FieldName for Description
        private const string _businessUnitField = "/*Removed*/";    // FieldName for Business Unit
        private const string _serverTypeField = "/*Removed*/";  // FieldName for Server Type
        private const string _environmentField = "/*Removed*/"; // FieldName for Environment

        private ServerCollection _serverInventory;

        /// <summary>
        /// Represents a Server in the DGS Server Inventory
        /// </summary>
        public class ServerItem
        {
            private String _serverName;
            private List<String> _supportStaff;
            private string _serverDescription;
            private string _businessUnit;
            private string _serverType;
            private string _environment;

            // Constructs the Server Item based off of the server name
            public ServerItem(String name)
            {
                _serverName = name.ToLower();
                _supportStaff = new List<string>();
                _serverDescription = "";
                _businessUnit = "";
                _serverType = "";
                _environment = "";
            }

            // Accessor methods
            public String ServerName { set { _serverName = value.ToLower(); } get { return _serverName; } }
            public List<String> SupportStaff { set { _supportStaff = value; } get { return _supportStaff; } }
            public string ServerDescription { set { _serverDescription = value; } get { return _serverDescription; } }
            public string BusinessUnit { set { _businessUnit = value; } get { return _businessUnit; } }
            public string ServerType { set { _serverType = value; } get { return _serverType; } }
            public string Environment { set { _environment = value; } get { return _environment; } }

            // Support staff List methods
            public void AddStaff(string staff) { _supportStaff.Add(staff); }
            public bool RemoveStaff(string staff) { return _supportStaff.Remove(staff); }

            /// <summary>
            /// Returns a string detailing this ServerItem.
            /// </summary>
            /// <returns>String representation of the ServerItem.</returns>
            public override string ToString()
            {
                string str = string.Format("{0, -15} {1, 1}", "Server Name", ": ") + _serverName + "\n";
                string staffStr = "";
                foreach (string staff in _supportStaff) staffStr += staff;
                str += string.Format("{0, -15} {1, 1}", "Support Staff", ": ") + staffStr + "\n";
                str += string.Format("{0, -15} {1, 1}", "Server Type", ": ") + _serverType + "\n";
                str += string.Format("{0, -15} {1, 1}", "Description", ": ") + _serverDescription + "\n";
                str += string.Format("{0, -15} {1, 1}", "Business Unit", ": ") + _businessUnit + "\n";
                str += string.Format("{0, -15} {1, 1}", "Environment", ": ") + _environment;
                return str;
            }
        }

        /// <summary>
        /// Collection of ServerItems.
        /// </summary>
        public class ServerCollection : Dictionary<string, ServerItem>
        {
            // Changes the function of the indexer to only search for lowercase servers.
            public new ServerItem this[string key]
            {
                get
                {
                    key = key.ToLower();
                    if (ContainsKey(key))
                        return base[key];
                    else
                        return null;
                }
            }

            /// <summary>
            /// Adds the ServerItem to the ServerInventory.
            /// </summary>
            /// <param name="server"></param>
            public void Add(ServerItem server)
            {
                Add(server.ServerName, server);
            }
            
            /// <summary>
            /// Remove the ServerItem from the Collection.
            /// </summary>
            /// <param name="server">Server being removed.</param>
            /// <returns>True if successfully removed.</returns>
            public bool Remove(ServerItem server)
            {
                if (Remove(server.ServerName)) { return true; }
                else {  return false; }
            }
        }

        /// <summary>
        /// Constructor initializes using DGS SharePoint URL and creates the DGS Server Inventory
        /// as a Dictionary of ServerItem objects.
        /// </summary>
        public DGSServerInventory()
        {
            _serverInventory = CreateServerItemList();
        }

        #region Methods
        public static string ServerInventoryListName { get { return _serverInventoryList; } }

        // Accessor for the Server Inventory Dictionary
        public ServerCollection ServerInventory { get { return _serverInventory; } }

        // Returns the ListItemCollection containing the DGS Server Inventory items
        private static ListItemCollection GetServerListItems()
        {
            return GetListItemCollection(DGSSharePointUrl, ServerInventoryListName);
        }

        // Initializes and returns a ListItemCollection of the Server Inventory entries
        private new static ListItemCollection GetListItemCollection(string url, string listTitle)
        {
            ListItemCollection itemCollection = null;

            // Create a context using SharePoint URL, initialize the web using this URL.
            ClientContext context = new ClientContext(url);
            Web web = context.Web;

            try
            {
                // Get the desired list from the web page.
                List lists = context.Web.Lists.GetByTitle(listTitle);

                // Create a query to grab all items.
                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                itemCollection = lists.GetItems(query);

                // Load only the necessary values for the SharePoint ListItem
                context.Load(
                    itemCollection,
                    items => items.Include(
                    item => item["Title"],
                    item => item[_supportStaffField],
                    item => item[_businessUnitField],
                    item => item[_environmentField],
                    item => item[_serverTypeField],
                    item => item[_serverDescriptionField],
                    item => item.FieldValuesAsText[_serverDescriptionField]));

                context.ExecuteQuery();
            }
            catch (Microsoft.SharePoint.Client.ClientRequestException ex)
            {
                Console.WriteLine("Error Connecting to URL: " + ex.Message);
            }

            return itemCollection;
        }

        // Creates a Dictionary of DGS Servers, using the server name as the key
        private static ServerCollection CreateServerItemList()
        {
            ServerCollection itemDic = new ServerCollection();

            // Get the Server Inventory as a ListItemCollection
            ListItemCollection list = GetServerListItems();
            // Itterate through each item in the list
            foreach (ListItem lItem in list)
            {
                try
                {
                    ServerItem serv = new ServerItem(lItem["Title"].ToString().ToLower());

                    // Get each user in the "Support Staff" field
                    var str = (FieldUserValue[])lItem.FieldValues[_supportStaffField];
                    // Add support staff if it exists
                    if (str != null)
                    {
                        foreach (FieldUserValue user in str)
                        {
                            serv.AddStaff(user.LookupValue);
                        }
                    }

                    // Set the simple text fields
                    serv.BusinessUnit = lItem[_businessUnitField] != null ? (string)lItem[_businessUnitField] : "";
                    serv.ServerType = lItem[_serverTypeField] != null ? (string)lItem[_serverTypeField] : "";
                    serv.Environment = lItem[_environmentField] != null ? (string)lItem[_environmentField] : "";

                    // Set the FieldNote value for Server Description
                    string description = lItem[_serverDescriptionField] != null ? lItem.FieldValuesAsText[_serverDescriptionField] : "";

                    // Remove non-ASCII characters from the string
                    description = description.Replace("\n", "");
                    description = Regex.Replace(description, @"[^\u0000-\u007F]", string.Empty);

                    serv.ServerDescription = description;

                    // Add server to the Dictionary
                    if (!itemDic.ContainsKey(serv.ServerName))
                        itemDic.Add(serv);

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
                
            return itemDic;
        }
        #endregion
    }
}
