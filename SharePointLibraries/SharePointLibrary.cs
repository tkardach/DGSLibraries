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
        private const string _dgsSharePointUrl = "http://dgssp.dgs.ca.gov/projects/ETS/ETSDocumentation/";

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

    public class DGSServerInventory : DGSSharePoint
    {
        private const string _serverInventoryList = "ETS Server Inventory";
        private const string _supportStaffField = "Support_x0020_Staff_x0020_2";
        //private const string _applicationNameField = "Application_x0020_Name";
        //private const string _businessUnitField = "Support_x0200_Unit";
        //private const string _serverTypeField = "Server_x0020_Type";
        //private const string _environmentField = "Physical_x002f_Virtual";

        private Dictionary<string, ServerItem> _serverInventory;

        // Represents a Server in the Server Inventory
        public class ServerItem
        {
            private String _serverName;
            private List<String> _supportStaff;
            //private string _applicationName;
            //private string _businessUnit;
            //private string _serverType;
            //private string _environment;

            // Constructs the Server Item based off of the server name
            public ServerItem(String name)
            {
                _serverName = name;
                _supportStaff = new List<string>();
            }

            // Accessor methods
            public String ServerName { set { _serverName = value; } get { return _serverName; } }
            public List<String> SupportStaff { set { _supportStaff = value; } get { return _supportStaff; } }
            //public string ApplicationName { set { _applicationName = value; } get { return _applicationName; } }
            //public string BusinessUnit { set { _businessUnit = value; } get { return _businessUnit; } }
            //public string ServerType { set { _serverType = value; } get { return _serverType; } }
            //public string Environment { set { _environment = value; } get { return _environment; } }

            // Support staff List methods
            public void AddStaff(string staff) { _supportStaff.Add(staff); }
            public bool RemoveStaff(string staff) { return _supportStaff.Remove(staff); }

            // Override ToString
            public override string ToString()
            {
                return _serverName;
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
        public static string ServerInventoryList { get { return _serverInventoryList; } }
        // Accessor for the Server Inventory Dictionary
        public Dictionary<string, ServerItem> ServerInventory { get { return _serverInventory; } }

        // Returns the ListItemCollection containing the DGS Server Inventory items
        private static ListItemCollection GetServerListItems()
        {
            return GetListItems(ServerInventoryList);
        }


        // Creates a Dictionary of DGS Servers, using the server name as the key
        private static Dictionary<string, ServerItem> CreateServerItemList()
        {
            Dictionary<string, ServerItem> itemDic = new Dictionary<string, ServerItem>();

            // Get the Server Inventory as a ListItemCollection
            ListItemCollection list = GetServerListItems();
            
            // Itterate through each item in the list
            foreach (ListItem lItem in list)
            {
                ServerItem serv = new ServerItem(lItem["Title"].ToString());

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
                // Add server to the Dictionary
                if (!itemDic.ContainsKey(serv.ServerName))
                    itemDic.Add(serv.ServerName, serv);
            }

            return itemDic;
        }
        #endregion
    }
    
    public class DGSApplicationInventory : DGSSharePoint
    {
        private const String ApplicationInventoryList = "ETS Application Inventory";

        // Constructor initializes using DGS SharePoint URL and creates 
        public DGSApplicationInventory()
        {
        }
    }
}
