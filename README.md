# DGSLibraries
Here is a collection of the libraries I made while working at DGS.

Project:
 Certificate Scanner. The objective was to scan SSL and other certificates
 on remote servers throughout the intranet. This process involved:
  - Getting a list of IP addresses associated with each server
  - Getting a list of active and listening ports for each server
  - Scanning for a Certificate on that port
  - Gathering and storing the certificate information
  
Goal: of the project was to discover certificates that were nearly
 expired, and to contact the person responsible for that server to notify
 them.
 
Process:
 Remoting into the servers was require because many of the servers were 
 multi-homed, meaning they have many IP Addresses hosted on the server and
 many of them were not registered with the DNS (making it impossible to find
 this information without a direct connection). To get the list of IPs per
 server, remote into the server, get "ipconfig /all", and parse the returned
 information into a list of IP addresses as strings. This same process was
 used for gathering a list of active ports per server. 
 
 The program was initially created and run in PowerShell (which I will include), 
 but I created a version written in C# to expirement with time-efficiency. 
 The SharePoint libraries were created in C# due to the 
 Microsoft.SharePoint.Client libraries. This library was used to pull a server's
 information from the SharePoint repository and find the support contact for
 that server (so that if a certificate is near expiration we could 
 automatically send them an email reminding them to renew it).
 
Libraries:
  PowerShellFuncions:
   A collection of classes that allow you to run PowerShell script from C#. The
   general class allows you to run code and return a Collection<PSObject>, the
   children classes are specific for the purpose of the Certificate Scanner.

  CertificateLibraries:
   Contains the certificate scanner. The certificate scanner takes a machine
   name, and returns a structured object containing all IP Addresses with all
   Ports that have a certificate bound to it, as well as the certificate on each 
   port.

  SharePointLibraries:
   A Collection of SharePoint functions that can be used to query a SharePoint
   website. Contains methods that allow you to retrieve all Lists associated
   with a SharePoint website, all Field items associated with a List, and all
   ListItems in a List.
   There are also more specialized classes meant for DGS use only. The 
   DGSServerInventory automatically creates a structured object that represents
   the Server repository (with just the server name and the support staff, because
   that is all we were interested in at the time).
 
