using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using PowerShellFunctions;
using System.Collections.ObjectModel;

namespace CertificateLibraries
{
    public static class Certificates
    {
        /// <summary>
        /// Return an X509Certificate that exists on the specified file path.
        /// </summary>
        /// <param name="filepath">Location of the certificate.</param>
        /// <returns>X509Certificate stored in the file.</returns>
        public static X509Certificate2 ImportCertificate(string filepath)
        {
            X509Certificate2 cert = null;
            try
            {
                cert = new X509Certificate2();
                cert.Import(filepath);
            }
            catch (CryptographicException ex)
            {
                Console.WriteLine(ex.Message);
            }

            return cert;
        }

        /// <summary>
        /// Returns an X509Certificate2 based off the rawdata given.
        /// </summary>
        /// <param name="rawdata">Raw data of the certificate.</param>
        /// <returns>X509Certificate from the raw data.</returns>
        public static X509Certificate2 ImportCertificate(byte[] rawdata)
        {
            X509Certificate2 cert = null;
            try
            {
                cert = new X509Certificate2();
                cert.Import(rawdata);
            }
            catch (CryptographicException ex)
            {
                Console.WriteLine(ex.Message);
            }

            return cert;
        }
    }

    namespace CertificateScanning
    {
        /// <summary>
        /// Represents a scanned IP address. Contains a dictionary of all ports that have
        /// certificates on them.
        /// </summary>
        public class IPScanObject
        {
            private string _ipAddress;
            private int _numCertificates;
            private Dictionary<int, X509Certificate2> _certificates;

            public IPScanObject(string ipAdd)
            {
                _ipAddress = ipAdd;
                _numCertificates = 0;
                _certificates = new Dictionary<int, X509Certificate2>();
            }

            public string IPAddress { set { _ipAddress = value; } get { return _ipAddress; } }
            public int NumCertificates { get { return _numCertificates; } }
            public Dictionary<int, X509Certificate2> Certificates { get { return _certificates; } }

            /// <summary>
            /// Adds the certificate on the port to the Dictionary.
            /// </summary>
            /// <param name="port">Port number the certificate resides on.</param>
            /// <param name="cert">Certificate bound to the port.</param>
            public void AddCertificate(int port, X509Certificate2 cert)
            {
                _certificates.Add(port, cert);
                _numCertificates++;
            }

            /// <summary>
            /// Removes the entry in the Dictionary with the matching port.
            /// </summary>
            /// <param name="port">Port to remove from the Dictionary.</param>
            /// <returns>Returns true if the entry was successfully removed.</returns>
            public bool RemoveCertificate(int port)
            {
                if (_certificates.Remove(port))
                {
                    _numCertificates--;
                    return true;
                }
                else { return false; }
            }

            public override string ToString()
            {
                string str = IPAddress + ":\n";
                foreach (KeyValuePair<int, X509Certificate2> cert in Certificates)
                    str += $"\tPort {cert.Key}  :  " + cert.Value.Subject + "\n";
                return str;
            }
        }

        /// <summary>
        /// Represents a server machine, which can host multiple IP addresses with 
        /// many certificates bound to different ports.
        /// </summary>
        public class Server
        {
            private string _machineName;
            private Dictionary<string, IPScanObject> _scannedAddresses;

            public string MachineName { set { _machineName = value; } get { return _machineName; } }

            public Server(string serverName)
            {
                _machineName = serverName;
                _scannedAddresses = new Dictionary<string, IPScanObject>();
            }

            /// <summary>
            /// Returns a Dictionary of all scanned IP addresses
            /// </summary>
            public Dictionary<string, IPScanObject> ScannedAddresses
            {
                get { return _scannedAddresses; }
            }

            /// <summary>
            ///  Adds a scanned IP address to the dictionary
            /// </summary>
            /// <param name="ipScan">The IPScanObject being added</param>
            public void Add(IPScanObject ipScan)
            {
                _scannedAddresses.Add(ipScan.IPAddress, ipScan);
            }

            public override string ToString()
            {
                string str = MachineName + ":\n";
                foreach (KeyValuePair<string, IPScanObject> ip in _scannedAddresses)
                    str += ip.Value.ToString();
                return str;
            }
        }

        /// <summary>
        /// A collection of all servers that have been scanned in the current instance.
        /// </summary>
        public class ServerList : ObservableCollection<Server>
        {
            /// <summary>
            /// Adds a Server to the collection.
            /// </summary>
            /// <param name="server">Server being added to the collection.</param>
            public void AddServer(Server server)
            {
                Add(server);
            }

            public override string ToString()
            {
                string str = "";
                foreach(Server server in Items)
                {
                    str += server;
                }
                return str;
            }
        }

        /// <summary>
        /// Scans for certificates on a remote server using PowerShell remoting.
        /// </summary>
        public class CertificateScanner
        {
            private int _timeout;
            private ServerList _scannedServerList;

            /// <summary>
            /// Creates a CertificateScanner with a specified timeout for scanning.
            /// </summary>
            /// <param name="timeout">Amount of time in milliseconds before cancelling scan.</param>
            public CertificateScanner(int timeout)
            {
                _timeout = timeout;
                _scannedServerList = new ServerList();
            }

            public CertificateScanner()
            {
                _timeout = 2000;
                _scannedServerList = new ServerList();
            }

            /// <summary>
            /// Returns the list of Servers that have been scanned.
            /// </summary>
            public ServerList ScannedServerList { get { return _scannedServerList; } }

            /// <summary>
            /// Runs a certificate scan on a server based on the server name.
            /// </summary>
            /// <param name="serverName">Name of the server being scanned.</param>
            public void RunCertificateScan(string serverName)
            {
                // Initialize a new DGSServer based off the inputted name
                Server server = new Server(serverName);
                // Gather all IPAddresses on the server and scan for certificates on each
                var ipconfig = IPInformationFunctions.InvokeGetIPConfigAllIPv4(serverName);
                foreach (string ip in ipconfig)
                {
                    // Create a new IPScanObject to store certificate information relevant to the server
                    IPScanObject ipScan = new IPScanObject(ip);
                    var listeningPorts = IPInformationFunctions.InvokeGetAllListeningPorts(serverName);
                    foreach (string port in listeningPorts)
                    {
                        // Scan the port for a certificate, add it if it exists
                        var cert = PowerShellCertificateScanner.GetCertificateOnIP(ip, Convert.ToInt32(port));
                        if (cert != null)
                        {
                            ipScan.AddCertificate(Convert.ToInt32(port), cert);
                        }
                    }
                    // Add IPScanObject to scanned collection
                    server.Add(ipScan);
                }
                // Add the server to the server list
                ScannedServerList.Add(server);
            }
        }
    }
}
