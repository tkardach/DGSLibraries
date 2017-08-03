using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security.Cryptography.X509Certificates;

namespace PowerShellFunctions
{
    public class PowerShellFunctions
    {
        // Convert a Dictionary of parameters into a PowerShell argument list string
        private static string ConvertParameters(Dictionary<string, Object> parameters)
        {
            if (parameters == null) return "";
            string result = "";
            // Itterate through each parameter and add it using PowerShell syntax
            foreach (KeyValuePair<string, Object> param in parameters)
                result += $" -{param.Key} {param.Value}";

            return result;
        }

        #region Methods
        /// <summary>
        /// Returns the result of running the parameter "command" against the PowerShell
        /// script "script".
        /// </summary>
        /// <param name="script">Text version of the PowerShell script being run.</param>
        /// <param name="command">Command being run on the PowerShell script</param>
        /// <returns>
        /// A Collection of PSObjects containing the result of running the command against
        /// the script.
        /// </returns>
        public static Collection<PSObject> RunScript(string command, string script = "")
        {
            InitialSessionState initial = InitialSessionState.CreateDefault();
            Runspace runspace = RunspaceFactory.CreateRunspace(initial);
            Collection<PSObject> result = null;
            try
            {
                runspace.Open();
                // Create a PowerShell object to manage the runspace
                PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;
                // Add parameters to runspace
                ps.AddScript(script);
                ps.AddScript(command);
                // Invoke the command
                result = ps.Invoke();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                runspace.Close();
            }
            return result;
        }

        /// <summary>
        /// Returns the result of running the specified command with parameters on the given script.
        /// </summary>
        /// <param name="command">Command being executed on the script.</param>
        /// <param name="parameters">Parameters being passed to the command.</param>
        /// <param name="script">Script being used.</param>
        /// <returns>A collection of the results returned from running the command on the given script.</returns>
        public static Collection<PSObject> RunScript(string command, Dictionary<string, Object> parameters, string script = "")
        {
            return RunScript(command + ConvertParameters(parameters), script);
        }

        /// <summary>
        /// Invokes the script block on the specified remote machine.
        /// </summary>
        /// <param name="machineName">Name of the remote machine where the command will be run.</param>
        /// <param name="scriptBlock">Content of the script being run on the remote machine.</param>
        /// <returns>A collection of the results returned from running the script on the remote machine.</returns>
        public static Collection<PSObject> InvokeCommand(string machineName, string scriptBlock)
        {
            Dictionary<string, Object> parameters = new Dictionary<string, Object>();
            // Create the script block and add it as a parameter
            ScriptBlock filter = ScriptBlock.Create(scriptBlock);
            parameters.Add("ScriptBlock", filter);
            // Add the ComputerName parameter
            parameters.Add("ComputerName", machineName);
            // Call Invoke-Command
            return RunCommand("Invoke-Command", parameters);
        }

        /// <summary>
        /// Returns the result of running the specified command with optional parameters.
        /// </summary>
        /// <param name="command">Command being executed.</param>
        /// <param name="parameters">Optional parameters being sent to the command.</param>
        /// <returns>A collection of the results returned from executing the command.</returns>
        public static Collection<PSObject> RunCommand(string command, Dictionary<string, Object> parameters = null)
        {
            InitialSessionState initial = InitialSessionState.CreateDefault();
            Runspace runspace = RunspaceFactory.CreateRunspace(initial);
            Collection<PSObject> result = null;
            try
            {
                runspace.Open();
                // Create a PowerShell object to manage the runspace
                PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;
                // Add parameters to runspace
                ps.AddCommand(command);
                ps.AddParameters(parameters);
                // Invoke the command
                result = ps.Invoke();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                runspace.Close();
            }
            return result;
        }
        #endregion
    }

    public class IPInformationFunctions : PowerShellFunctions
    {
        // No error checking because we want the program to crash if these initializations don't work
        private static string _scriptPath =
            AppDomain.CurrentDomain.BaseDirectory + @"\Script\IPInformation.ps1";
        private static string _scriptText = File.ReadAllText(_scriptPath);

        private static string _ipConfigCommand = "Get-IPConfigAllIPv4";
        private static string _netstatCommand = "Get-NetstatInformation";

        #region Accessors
        public static string ScriptPath { get { return _scriptPath; } }
        public static string ScriptText { get { return _scriptText; } }
        public static string IPConfigCommand { get { return _ipConfigCommand; } }
        public static string NetstatCommand { get { return _netstatCommand; } }
        #endregion

        #region Methods
        /// <summary>
        /// Returns all IPv4 IP addresses in the form of an array.
        /// </summary>
        /// <returns>
        /// Returns a Collection of PSObjects containing each IPv4 IP address
        /// on the local machine.
        /// </returns>
        public static List<string> GetIPConfigAllIPv4 ()
        {
            List<string> ips = new List<string>();
            // Create the proper PowerShell command and run the script
            string totalScript = _scriptText + " " + _ipConfigCommand;
            var result = RunScript(_ipConfigCommand, _scriptText);
            // Add each IP results to the List
            foreach (PSObject ip in result)
            {
                ips.Add(ip.Properties["Server"].Value.ToString());
            }

            return ips;
        }

        /// <summary>
        /// Returns all active ports on the local machine and their status.
        /// </summary>
        /// <returns>
        /// Returns a Collection of PSObjects containing each active port
        /// number and their status.
        /// </returns>
        public static List<string[]> GetNetstatInformation()
        {
            List<string[]> ports = new List<string[]>();
            // Create the proper PowerShell command and run the script
            string totalScript = _scriptText + " " + _netstatCommand;
            var result = RunScript(_netstatCommand, _scriptText);
            // Add each Port results to the List
            foreach (PSObject port in result)
            {
                String[] portInfo = new String[2];
                portInfo[0] = port.Properties["Port"].Value.ToString();
                portInfo[1] = port.Properties["Status"].Value.ToString();
                ports.Add(portInfo);
            }

            return ports;
        }


        /// <summary>
        /// Returns all IPv4 IP addresses on a remote machine in the form of a collection.
        /// </summary>
        /// <returns>
        /// Returns a Collection of PSObjects containing each IPv4 IP address
        /// on the local machine.
        /// </returns>
        public static List<string> InvokeGetIPConfigAllIPv4(string machineName)
        {
            List<string> ips = new List<string>();
            // Create the proper PowerShell command and get the results
            string totalScript = _scriptText + " " + _ipConfigCommand;
            var result = InvokeCommand(machineName, totalScript);
            // Add each IP from the results to the List
            foreach (PSObject ip in result)
            {
                ips.Add(ip.Properties["Server"].Value.ToString());
            }

            return ips;
        }

        /// <summary>
        /// Returns all active ports on the remote machine and their status.
        /// </summary>
        /// <returns>
        /// Returns a Collection of PSObjects containing each active port
        /// number and their status.
        /// </returns>
        public static List<string[]> InvokeGetNetstatInformation(string machineName)
        {
            List<string[]> ports = new List<string[]>();
            // Create the proper PowerShell command and run the script
            string totalScript = _scriptText + " " + _netstatCommand;
            var result = InvokeCommand(machineName, totalScript);
            // Add each Port results to the List
            foreach (PSObject port in result)
            {
                String[] portInfo = new String[2];
                portInfo[0] = port.Properties["Port"].Value.ToString();
                portInfo[1] = port.Properties["Status"].Value.ToString();
                ports.Add(portInfo);
            }

            return ports;
        }

        // Returns the active ports that meet the specified status
        private static List<string> GetPortByStatus(string status)
        {
            List<string> ports = new List<string>();
            // Create the proper PowerShell command and run the script
            string ext = $" | where {{$_.status -eq \"{status}\"}}";
            var result = RunScript(_netstatCommand + ext, _scriptText);
            // Add each Port results to the List
            foreach (PSObject port in result)
            {
                ports.Add(port.Properties["Port"].Value.ToString());
            }

            return ports;
        }

        /// <summary>
        /// Returns all active ports that have an Established status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with an Established status
        /// </returns>
        public static List<string> GetAllEstablishedPorts()
        {
            return GetPortByStatus("ESTABLISHED");
        }

        /// <summary>
        /// Returns all active ports that have a Listening status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with a Listening status
        /// </returns>
        public static List<string> GetAllListeningPorts()
        {
            return GetPortByStatus("LISTENING");
        }

        /// <summary>
        /// Returns all active ports that have a Time-Wait status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with a Time-Wait status
        /// </returns>
        public static List<string> GetAllTimeWaitPorts()
        {
            return GetPortByStatus("TIME-WAIT");
        }

        /// <summary>
        /// Returns all active ports that have a Close-Wait status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with a Close-Wait status
        /// </returns>
        public static List<string> GetAllCloseWaitPorts()
        {
            return GetPortByStatus("CLOSE-WAIT");
        }

        // Returns all the active ports on the remote machine that meet the specified status
        private static List<string> InvokeGetPortByStatus(string machineName, string status)
        {
            List<string> ports = new List<string>();
            // Create the proper PowerShell command and run the script
            string ext = $" | where {{$_.status -eq \"{status}\"}}";
            string totalScript = _scriptText + Environment.NewLine + _netstatCommand + ext;
            var result = InvokeCommand(machineName, totalScript);
            // Add each IP results to the List
            foreach (PSObject port in result)
            {
                ports.Add(port.Properties["Port"].Value.ToString());
            }

            return ports;
        }

        /// <summary>
        /// Returns all active ports on the specified machine that have an Established status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with an Established status
        /// </returns>
        public static List<string> InvokeGetAllEstablishedPorts(string machineName)
        {
            return InvokeGetPortByStatus(machineName, "ESTABLISHED");
        }

        /// <summary>
        /// Returns all active ports on the specified machine that have a Listening status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with a Listening status
        /// </returns>
        public static List<string> InvokeGetAllListeningPorts(string machineName)
        {
            return InvokeGetPortByStatus(machineName, "LISTENING");
        }

        /// <summary>
        /// Returns all active ports on the specified machine that have a Time-Wait status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with a Time-Wait status
        /// </returns>
        public static List<string> InvokeGetAllTimeWaitPorts(string machineName)
        {
            return InvokeGetPortByStatus(machineName, "TIME-WAIT");
        }

        /// <summary>
        /// Returns all active ports on the specified machine that have a Close-Wait status
        /// </summary>
        /// <returns>
        /// Collection of PSObjects containing every port with a Close-Wait status
        /// </returns>
        public static List<string> InvokeGetAllCloseWaitPorts(string machineName)
        {
            return InvokeGetPortByStatus(machineName, "CLOSE-WAIT");
        }
        #endregion
    }

    public class ActiveDirectoryFunctions : PowerShellFunctions
    {
        // No error checking because we want the program to crash if these initializations don't work
        private static string _scriptPath =
            AppDomain.CurrentDomain.BaseDirectory + @"\Script\MSOffice.ps1";
        private static string _scriptText = File.ReadAllText(_scriptPath);
        
        private const string _getADUserEmail = "Get-ADUserEmail";

        public static string ScriptText { get { return _scriptText; } }
        public static string ScriptPath { get { return _scriptPath; } }
        public static string GetADUserEmailCommand { get { return _getADUserEmail; } }
        
        /// <summary>
        /// Takes an Active Directory domain and a username for that domain and returns their email address.
        /// </summary>
        /// <param name="domain">Domain name for the Active Directory server being queried.</param>
        /// <param name="user">Active Directory Username that will be used to find their email.</param>
        /// <returns>Collection of PSObjects of emails that were found for the user.</returns>
        public static string GetActiveDirectoryUserEmail(string domain, string user)
        {
            string email = "";
            // Create the proper PowerShell command and run the script
            string ext = $" -Domain \"{domain}\" -CommonName \"{user}\"";
            var result = RunScript(_getADUserEmail + ext, _scriptText);
            foreach (PSObject ps in result)
            {
                email = ps.Properties["email"].Value.ToString();
            }

            return email;
        }
    }

    public class DGSActiveDirectoryFunctions : ActiveDirectoryFunctions
    {
        private static string _dgsDomain = "/*DGS Active Directory Domain removed for public repository*/";

        /// <summary>
        /// Takes a DGS Active Directory Username in the proper format {last, first@DGS} and returns their email.
        /// </summary>
        /// <param name="user">DGS Active Directory Username that will be used to find their email.</param>
        /// <returns>Collection of PSObjects of emails that were found for the user.</returns>
        public static string GetActiveDirectoryUserEmail(string user)
        {
            return GetActiveDirectoryUserEmail(_dgsDomain, user);
        }
    }

    public class PowerShellCertificateScanner : PowerShellFunctions
    {
        // No error checking because we want the program to crash if these initializations don't work
        private static string _scriptPath =
             AppDomain.CurrentDomain.BaseDirectory + @"\Script\CertificateScanner.ps1";
        private static string _scriptText = File.ReadAllText(_scriptPath);

        private static string _getCertificate = "Get-CertificateInformation";

        public static string ScriptPath { get { return _scriptPath; } }
        public static string ScriptText { get { return _scriptText; } }
        public static string GetCertificate { get { return _getCertificate; } }

        /// <summary>
        /// Creates an X509Certificate2 based of the certificate at the given IP and Port.
        /// </summary>
        /// <param name="ipAddress">IP Address of the machine being scannned for certificates.</param>
        /// <param name="port">Port number being checked for bound certificates.</param>
        /// <param name="timeout">Number of milliseconds to wait before giving up on SSL/TLS negotiation</param>
        /// <returns>Returns the X509Certificate2 for the given IP and Port.</returns>
        public static X509Certificate2 GetCertificateOnIP (string ipAddress, int port, int timeout = 2000)
        {
            // Get the RawData of the certificate from the ipaddress:port
            Collection<PSObject> result = GetRawCertificateOnIP(ipAddress, port, timeout);
            X509Certificate2 certificate = null;

            // Itterate through the results to extract the RawData attribute from the PSObject
            foreach (PSObject psobject in result)
            {
                // If the PSObject is null, no certificate was found. Return null
                if (psobject == null) return null;
                foreach (PSMemberInfo info in psobject.Members)
                {
                    if (info.Name == "RawData")
                    {
                        // Create the certificate based off the RawData returned
                        certificate = new X509Certificate2((byte[])info.Value);
                    }
                    else if (info.Name == "Error")
                    {
                        Console.WriteLine(info.Value);
                    }
                }
            }

            if (certificate != null) { return certificate; }
            else { return null; }
        }

        // Runs the PowerShell command for getting the RawData of a certificate on a specified IP and Port
        private static Collection<PSObject> GetRawCertificateOnIP (string ipAddress, int port, int timeout)
        {
            string ext = $" -MachineName \"{ipAddress}\" -Port {port} -Timeout {timeout}";
            return RunScript(_getCertificate + ext, _scriptText);
        } 
    }
}
