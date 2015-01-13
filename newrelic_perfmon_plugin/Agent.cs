using System;
using System.Collections.Generic;
using System.Security;
using System.Security.Principal;
using System.Text.RegularExpressions;
using NewRelic.Platform.Sdk;
using NewRelic.Platform.Sdk.Utils;
using System.Management;
using System.Configuration;

namespace newrelic_perfmon_plugin
{
    class PerfmonAgent : Agent
    {
        private const string DefaultGuid = "com.automatedops.perfmon_plugin";
        private static readonly Logger _logger = Logger.GetLogger("newrelic_perfmon_plugin");
        private static readonly Regex _regex = new Regex("-ap \"(?<poolName>\\w+)\"", RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
 
        private readonly string _name;
        private readonly List<Object> _counters;
        private readonly ManagementScope _scope;

        public override string Guid
        {
            get
            {
                if (ConfigurationManager.AppSettings.HasKeys())
                {
                    if (! string.IsNullOrEmpty(ConfigurationManager.AppSettings["guid"]))
                    {
                        return ConfigurationManager.AppSettings["guid"];
                    }
                }
                return DefaultGuid;
            }
        }

        public override string Version
        {
            get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(); }
        }
       
        public PerfmonAgent(string name, List<Object> paths)
        {
            _name = name;
            _counters = paths;
            _scope = new ManagementScope("\\\\" + _name + "\\root\\cimv2");
#if (DEBUG)
            var domain = name.Substring(0, name.IndexOf("-", StringComparison.InvariantCultureIgnoreCase));
            var username = WindowsIdentity.GetCurrent().Name;
            username = domain + username.Substring(username.IndexOf("\\", StringComparison.InvariantCultureIgnoreCase));
            _scope.Options.Username = username;
#endif
        }

        public override string GetAgentName()
        {
            return _name;
        }

        public override void PollCycle()
        {
            try
            {
                _scope.Connect();

                var processInstances = new Dictionary<string, int>();       // key = w3wp instance, value = pid
                var applicationPools = CreateApplicationPoolDictionary();   // key = pid, value = application pool

                foreach (Dictionary<string, Object> counter in _counters)
                {
                    var providerName = counter["provider"].ToString();
                    var categoryName = counter["category"].ToString();
                    var counterName = counter["counter"].ToString();
                    var predicate = string.Empty;
                    if (counter.ContainsKey("instance"))
                    {
                        predicate = string.Format(" Where Name Like '{0}'", counter["instance"]);
                    }
                    var unitValue = counter["unit"].ToString();

                    var queryString = string.Format("Select Name, {2} from Win32_PerfFormattedData_{0}_{1}{3}", providerName, categoryName, counterName, predicate);
                
                    var search = new ManagementObjectSearcher(_scope, new ObjectQuery(queryString));

                    try
                    {
                        var queryResults = search.Get();

                        foreach (ManagementObject result in queryResults)
                        {
                            try
                            {
                                var value = Convert.ToSingle(result[counterName]);

                                if (null == result["Name"])
                                {
                                    _logger.Warn("Query returned null Name: {0}", queryString);
                                    continue;
                                }

                                //Process ID is a meta-counter which will be stored separately to help derive the application pool
                                if (counterName.Equals("ProcessID", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    processInstances[result["Name"].ToString()] = (int) value;
                                    _logger.Debug("ProcessID {0:N0} stored for instance {1}", value, result["Name"]);
                                    continue;
                                }

                                int pid;
                                string applicationPool;
                                var instanceName = result["Name"].ToString();
                                // attempt to use the application pool as the instance name
                                if (processInstances.TryGetValue(instanceName, out pid) && applicationPools.TryGetValue(pid, out applicationPool))
                                {
                                    instanceName = applicationPool;
                                }

                                var metricName = string.Format("{0}({2})/{1}", categoryName, counterName, instanceName);

                                _logger.Debug("{0}/{1}: {2} {3}", _name, metricName, value, unitValue);

                                ReportMetric(metricName, unitValue, value);
                            }
                            catch (Exception e)
                            {
                                _logger.Error("Exception occurred in processing results. {0}\r\n{1}", e.Message, e.StackTrace);
                            }
                        }
                    }
                    catch (ManagementException e)
                    {
                        _logger.Error("Exception occurred in polling. {0}\r\n{1}", e.Message, queryString);
                    }
                    catch (Exception e)
                    {
                        _logger.Error("Unable to connect to \"{0}\". {1}", _name, e.Message);
                    }     
                } 
         
            }    
            catch (Exception e)
            {
                _logger.Error("Unable to connect to \"{0}\". {1}", _name, e.Message);
            }
        }

        /// <summary>
        /// Queries running processes for all instances of w3wp to create a mapping of pid to application pool
        /// </summary>
        /// <returns>Returns a Dictionary with pid as key and application pool as value.</returns>
        private IDictionary<int, string> CreateApplicationPoolDictionary()
        {
            var pools = new Dictionary<int, string>();
            using (var searcher = new ManagementObjectSearcher(_scope, new SelectQuery("SELECT * FROM Win32_Process where Name = 'w3wp.exe'")))
            {
                foreach (ManagementObject process in searcher.Get())
                {
                    var pid = int.Parse(process["ProcessId"].ToString());
                    var poolName = _regex.Match(process["CommandLine"].ToString()).Groups["poolName"].Value;
                    pools.Add(pid, poolName);
                    _logger.Debug("Application Pool {0} stored with ProcessID {1}", pools[pid], pid);
                }
            }

            return pools;
        }
    }

    class PerfmonAgentFactory : AgentFactory
    {
        public override Agent CreateAgentWithConfiguration(IDictionary<string, object> properties)
        {
            var name = (string)properties["name"];
            var counterlist = (List<Object>)properties["counterlist"];

            if (counterlist.Count == 0)
            {
                throw new Exception("'counterlist' is empty. Do you have a 'config/plugin.json' file?");
            }

            return new PerfmonAgent(name, counterlist);
        }
    }
}
