using System;
using NewRelic.Platform.Sdk;
using NewRelic.Platform.Sdk.Utils;
using Topshelf;
using System.Threading;

namespace newrelic_perfmon_plugin
{
    class Program
    {
        static void Main(string[] args)
        {
            HostFactory.Run(x =>
            {
                x.Service<PluginService>(sc =>
                {
                    sc.ConstructUsing(() => new PluginService());

                    sc.WhenStarted(s => s.Start());
                    sc.WhenStopped(s => s.Stop());
                });
                x.SetServiceName("newrelic_perfmon_plugin");
                x.SetDisplayName("NewRelic Windows Perfmon Plugin");
                x.SetDescription("Sends Perfmon Metrics to NewRelic Platform");
                x.StartAutomatically();
                x.RunAsPrompt();
            });
        }
    }

    class PluginService
    {
        private Runner _runner;
        private Thread _thread;

        private static readonly Logger logger = Logger.GetLogger("newrelic_perfmon_plugin");

        public PluginService()
        {
            _runner = new Runner();
            _runner.Add(new PerfmonAgentFactory());
        }

        public void Start()
        {
            logger.Info("Starting service.");
            _thread = new Thread(_runner.SetupAndRun);
            try
            {
                _thread.Start();
            }
            catch (Exception e)
            {
                logger.Error("Exception occurred, unable to continue. {0}\r\n{1}", e.Message, e.StackTrace);
            }
        }

        public void Stop()
        {
            logger.Info("Stopping service.");
            Thread.Sleep(5000);

            if (!_thread.IsAlive) return;

            _runner = null;
            _thread.Abort();
        }
    }
}
