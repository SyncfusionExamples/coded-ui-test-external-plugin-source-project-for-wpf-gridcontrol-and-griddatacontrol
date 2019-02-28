namespace Syncfusion.UITest.GridCommunication
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Windows;
    using System.Runtime.Remoting.Channels;
    using System.Runtime.Remoting;
    using System.Runtime.Remoting.Channels.Ipc;

    public class GridControlTestApplication : Application
    {
        private IChannel channel;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            this.channel = new IpcChannel("GridControl");
            ChannelServices.RegisterChannel(this.channel, false);
            RemotingConfiguration.RegisterWellKnownServiceType(typeof(GridInteropService), "GridTestService", WellKnownObjectMode.Singleton);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            if (this.channel != null)
            {
                ChannelServices.UnregisterChannel(this.channel);
            }
        }
    }
}
