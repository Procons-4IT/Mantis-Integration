using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.ServiceProcess;

static class Program
{
    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    static void Main()
    {
        ServiceBase[] ServicesToRun;
        ServicesToRun = new ServiceBase[] 
			{ 
				new Bayan_Service.Bayan_Service()
			};
        ServiceBase.Run(ServicesToRun);
    }
}