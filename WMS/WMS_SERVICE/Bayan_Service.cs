using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.IO;
using System.Timers;
using System.Configuration;

namespace Bayan_Service
{
    partial class Bayan_Service : ServiceBase
    {

        System.Timers.Timer tmrSync = new System.Timers.Timer();

        private bool blnInProcess = false;
        private BAYAN_LIB.clsLibrary objclsLibrary = null;
        private string[] strTime;

        public Bayan_Service()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                traceService("Bayan Service Started At : " + DateTime.Now);
                tmrSync.Elapsed += new System.Timers.ElapsedEventHandler(tmrSync_Elapsed);
                tmrSync.Enabled = true;
                int intSyncInt = Convert.ToInt32(ConfigurationManager.AppSettings["SynInterval"].ToString());
                tmrSync.Interval = intSyncInt;
            }
            catch (Exception ex)
            {
                traceService(ex.Message.ToString() + DateTime.Now);
            }
        }

        protected override void OnStop()
        {
            try
            {

            }
            catch (Exception ex)
            {
                traceService(ex.StackTrace.ToString());
            }
        }

        protected override void OnPause()
        {
            try
            {
                traceService("Service Pause :" + DateTime.Now);
                tmrSync.Stop();
            }
            catch (Exception ex)
            {
                traceService(ex.StackTrace.ToString());
            }
        }

        protected override void OnContinue()
        {
            try
            {
                traceService("Service Continues :" + DateTime.Now);
                tmrSync.Start();
            }
            catch (Exception ex)
            {
                traceService(ex.StackTrace.ToString());
            }
        }

        private void tmrSync_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                traceService("Timer Reset At :" + DateTime.Now);
                if (!blnInProcess)
                //if (1 == 1)
                {
                    blnInProcess = true;
                    traceService("Started Processing Customer Sync Process..." + DateTime.Now);
                    create_Customer_Sync();
                    traceService("Completed Processing Customer Sync Process..." + DateTime.Now);
                    blnInProcess = false;
                }
                else
                {
                    traceService("Still In Process..." + DateTime.Now);
                }
                traceService("Timer Elapses At :" + DateTime.Now);
            }
            catch (Exception ex)
            {
                blnInProcess = false;
                traceService(ex.StackTrace.ToString());
                traceService(ex.Message.ToString());
                traceService(ex.InnerException.ToString());
            }
            finally
            {
                
            }
        }
        
        private void create_Customer_Sync()
        {
            objclsLibrary = new BAYAN_LIB.clsLibrary();
            try
            {
                if (objclsLibrary.mainFunction())
                {
                    traceService("Bayan Sync Success..." + DateTime.Now);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                objclsLibrary = null;
            }
        }

        private void traceService(string content)
        {
            try
            {
                string strFile = @"\BAYAN_SERVICE_" + System.DateTime.Now.ToString("yyyyMMdd") + ".txt";
                string strPath = System.Windows.Forms.Application.StartupPath.ToString() + strFile;

                if (!File.Exists(strPath))
                {
                    FileStream fs = new FileStream(strPath, FileMode.Create, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.BaseStream.Seek(0, SeekOrigin.End);
                    sw.WriteLine(content);
                    sw.Flush();
                    sw.Close();
                }
                else
                {
                    FileStream fs = new FileStream(strPath, FileMode.Append, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.BaseStream.Seek(0, SeekOrigin.End);
                    sw.WriteLine(content);
                    sw.Flush();
                    sw.Close();
                }
            }
            catch (Exception)
            {

            }
        }

    }
}
