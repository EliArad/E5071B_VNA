using Intel_unit_Test;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace E5071BLib

{
    public class E5071BApi
    {

        bool m_abort = false;
        protected string m_staCalFile;
        protected Thread m_waitTaskThread = null;
        public delegate void VNACallback(int code, string message);
        protected Object thisLock = new Object();
        protected Thread m_thread;
        protected EventWaitHandle m_notifyEvent = new AutoResetEvent(false);
        private bool m_pause = false;
        protected string m_errorMessage;
        protected bool m_copyInside = false;
        protected EventWaitHandle m_copyInsideEvent = new AutoResetEvent(false);
        protected string m_destDirectory;
        protected EventWaitHandle m_timer = new AutoResetEvent(false);
        protected string m_firstS4pFile;
        protected string m_lastS4pFile;
        
        protected bool m_error = false;
        protected Queue<string> m_monitorDirQueue = new Queue<string>();
        protected Task<string> m_copyTask = null;
        bool m_running = false;
        ScpiDemo m_scpi = null;
        //ScpiNull m_scpi;
        string m_VNAMapDir;
        protected bool m_continues = false;

        VNACallback m_callback;
        public E5071BApi(VNACallback callback)
        {
            m_callback = callback;
        }
        void Wait()
        {
            string res = string.Empty;
            res = m_scpi.QueryScpi("*OPC?");
            while (res != "+1")
            {
                res = m_scpi.QueryScpi("*OPC?");
            }
        }
        protected void StartWaitTaskThread()
        {
            m_waitTaskThread = new Thread(WaitTask);
            m_waitTaskThread.Start();
        }
        protected void WaitTask()
        {
            if (m_copyInside == true)
            {
                while (m_copyTask == null)
                {
                    Thread.Sleep(100);
                }
                m_copyTask.Wait();
                Stop();
            }
        }


        public string CopyOnStop(string destDirectory, string VNAMapDir)
        {

            if (VNAMapDir[VNAMapDir.Length - 1] != '\\')
                VNAMapDir += "\\";

            ulong fileIndex = 0;
            string sourceFileName, destFileName;

            m_copyTask = new Task<string>(() =>
            {
                try
                {
                    while (true)
                    {
                        if (m_monitorDirQueue.Count == 0)
                            return "ok";

                        sourceFileName = m_monitorDirQueue.Dequeue();

                        destFileName = destDirectory + "\\" + sourceFileName;
                        try
                        {
                            if (File.Exists(destFileName) == true)
                            {
                                File.Delete(destFileName);
                            }
                            int timeOut = 5 * 10;
                            while (File.Exists(VNAMapDir + sourceFileName.ToUpper()) == false)
                            {
                                Thread.Sleep(200);
                                if (timeOut == 0)
                                {
                                    throw (new SystemException("File: " + VNAMapDir + sourceFileName + " is not in VNA yet"));
                                }
                                timeOut--;
                            }
                            //Debug.WriteLine("Start move");
                            File.Move(VNAMapDir + sourceFileName, destFileName);
                            //Console.WriteLine(destFileName);
                            if (m_callback != null)
                                m_callback(868, destFileName);

                            fileIndex++;
                        }
                        catch (Exception err)
                        {
                            m_error = true;
                            m_errorMessage = err.Message;
                            m_notifyEvent.Set();
                            return err.Message;
                        }
                    }
                }
                catch (Exception err)
                {
                    m_error = true;
                    m_errorMessage = err.Message;
                    m_notifyEvent.Set();
                    return err.Message;
                }
            });
            m_copyTask.Start();
            return "ok";
        }
        protected virtual void VNACaptureProcess(TimeSpan sleepTime)
        {
            string fileName;
            uint fileIndex = 0;
            TimeSpan diffTime = new TimeSpan(0, 0, 0);
            while (m_running)
            {
                if (sleepTime > diffTime)
                    sleepTime = sleepTime - diffTime;
                m_timer.WaitOne(sleepTime);
                //Console.WriteLine(sleepTime);
                if (m_running == false)
                    return;

                if (m_pause == true)
                {
                    Thread.Sleep(10);
                    continue;
                }
                DateTime start = DateTime.Now;
                //Console.WriteLine(start.ToString());
                fileName = start.Date.Day.ToString("00") + "_" + start.Date.Month.ToString("00") + "_" + start.Date.Year.ToString("0000") + "_" + start.Hour.ToString("00") + "_" + start.Minute.ToString("00") + "_" + start.Second.ToString("00") + "_" + start.Millisecond.ToString("000") + "_Snap.s4p";

                StartVNAS4PCapture(fileName);
                DateTime end = DateTime.Now;
                diffTime = end - start;

                if (fileIndex == 0)
                {
                    m_firstS4pFile = fileName;
                }
                else
                {
                    m_lastS4pFile = fileName;
                }
                fileIndex++;
                m_monitorDirQueue.Enqueue(fileName);
                if (m_copyInside)
                    m_copyInsideEvent.Set();
            }
        }
        public virtual void Start(string destDirectory,
                                  TimeSpan sleepTime,                                
                                  bool copyInside)
        {

          
            m_firstS4pFile = string.Empty;
            m_lastS4pFile = string.Empty;
            m_errorMessage = string.Empty;
            m_copyInside = copyInside;

            m_monitorDirQueue.Clear();
            m_error = false;
            m_copyInsideEvent = new AutoResetEvent(false);

            m_destDirectory = destDirectory;
            m_thread = new Thread(() => VNACaptureProcess(sleepTime));
            m_running = true;
            m_thread.Start();
            if (m_VNAMapDir[m_VNAMapDir.Length - 1] != '\\')
                m_VNAMapDir += "\\";
            m_copyTask = null;
            StartWaitTaskThread();
            if (copyInside == true)
                CopyVNAFilesInTask(destDirectory + "VNA\\", m_VNAMapDir);
        }


        public virtual void Start(string destDirectory,
                                  string VNAMapDir,
                                  bool copyInside = true)
        {

            m_firstS4pFile = string.Empty;
            m_lastS4pFile = string.Empty;
            m_errorMessage = string.Empty;
            m_copyInside = copyInside;

            m_monitorDirQueue.Clear();
            m_error = false;
            m_copyInsideEvent = new AutoResetEvent(false);

            m_destDirectory = destDirectory;
            m_running = true;
            m_thread.Start();
            if (VNAMapDir[VNAMapDir.Length - 1] != '\\')
                VNAMapDir += "\\";
            m_copyTask = null;
            StartWaitTaskThread();
            if (copyInside == true)
                CopyVNAFilesInTask(destDirectory + "VNA\\", VNAMapDir);
        }
        private string CopyVNAFilesInTask(string destDirectory, string VNAMapDir)
        {
            ulong fileIndex = 0;
            string sourceFileName, destFileName;

            m_copyTask = new Task<string>(() =>
            {
                try
                {
                    while (true)
                    {
                        if (m_monitorDirQueue.Count < 2)
                            m_copyInsideEvent.WaitOne();
                        if (m_monitorDirQueue.Count == 0 && (m_running == false))
                            return "ok";

                        if (m_monitorDirQueue.Count < 2 && m_running == true)
                        {
                            continue;
                        }
                        sourceFileName = m_monitorDirQueue.Dequeue();

                        destFileName = destDirectory + "\\" + sourceFileName;
                        try
                        {
                            if (File.Exists(destFileName) == true)
                            {
                                File.Delete(destFileName);
                            }
                            int timeOut = 5 * 10;
                            while (File.Exists(VNAMapDir + sourceFileName.ToUpper()) == false)
                            {
                                Thread.Sleep(200);
                                if (timeOut == 0)
                                {
                                    throw (new SystemException("File: " + VNAMapDir + sourceFileName + " is not in VNA yet"));
                                }
                                timeOut--;
                            }
                            //Debug.WriteLine("Start move");
                            File.Move(VNAMapDir + sourceFileName, destFileName);
                            //Console.WriteLine(destFileName);

                            if (m_callback != null)
                            {
                                m_callback(868, destFileName);
                                m_callback(860, m_monitorDirQueue.Count.ToString());
                            }
        

                            fileIndex++;
                        }
                        catch (Exception err)
                        {
                            m_error = true;
                            m_errorMessage = err.Message;
                            m_notifyEvent.Set();
                            return err.Message;
                        }
                    }
                }
                catch (Exception err)
                {
                    m_error = true;
                    m_errorMessage = err.Message;
                    m_notifyEvent.Set();
                    return err.Message;
                }
            });
            m_copyTask.Start();
            return "ok";
        }
        public void Close()
        {
            if (m_scpi != null)
                m_scpi.Close();
        }
        public virtual void Stop(bool writeVNASummery = true)
        {
            TimeSpan timeout = new TimeSpan(0, 0, 6);
            lock (thisLock)
            {
                if (m_running == false)
                {
                    if (m_scpi != null)
                        m_scpi.Close();
                    m_scpi = null;
                    return;
                }
                m_running = false;
                m_pause = false;
                m_timer.Set();
                m_thread.Join();
                m_copyInsideEvent.Set();
                if (m_error == false)
                {
                    if (m_copyInside == true && m_copyTask != null)
                    {
                        while (m_copyTask.Wait(timeout) == false)
                        {
                            m_copyInsideEvent.Set();
                        }
                    }
                }
                string p = string.Empty;
                CopySelectedFilesToDestResults(m_destDirectory + "Results_" + p);

                if (m_scpi != null)
                    m_scpi.Close();
                m_scpi = null;
            }
        }
        void CopySelectedFilesToDestResults(string dest)
        {
            string f1 = m_destDirectory + "VNA\\" + m_firstS4pFile;
            if (File.Exists(f1) == true)
                File.Copy(f1, dest + "\\" + m_firstS4pFile);
            string f2 = m_destDirectory + "VNA\\" + m_lastS4pFile;
            if (File.Exists(f2) == true)
                File.Copy(f2, dest + "\\" + m_lastS4pFile);
        }
        public void Initialize(string visaAddress, string STACalFile)
        {
            try
            {
                m_errorMessage = string.Empty;
                m_staCalFile = STACalFile;
                if (visaAddress == null || visaAddress == string.Empty)
                    throw (new SystemException("Invalid vna visa address"));
                m_scpi = new ScpiDemo(visaAddress);


                m_scpi.SendScpi(":STATus:PRESet");
                Wait();
                m_scpi.SendScpi(":MMEMory:LOAD:STATe \"" + STACalFile + "\"");
                m_scpi.SendScpi(":MMEMory:STORe:SNP:TYPE:S1P 1");
                if (m_continues == true)
                {
                    m_scpi.SendScpi(":INITiate:CONTinuous 1");
                }
                else
                {
                    m_scpi.SendScpi(":DISPlay:WINDow:ACTivate");
                    Wait();
                    m_scpi.SendScpi(":TRIGger:SEQuence:SOURce BUS");
                    Wait();
                    m_scpi.SendScpi(":INITiate:CONTinuous 0");
                }
                Wait();
                m_scpi.SendScpi(":DISPlay:ENABle 1");
                Wait();
            }
            catch (Exception err)
            {
                throw (new SystemException(err.Message));
            }
        }

        public virtual void Initialize(string visaAddress, string VNAMapDir, 
                                       string STACalFile)
        {
            try
            {
                string STAOnVna = STACalFile;
                STAOnVna = VNAMapDir.Substring(0, 3) + STACalFile; //returns " Hello World"
                m_VNAMapDir = VNAMapDir;

                /*
                if (File.Exists(STAOnVna) == false)
                {
                    throw (new SystemException("VNA calibration file sensing.sta is missing\nThe file should be located at d:\\sensing\\sensing.sta on the network drive"));
                }
                 */

                m_errorMessage = string.Empty;
                m_staCalFile = STACalFile;
                if (visaAddress == null || visaAddress == string.Empty)
                    throw (new SystemException("Invalid vna visa address"));
                m_scpi = new ScpiDemo(visaAddress);
                //m_scpi = new ScpiNull();
               

                m_scpi.SendScpi(":STATus:PRESet");
                Wait();
                m_scpi.SendScpi(":MMEMory:LOAD:STATe \"" + STACalFile + "\"");
                //m_scpi.SendScpi(":MMEMory:STORe:SNP:TYPE:S4P 1,2,3,4");
                m_scpi.SendScpi(":MMEMory:STORe:SNP:TYPE:S1P 1");
                if (m_continues == true)
                {
                    m_scpi.SendScpi(":INITiate:CONTinuous 1");
                }
                else
                {
                    m_scpi.SendScpi(":DISPlay:WINDow:ACTivate");
                    Wait();
                    m_scpi.SendScpi(":TRIGger:SEQuence:SOURce BUS");
                    Wait();
                    m_scpi.SendScpi(":INITiate:CONTinuous 0");
                }
                Wait();
                m_scpi.SendScpi(":DISPlay:ENABle 0");
                Wait();
            }
            catch (Exception err)
            {
                if (m_scpi != null)
                    m_scpi.Close();
                m_scpi = null;
                throw (new SystemException(err.Message));
            }
        }

        public void SingleShotCapture(string fileName)
        {
            StartVNAS4PCapture(fileName);
            if (m_copyInside)
            {
                m_monitorDirQueue.Enqueue(fileName);
                m_copyInsideEvent.Set();
            }

        }
        public void Abort()
        {
            m_abort = true;
        }
        protected  virtual bool StartVNAS4PCapture(string fileName)
        {
            try
            {
                string sss = Path.GetFileName(fileName);
                sss = " \"" + "d:\\homertuner\\" + sss + "\"";
                if (m_continues == false)
                {
                    m_scpi.SendScpi(":INITiate:IMMediate");
                    Wait();
                    m_scpi.SendScpi("*TRG");
                    Wait();
                }

                if (m_continues == false)
                {
                    string q;
                    for (int i = 0; i < 1000; i++)
                    {
                        q = m_scpi.QueryScpi(":STATus:OPERation:CONDition?");
                        if (q == "+0")
                            break;
                        if (m_abort == true)
                            return false;
                    }
                }
                else
                {
                    string q;
                    for (int i = 0; i < 1000; i++)
                    {
                        q = m_scpi.QueryScpi(":STATus:OPERation:CONDition?");
                        if (q == "+16")
                            break;
                        if (m_abort == true)
                            return false;
                    }
                }
                m_scpi.SendScpi(":MMEMory:STORe:SNP:DATA" + sss);
                Wait();

                if (m_callback != null)
                    m_callback(200, "Single shot completed");

                return true;
            }
            catch (Exception err)
            {
                throw (new SystemException(err.Message));
            }
        }
    }
}
