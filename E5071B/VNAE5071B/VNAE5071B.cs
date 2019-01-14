using Agilent.AgilentNA.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AgilentNA_CS_Example1
{
    public class VNAE5071B
    {


        AgilentNA m_driver = null;
        IAgilentNAChannel Ch1;

        public VNAE5071B(string visaName = "TCPIP0::192.168.15.3::5025::SOCKET")
        {
            // Create driver instance
            m_driver = new AgilentNA();

            // Edit resource and options as needed.  Resource is ignored if option Simulate=true
            //string resourceDesc = "GPIB0::16::INSTR";
            string resourceDesc = visaName;

            string initOptions = "QueryInstrStatus=true, Simulate=false, DriverSetup= Model=E5071B, Trace=false, TraceName=c:\\temp\\traceOut";

            bool idquery = true;
            bool reset = true;

            // Initialize the driver.  See driver help topic "Initializing the IVI-COM Driver" for additional information
            m_driver.Initialize(resourceDesc, idquery, reset, initOptions);
            Console.WriteLine("Driver Initialized");

            // Print a few IIviDriverIdentity properties
            Console.WriteLine("Identifier:  {0}", m_driver.Identity.Identifier);
            Console.WriteLine("Revision:    {0}", m_driver.Identity.Revision);
            Console.WriteLine("Vendor:      {0}", m_driver.Identity.Vendor);
            Console.WriteLine("Description: {0}", m_driver.Identity.Description);
            Console.WriteLine("Model:       {0}", m_driver.Identity.InstrumentModel);
            Console.WriteLine("FirmwareRev: {0}", m_driver.Identity.InstrumentFirmwareRevision);
            Console.WriteLine("Serial #:    {0}", m_driver.System.SerialNumber);
            Console.WriteLine("\nSimulate:    {0}\n", m_driver.DriverOperation.Simulate);


            // Setup convineient pointers to Channels and Measurements
            Ch1 = m_driver.Channels.get_Item("Channel1");

            m_driver.System.RecallState(@"d:\icvjig\ICVJIG1.STA");


        }


        public void Close()
        {
            if (m_driver != null && m_driver.Initialized)
            {
             
                m_driver.Close();
            }
        }

        public void MeasureS22(out double[] data, out double[] freq)
        {
            IAgilentNAMeasurement Ch1xx = Ch1.Measurements.get_Item("Measurement1");
            
            Ch1xx.Create(2, 2);  
            m_driver.Trigger.Source = AgilentNATriggerSourceEnum.AgilentNATriggerSourceManual;
            Ch1.TriggerMode = AgilentNATriggerModeEnum.AgilentNATriggerModeContinuous;
            Ch1xx.Format = AgilentNAMeasurementFormatEnum.AgilentNAMeasurementLogMag;
            Console.WriteLine("Measuring Channel1 S22 Data...");
            Ch1.TriggerSweep(2000); // Take sweep and wait up to 2 seconds for sweep to complete
            Ch1xx.Trace.AutoScale();

            FetchData(Ch1xx, out data, out freq);
        }

        public void MeasureS12(out double[] data, out double[] freq)
        {
            IAgilentNAMeasurement Ch1xx = Ch1.Measurements.get_Item("Measurement1");
            
            Ch1xx.Create(1, 2);  
            m_driver.Trigger.Source = AgilentNATriggerSourceEnum.AgilentNATriggerSourceManual;
            Ch1.TriggerMode = AgilentNATriggerModeEnum.AgilentNATriggerModeContinuous;
            Ch1xx.Format = AgilentNAMeasurementFormatEnum.AgilentNAMeasurementLogMag;
            Console.WriteLine("Measuring Channel1 S12 Data...");
            Ch1.TriggerSweep(2000); // Take sweep and wait up to 2 seconds for sweep to complete
            Ch1xx.Trace.AutoScale();

            FetchData(Ch1xx, out data, out freq);
        }

        public void MeasureS21(out double[] data, out double[] freq)
        {
            IAgilentNAMeasurement Ch1xx = Ch1.Measurements.get_Item("Measurement1");
            Ch1xx.Create(2, 1);  
            m_driver.Trigger.Source = AgilentNATriggerSourceEnum.AgilentNATriggerSourceManual;
            Ch1.TriggerMode = AgilentNATriggerModeEnum.AgilentNATriggerModeContinuous;
            Ch1xx.Format = AgilentNAMeasurementFormatEnum.AgilentNAMeasurementLogMag;
            Console.WriteLine("Measuring Channel1 S21 Data...");
            Ch1.TriggerSweep(2000); // Take sweep and wait up to 2 seconds for sweep to complete
            Ch1xx.Trace.AutoScale();

            FetchData(Ch1xx, out data, out freq);
        }

        public void MeasureS11(out double[] data, out double[] freq)
        {
              
            IAgilentNAMeasurement Ch1xx = Ch1.Measurements.get_Item("Measurement1");

            // Setup S11 imaginary data measurement
            Ch1xx.Create(1, 1); // Define S11 measurement
            m_driver.Trigger.Source = AgilentNATriggerSourceEnum.AgilentNATriggerSourceManual;
            Ch1.TriggerMode = AgilentNATriggerModeEnum.AgilentNATriggerModeContinuous;
            Ch1xx.Format = AgilentNAMeasurementFormatEnum.AgilentNAMeasurementLogMag;
            Console.WriteLine("Measuring Channel1 S11 Data...");
            Ch1.TriggerSweep(2000); // Take sweep and wait up to 2 seconds for sweep to complete
            Ch1xx.Trace.AutoScale();

            FetchData(Ch1xx, out data, out freq);
        }

        void FetchData(IAgilentNAMeasurement Ch1xx, out double[] data, out double[] freq)
        {
           
            int i, points;

            // Read and output data
            data = Ch1xx.FetchFormatted();
            freq = Ch1xx.FetchX();
            points = data.Length;
            Console.WriteLine("First 5 of {0} points, Imaginary Data (Frequency, Magnitude):", points);
            for (i = 0; i < 5; i++)
            {
                Console.WriteLine("  {0}\t{1}", freq[i], data[i]);
            }

            // Check instrument for errors
            int errorNum = -1;
            string errorMsg = null;
            Console.WriteLine();
            while (errorNum != 0)
            {
                m_driver.Utility.ErrorQuery(ref errorNum, ref errorMsg);
                Console.WriteLine("ErrorQuery: {0}, {1}", errorNum, errorMsg);
            }
        }
    }
}
