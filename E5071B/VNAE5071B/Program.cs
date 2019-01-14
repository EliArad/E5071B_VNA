/***************************************************
 *        Copyright Keysight Technologies 2006-2016
 **************************************************/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Agilent.AgilentNA.Interop;

namespace AgilentNA_CS_Example1
{
    /// <summary>
    /// AgilentNA IVI-COM Driver Example Program
    /// 
    /// Creates a driver object, reads a few Identity interface properties, and checks the instrument error queue.
    /// May include additional instrument specific functionality.
    /// 
    /// See driver help topic "Programming with the IVI-COM Driver in Various Development Environments"
    /// for additional programming information.
    ///
    /// Runs in simulation mode without an instrument.
    /// 
    /// Requires a reference to the driver's interop or COM type library.
    /// 
    /// </summary>
    class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {

            VNAE5071B m_vnaE5071B = new VNAE5071B();
            double[] data;
            double[] freq;
            m_vnaE5071B.MeasureS11(out data, out freq);
            m_vnaE5071B.MeasureS22(out data, out freq);
            m_vnaE5071B.MeasureS12(out data, out freq);
            m_vnaE5071B.MeasureS21(out data, out freq);

        }
        
    }
}
