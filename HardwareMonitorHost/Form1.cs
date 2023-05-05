using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Lifetime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenHardwareMonitor.Hardware;
using Newtonsoft.Json;
using System.IO.Ports;
using System.Diagnostics;

namespace HardwareMonitorHost
{
    public partial class Form1 : Form
    {
        public static SerialPort _serialPort = new SerialPort("COM11", 115200, Parity.None, 8, StopBits.One);

        public class UpdateVisitor : IVisitor
        {
            public void VisitComputer(IComputer computer)
            {
                computer.Traverse(this);
            }
            public void VisitHardware(IHardware hardware)
            {
                hardware.Update();
                foreach (IHardware subHardware in hardware.SubHardware) subHardware.Accept(this);
            }
            public void VisitSensor(ISensor sensor) { }
            public void VisitParameter(IParameter parameter) { }
        }
        public class HwInfo
        {
            public int CT { get; set; }
            public int CU { get; set; }
            public int CF { get; set; }
            public int GT { get; set; }
            public int GU { get; set; }
            public int GF { get; set; }
            public int RT { get; set; }
            public int RU { get; set; }
        }
        public static void GetSystemInfo(Form1 theForm)
        {
            UpdateVisitor updateVisitor = new UpdateVisitor();
            HwInfo hwInfo = new HwInfo();
            Computer computer = new Computer();
            computer.Open();
            computer.CPUEnabled = true;
            computer.GPUEnabled = true;
            computer.RAMEnabled = true;
            computer.MainboardEnabled = true;
            computer.Accept(updateVisitor);
            for (int i = 0; i < computer.Hardware.Length; i++)
            {
                if (computer.Hardware[i].HardwareType == HardwareType.Mainboard)
                {

                    if (computer.Hardware[i].SubHardware.Length > 0)
                    {
                        int temp = 0;
                        for (int j = 0; j < computer.Hardware[i].SubHardware[0].Sensors.Length; j++)
                        {
                            if (computer.Hardware[i].SubHardware[0].Sensors[j].SensorType == SensorType.Temperature)
                            {
                                temp += (int)computer.Hardware[i].SubHardware[0].Sensors[j].Value;
                            }
                        }
                        hwInfo.RT = (int)temp / 6;
                    }

                }
                if (computer.Hardware[i].HardwareType == HardwareType.CPU)
                {
                    int fcpu = 0;

                    for (int j = 0; j < computer.Hardware[i].Sensors.Length; j++)
                    {
                        if (computer.Hardware[i].Sensors[j].SensorType == SensorType.Clock && computer.Hardware[i].Sensors[j].Name != "Bus Speed")
                            fcpu += (int)computer.Hardware[i].Sensors[j].Value;

                        if (computer.Hardware[i].Sensors[j].SensorType == SensorType.Temperature)
                            hwInfo.CT = (int)computer.Hardware[i].Sensors[j].Value;

                        if (computer.Hardware[i].Sensors[j].SensorType == SensorType.Load && computer.Hardware[i].Sensors[j].Name == "CPU Total")
                            hwInfo.CU = (int)computer.Hardware[i].Sensors[j].Value;

                    }
                    hwInfo.CF = (int)fcpu / 8;
                }

                if (computer.Hardware[i].HardwareType == HardwareType.GpuNvidia)
                {
                    for (int j = 0; j < computer.Hardware[i].Sensors.Length; j++)
                    {
                        if (computer.Hardware[i].Sensors[j].SensorType == SensorType.Clock && computer.Hardware[i].Sensors[j].Name == "GPU Core")
                            hwInfo.GF = (int)computer.Hardware[i].Sensors[j].Value;

                        if (computer.Hardware[i].Sensors[j].SensorType == SensorType.Temperature)
                            hwInfo.GT = (int)computer.Hardware[i].Sensors[j].Value;

                        if (computer.Hardware[i].Sensors[j].SensorType == SensorType.Load && computer.Hardware[i].Sensors[j].Name == "GPU Core")
                            hwInfo.GU = (int)computer.Hardware[i].Sensors[j].Value;

                    }
                }
                if (computer.Hardware[i].HardwareType == HardwareType.RAM)
                {
                    for (int j = 0; j < computer.Hardware[i].Sensors.Length; j++)
                    {
                        if (computer.Hardware[i].Sensors[j].SensorType == SensorType.Load)
                            hwInfo.RU = (int)computer.Hardware[i].Sensors[j].Value;
                    }
                }
            }
            computer.Close();
            string stringjson = JsonConvert.SerializeObject(hwInfo);
            Debug.WriteLine(stringjson);

            if (!_serialPort.IsOpen)
            {
                try { _serialPort.Open(); }
                catch 
                { 
                    Debug.WriteLine("Cannot Open device");
                    theForm.toolStripMenuItem1.Text = "Not Connected";
                }
            }
            if (_serialPort.IsOpen)
            {
                try { _serialPort.Write(stringjson); }
                catch 
                { 
                    _serialPort.Close();
                    theForm.toolStripMenuItem1.Text = "Not Connected";
                    return;
                }
                theForm.toolStripMenuItem1.Text = "Connected";
                //    

            }
        }
        public Form1()
        {
            InitializeComponent();
            _serialPort.WriteTimeout = 1000;
            this.ShowInTaskbar = false;
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            notifyIcon1.Dispose();
            Application.Exit();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            GetSystemInfo(this);
        }
    }
}
