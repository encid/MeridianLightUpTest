using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Management;
using System.IO.Ports;
using System.Threading;
using System.Text.RegularExpressions;

namespace MeridianLightUpTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AddPortsToListBox();
        }

        public void AddPortsToListBox()
        {
            try
            {
                listBox1.Items.Clear();
                var ports = GetAllPorts3();
                foreach (var item in ports)
                {
                    listBox1.Items.Add(item.ToString());
                }

                if (listBox1.Items.Count > 0)
                {
                    listBox1.SelectedIndex = 0;
                }
            }
            catch 
            {
            }           
        }

        public List<string> GetAllPorts()
        {
            List<string> allPorts = new List<string>();
            foreach (string portName in SerialPort.GetPortNames())
            {
                allPorts.Add(portName);
            }
            return allPorts;
        }

        public List<string> GetAllPorts2()
        {
            var pList = new List<string>();
            listBox1.Items.Clear();
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = GetAllPorts();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());

                var portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();

                foreach (string s in portList)
                {
                    pList.Add(s);
                }
            }

            return pList;
        }

        /// <summary>
        /// Preferred method for getting ports on Windows
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllPorts3()
        {
            var pList = new List<string>();
            
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%USB Serial Port%'"))
            {
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());

                foreach (string port in ports)
                {
                    // hacky way of getting COM port names
                    string substr = port.Substring(port.Length - 7);
                    string portName = Regex.Replace(substr, @"[^a-zA-Z0-9]", "");                    

                    pList.Add(portName);            
                }
            }

            return pList;
        }

        public List<string> GetAllPorts4()
        {
            var pList = new List<string>();

            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_USBController WHERE DeviceID like '%'"))
            {
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["DeviceID"].ToString());

                foreach (string port in ports)
                {
                    pList.Add(port);
                    //if (port.Contains("USB Serial Port"))
                    //{
                    //    pList.Add(port.Substring(17).Trim(')'));
                    //}
                }
            }

            return pList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddPortsToListBox();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //if (listBox1.Items.Count < 1)
            //{
            //    return;
            //}

            string comPort = Convert.ToString(listBox1.SelectedItem);
            int baudRate = 57600;

            SendCommand(comPort, baudRate);
        }

        private void SendCommand(string comPort, int baudRate)
        {
            try
            {
                SerialPort port = new SerialPort(comPort, baudRate);

                if (!port.IsOpen)
                    port.Open();
                else
                    return;

                // Enter Diagnostic mode
                port.Write("M 1 30 1 5000 1 0 1 1234\r\n");
                Thread.Sleep(150);
                // Turn on all display/LEDs
                port.Write("M 1 30 1 5016 0 0 1 1\r\n");
                port.Close();
            }
            catch 
            {
                
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string comPort = listBox1.SelectedItem.ToString();
            int baudRate = 9600;

            SendCommand(comPort, baudRate);
        }
    }
}
