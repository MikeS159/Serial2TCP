#region Licence
//MIT License(MIT)

/*     AsyncSerial.cs Version 4.3.0        */

/*     Copyright(c) 2018 Mike Simpson      */

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.
#endregion


using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Ports;

namespace AsyncSerial
{

    #region classRegion

    /// <summary>
    /// This class allows a serial port to be set up with different parameters.
    /// It contains functions to open and close the port safely
    /// Implements IDisposable
    /// </summary>
    public class AsyncSerial : IDisposable
    {

        #region Varibles

        /// <summary>
        /// Underlying SerialPort object supported by .NET
        /// </summary>
        SerialPort dotNETPort;

        /// <summary>
        /// Contains an error message if any of the functions return false
        /// </summary>
        private string portErrorMessage = "";

        #endregion

        #region Constructo
        /// <summary>
        /// Constructor take 5 arguments to correctly configure the serial port
        /// </summary>
        /// <param name="portName"> Sets the port name of the serial port (string) </param>
        /// <param name="baudRate"> Sets the baud rate of the serial port (int) </param>
        /// <param name="parity"> Sets the parity bit system to use (string) </param>
        /// <param name="dataBits"> Sets the number of data bits of the serial port (int) </param>
        /// <param name="stopBits"> Sets the number of stop bits of the serial port (string) </param>
        public AsyncSerial(string portName, int baudRate, string parity, int dataBits, string stopBits)
        {
            dotNETPort = new SerialPort();
            if (portName != null & parity != null & stopBits != null)
            {
                try
                {
                    switch (stopBits.ToLower(CultureInfo.CurrentCulture))
                    {
                        case "none": dotNETPort.StopBits = StopBits.None; break;
                        case "onepointfive": dotNETPort.StopBits = StopBits.OnePointFive; break;
                        case "two": dotNETPort.StopBits = StopBits.Two; break;
                        default: dotNETPort.StopBits = StopBits.One; break;
                    }

                    switch (parity.ToLower(CultureInfo.CurrentCulture))
                    {
                        case "even": dotNETPort.Parity = System.IO.Ports.Parity.Even; break;
                        case "odd": dotNETPort.Parity = System.IO.Ports.Parity.Odd; break;
                        case "mark": dotNETPort.Parity = System.IO.Ports.Parity.Mark; break;
                        case "space": dotNETPort.Parity = System.IO.Ports.Parity.Space; break;
                        default: dotNETPort.Parity = System.IO.Ports.Parity.None; break;
                    }

                    dotNETPort.PortName = portName;
                    dotNETPort.BaudRate = baudRate;
                    dotNETPort.DataBits = dataBits;
                }
                catch (IOException e)
                {
                    exceptionCatch(e);
                }
                catch (ArgumentNullException e)
                {
                    exceptionCatch(e);
                }
                catch (ArgumentOutOfRangeException e)
                {
                    exceptionCatch(e);
                }
                catch (ArgumentException e)
                {
                    exceptionCatch(e);
                }
                catch (InvalidOperationException e)
                {
                    exceptionCatch(e);
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Calls the .NET IDisposable method on the serial port and cleans up with the garbage collector
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Uses closes the port properly and disposes using the serial port IDisposable method
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
                try
                {
                    dotNETPort.Close();
                }
                catch (IOException e)
                {
                    exceptionCatch(e);
                }
                dotNETPort.Dispose();
            }
            // free native resources if there are any
        }

        /// <summary>
        /// Open the serial port with defined values, returns false and puts error message in 'portError' if the port fails to open
        /// </summary>
        public bool StartReceive()
        {
            if (dotNETPort != null)
            {
                try
                {
                    dotNETPort.Open();
                }
                catch (InvalidOperationException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (IOException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (UnauthorizedAccessException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (ArgumentOutOfRangeException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (ArgumentException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }

                finally
                {
                    dotNETPort.DataReceived += new SerialDataReceivedEventHandler(port_OnDataRecived);    //sets up new event handler
                }

                return true;
            }
            else
            {
                exceptionCatch("Port Null");
                return false;
            }
        }

        /// <summary>
        /// Exceptions can be passed to be recorded in the port error message
        /// </summary>
        /// <param name="Ex"></param>
        private void exceptionCatch(Exception Ex)
        {
            portErrorMessage = defaltErrorText() + Ex.Message;
        }

        /// <summary>
        /// Strings can be passed to be recorded in the port error message
        /// </summary>
        /// <param name="Ex"></param>
        private void exceptionCatch(string Ex)
        {
            portErrorMessage = defaltErrorText() + Ex;
        }

        /// <summary>
        /// Returns a stringing containing defalt text about errors which specific error information can be appended to
        /// </summary>
        /// <returns></returns>
        private string defaltErrorText()
        {
            return "An error occurred while trying to open the serial port " + dotNETPort.PortName.ToString() + "\nPlease check the port settings and try again\n" + "Error message - ";
        }

        /// <summary>
        /// Safely closes the port if open
        /// </summary>
        public void StopReceive()
        {
            try
            {
                if (dotNETPort != null)
                {
                    if (dotNETPort.IsOpen)
                    {
                        dotNETPort.Close();
                    }
                }
            }
            catch (ArgumentException Ex)
            {
                exceptionCatch(Ex);
            }
            catch (IOException Ex)
            {
                exceptionCatch(Ex);
            }
        }

        /// <summary>
        /// This function can be used to get the port status externally
        /// </summary>
        public bool IsPortOpen
        {
            get
            {
                if (dotNETPort != null)
                {
                    return dotNETPort.IsOpen;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns an empty/null string if there is no error, or the error message as a string to be displayed to the user, auto clears once read
        /// </summary>
        public string CurrentPortError
        {
            get
            {
                string returnError = portErrorMessage;
                portErrorMessage = "";
                return returnError;
            }
        }

        /// <summary>
        /// Transmitts a string over the serial port using the Serial.Write() method
        /// Encoding should be considered before passing string
        /// </summary>
        /// <param name="sendData"></param>
        public bool SendSerialData(string sendData)
        {
            if (IsPortOpen & sendData != null)
            {
                try
                {
                    dotNETPort.Write(sendData);
                    return true;
                }
                catch (InvalidOperationException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (ArgumentNullException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (OverflowException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
            }
            else
            {
                exceptionCatch("Port nto open");
                return false;
            }
        }

        /// <summary>
        /// Transmitts a byte array over the serial port using the Serial.write() method
        /// Can be used to send custom data formats without text overhead
        /// </summary>
        /// <param name="sendData"></param>
        public bool SendSerialData(byte[] sendData)
        {
            if (IsPortOpen & sendData != null)
            {
                try
                {
                    dotNETPort.Write(sendData, 0, sendData.Length);
                    return true;
                }

                catch (InvalidOperationException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (ArgumentNullException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (ArgumentOutOfRangeException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (ArgumentException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
                catch (OverflowException Ex)
                {
                    exceptionCatch(Ex);
                    return false;
                }
            }
            else
            {
                exceptionCatch("Port not open");
                return false;
            }
        }

        #endregion

        #region Events

        /// <summary>
        /// This function executes when data is recived on the serial port and triggers the event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void port_OnDataRecived(object sender, SerialDataReceivedEventArgs e)
        {
            int lengthToRead = dotNETPort.BytesToRead;
            byte[] rxBytes = new byte[lengthToRead];
            dotNETPort.Read(rxBytes, 0, lengthToRead);
            OnPacketReceived(rxBytes);
        }

        /// <summary>
        /// The event that triggers when ever data is recived on the serial port
        /// Operations started from this even run on their own thread
        /// </summary>
        public event EventHandler<PacketEventArgs> PacketReceived;

        /// <summary>
        /// Sends packets to calling function to be processed
        /// </summary>
        /// <param name="packet"></param>
        protected virtual void OnPacketReceived(byte[] packet)
        {
            PacketReceived?.Invoke(this, new PacketEventArgs(packet));
        }
    }

    /// <summary>
    /// Custom class for packet recived events, based on EventArgs
    /// </summary>
    public class PacketEventArgs : EventArgs
    {
        private byte[] packet;

        /// <summary>
        /// Allows new packets to be read by and set by constructor
        /// </summary>
        public byte[] Packet()
        {
            return packet;
        }

        /// <summary>
        /// Constructor for PacketEventArgs, reads new packets
        /// </summary>
        /// <param name="inPacket"></param>
        public PacketEventArgs(byte[] inPacket)
        {
            packet = inPacket;
        }
    }

    #endregion

    #endregion

    #region PortParameters

    /// <summary>
    /// Provides string arrays containing common (or enforced) port parameters which you can use to populate menus in foreach loops etc
    /// ReadOnlyCollections can be cast to a array using .ToArray()
    /// </summary>
    public static class SerialPortParameters
    {

        private static string[] commonBaudRates = new string[] {"110", "300", "600", "1200", "2400", "4800", "9600", "14400", "19200", "28800", "38400", "56000",
            "57600", "115200", "128000", "153600", "230400", "256000", "460800", "921600", "1843200"};
        private static ReadOnlyCollection<string> readOnlyCommonBaudRates = new ReadOnlyCollection<string>(commonBaudRates);

        private static string[] commonStopBits = new string[] { "none", "one", "onepointfive", "two" };
        private static ReadOnlyCollection<string> readOnlyCommonStopBits = new ReadOnlyCollection<string>(commonStopBits);

        private static string[] commonParityBits = new string[] { "none", "odd", "even", "mark", "space" };
        private static ReadOnlyCollection<string> readOnlyCommonParityBits = new ReadOnlyCollection<string>(commonParityBits);

        private static string[] commonDataBits = new string[] { "7", "8" };
        private static ReadOnlyCollection<string> readOnlyCommonDataBits = new ReadOnlyCollection<string>(commonDataBits);

        /// <summary>
        /// Commonly used serial baud rates, system dependant but usually produced by a 1.843200MHz oscillator
        /// </summary>
        public static ReadOnlyCollection<string> BaudRates
        {
            get
            {
                return readOnlyCommonBaudRates;
            }
        }

        /// <summary>
        /// Only supported stop bits as of .NET 4.6.1
        /// </summary>
        public static ReadOnlyCollection<string> StopBits
        {
            get
            {
                return readOnlyCommonStopBits;
            }
        }

        /// <summary>
        /// Only supported parity bits as of .NET 4.6.1
        /// </summary>
        public static ReadOnlyCollection<string> ParityBits
        {
            get
            {
                return readOnlyCommonParityBits;
            }
        }

        /// <summary>
        /// Genrally supported data bits (7 for ASCII, 8 for ease of use)
        /// </summary>
        public static ReadOnlyCollection<string> DataBits
        {
            get
            {
                return readOnlyCommonDataBits;
            }
        }

        /// <summary>
        /// Gets currently recognised serial port names connected to the computer, can be called when ever refreshed list is needed
        /// </summary>
        /// <returns></returns>
        public static ReadOnlyCollection<string> GetSerialPorts
        {
            get
            {
                ReadOnlyCollection<string> portNames = new ReadOnlyCollection<string>(SerialPort.GetPortNames());
                return portNames;
            }
        }

    }
    #endregion

    #region ExampleCode

    // <summary>
    // WINFORMS EXAMPLE
    // SerialPortParameters can be used to fill form elements for easier user selection
    // </summary>
    //private void Form_Load(object sender, EventArgs e)
    //{
    //    serialPortBaud_CB.Items.AddRange(SerialPortParameters.BaudRates.ToArray());
    //    serialPortData_CB.Items.AddRange(SerialPortParameters.DataBits.ToArray());
    //    serialPortParity_CB.Items.AddRange(SerialPortParameters.ParityBits.ToArray());
    //    serialPortStop_CB.Items.AddRange(SerialPortParameters.StopBits.ToArray());
    //}

    // <summary>
    // XAML EXAMPLE
    // Adds common values for each serial port setting to a combobox and selects a default
    // </summary>
    //private void populateSerialSettings()
    //{
    //    List<string> itemsToAdd = SerialPortParameters.BaudRates.ToList();
    //    itemsToAdd.ForEach(item => serialPortBaudRate_CB.Items.Add(item));
    //    serialPortBaudRate_CB.SelectedItem = "115200";

    //    itemsToAdd.Clear();
    //    itemsToAdd = SerialPortParameters.DataBits.ToList();
    //    itemsToAdd.ForEach(item => serialPortDataBits_CB.Items.Add(item));
    //    serialPortDataBits_CB.SelectedItem = "8";

    //    itemsToAdd.Clear();
    //    itemsToAdd = SerialPortParameters.ParityBits.ToList();
    //    itemsToAdd.ForEach(item => serialPortParity_CB.Items.Add(item));
    //    serialPortParity_CB.SelectedItem = "none";

    //    itemsToAdd.Clear();
    //    itemsToAdd = SerialPortParameters.StopBits.ToList();
    //    itemsToAdd.ForEach(item => serialPortStopBits_CB.Items.Add(item));
    //    serialPortStopBits_CB.SelectedItem = "one";
    //}

    // <summary>
    // Example of how to setup the serial port and event handler
    // </summary>
    //private void setupPort()
    //{
    //    string portName = serialPortName_TB.Text;
    //    int portBaud = Convert.ToInt32(serialPortBaud_CB.SelectedItem.ToString());
    //    string portParity = serialPortParity_CB.SelectedItem.ToString();
    //    int portData = Convert.ToInt32(serialPortData_CB.SelectedItem.ToString());
    //    string portStop = serialPortStop_CB.SelectedItem.ToString();

    //    comPort = new AsyncSerial(portName, portBaud, portParity, portData, portStop); //comPort should be declared at class level
    //if (comPort.CurrentPortError != null | comPort.CurrentPortError != "")
    //{
    //    //No Error
    //}
    //else
    //{
    //    //Handle Error
    //    comPort.clearPortError(); //Reset for next attempt
    //}
    //comPort.PacketReceived += onComSerial_receive;
    //if (comPort.StartReceive())
    //{

    //}
    //else
    //{
    //    //Port Error
    //}
    //}

    // <summary>
    // This is the code that gets called when serial data is recived
    // </summary>
    //private void onMySerial_receive(object sender, PacketEventArgs e)
    //{
    //    byte[] packet = e.Packet;
    //    string supplyMessage = Encoding.Default.GetString(packet); //Conver to string if needed

    //    //WINFORM EXAMPLE
    //    this.Invoke((MethodInvoker)delegate
    //    {
    //        someFormItem.property = someData; //Since the port is asynchronous you cannot change UI elements without a cross thread Invoker
    //    });

    //    //XAML EXAMPLE
    //    Dispatcher.Invoke(() => someFormItem.property = someData); //Since the port is asynchronous you cannot change UI elements without a cross thread Invoker
    //}

    //private void sendSerialData()
    //{
    //    if (comPort.IsPortOpen()) //you can check if the port is open (device could be removed etc)
    //    {
    //        comPort.SendSerialData("you can send strings or byte arrays");
    //    }
    //    else
    //    {
    //        //port isn't open
    //    }
    //}

    #endregion
}
