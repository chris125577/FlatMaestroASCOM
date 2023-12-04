// ASCOM CoverCalibrator hardware class for FlatMaestro
//
// Description:	 ASCOM driver for LED flat panel via Arduino FlatMaster project
//  Serial command sends brightness over serial interface and receives the same as string.0=off
// Implements:	ASCOM CoverCalibrator interface version: 1.1
// V 1.2    changes brightness range from 0-255 to 0-100
// Author:		(CJW) Chris Woodhouse cwoodhou@icloud.com
//

using ASCOM.Astrometry.AstroUtils;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;
using System;
using System.Collections;
using System.IO.Ports;
using System.Windows.Forms;

namespace ASCOM.FlatMaestro.CoverCalibrator
{
    /// <summary>
    /// ASCOM CoverCalibrator hardware class for FlatMaestro.
    /// </summary>
    [HardwareClass()] // Class attribute flag this as a device hardware class that needs to be disposed by the local server when it exits.
    internal static class CoverCalibratorHardware
    {
        // Constants used for Profile persistence
        internal const string comPortProfileName = "COM Port";
        internal const string comPortDefault = "COM1";
        internal const string traceStateProfileName = "Trace Level";
        internal const string traceStateDefault = "true";
        internal const byte maxlevel = 100;  // maximum brightness level

        private static string DriverProgId = ""; // ASCOM DeviceID (COM ProgID) for this driver, the value is set by the driver's class initialiser.
        private static string DriverDescription = ""; // The value is set by the driver's class initialiser.
        internal static string comPort; // COM port name (if required)

        private static bool connectedState; // Local server's connected state
        private static bool runOnce = false; // Flag to enable "one-off" activities only to run once.
        internal static Util utilities; // ASCOM Utilities object for use as required
        internal static AstroUtils astroUtilities; // ASCOM AstroUtilities object for use as required
        internal static TraceLogger tl; // Local server's trace logger object for diagnostic log with information that you specify
        
        internal static char endChar = '#';  // end character of Arduino response
        internal static char startChar = '$';  // start character of Arduino response
        internal static string buffer = "";
        internal static bool startCharRx;  // indicates start of message detected
        internal static byte LEDlevel = 0; // locally held and updated brigntness level
        static bool dataRx = false;  // indicates that data is available to read (not used actively)
        private static bool arduinoStatus = false;  // Panel on/off status (on can be zero level too)
        static int baud = 19200;
        private static SerialPort Serial;  // my serial port instance of ASCOM serial port

        /// <summary>
        /// Initializes a new instance of the device Hardware class.
        /// </summary>
        static CoverCalibratorHardware()
        {
            try
            {
                // Create the hardware trace logger in the static initialiser.
                // All other initialisation should go in the InitialiseHardware method.
                tl = new TraceLogger("", "FlatMaestro.Hardware");
                Serial = new SerialPort();  // standard .net serial port
                
                // DriverProgId has to be set here because it used by ReadProfile to get the TraceState flag.
                DriverProgId = CoverCalibrator.DriverProgId; // Get this device's ProgID so that it can be used to read the Profile configuration values

                // ReadProfile has to go here before anything is written to the log because it loads the TraceLogger enable / disable state.
                ReadProfile(); // Read device configuration from the ASCOM Profile store, including the trace state

                LogMessage("CoverCalibratorHardware", $"Static initialiser completed.");
            }
            catch (Exception ex)
            {
                try { LogMessage("CoverCalibratorHardware", $"Initialisation exception: {ex}"); } catch { }
                MessageBox.Show($"{ex.Message}", "Exception creating ASCOM.FlatMaestro.CoverCalibrator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        /// <summary>
        /// Place device initialisation code here that delivers the selected ASCOM <see cref="Devices."/>
        /// </summary>
        /// <remarks>Called every time a new instance of the driver is created.</remarks>
        internal static void InitialiseHardware()
        {
            // This method will be called every time a new ASCOM client loads your driver
            LogMessage("InitialiseHardware", $"Start.");

            // Make sure that "one off" activities are only undertaken once
            if (runOnce == false)
            {
                LogMessage("InitialiseHardware", $"Starting one-off initialisation.");

                DriverDescription = CoverCalibrator.DriverDescription; // Get this device's Chooser description

                LogMessage("InitialiseHardware", $"ProgID: {DriverProgId}, Description: {DriverDescription}");

                connectedState = false; // Initialise connected to false
                utilities = new Util(); //Initialise ASCOM Utilities object
                astroUtilities = new AstroUtils(); // Initialise ASCOM Astronomy Utilities object

                LogMessage("InitialiseHardware", "Completed basic initialisation");

                LogMessage("InitialiseHardware", $"One-off initialisation complete.");
                runOnce = true; // Set the flag to ensure that this code is not run again
            }
        }

        // PUBLIC COM INTERFACE ICoverCalibratorV1 IMPLEMENTATION

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialogue form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public static void SetupDialog()
        {
            // Don't permit the setup dialogue if already connected
            if (IsConnected)
                MessageBox.Show("Already connected, just press OK");

            using (SetupLEDdriver F = new SetupLEDdriver(tl))
            {
                var result = F.ShowDialog();
                if (result == DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        /// <summary>Returns the list of custom action names supported by this driver.</summary>
        /// <value>An ArrayList of strings (SafeArray collection) containing the names of supported actions.</value>
        public static ArrayList SupportedActions
        {
            get
            {
                LogMessage("SupportedActions Get", "Returning empty ArrayList");
                return new ArrayList();
            }
        }

        /// <summary>Invokes the specified device-specific custom action.</summary>
        /// <param name="ActionName">A well known name agreed by interested parties that represents the action to be carried out.</param>
        /// <param name="ActionParameters">List of required parameters or an <see cref="String.Empty">Empty String</see> if none are required.</param>
        /// <returns>A string response. The meaning of returned strings is set by the driver author.
        /// <para>Suppose filter wheels start to appear with automatic wheel changers; new actions could be <c>QueryWheels</c> and <c>SelectWheel</c>. The former returning a formatted list
        /// of wheel names and the second taking a wheel name and making the change, returning appropriate values to indicate success or failure.</para>
        /// </returns>
        public static string Action(string actionName, string actionParameters)
        {
            LogMessage("Action", $"Action {actionName}, parameters {actionParameters} is not implemented");
            throw new ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        /// <summary>
        /// Transmits an arbitrary string to the device and does not wait for a response.
        /// Optionally, protocol framing characters may be added to the string before transmission.
        /// </summary>
        /// <param name="Command">The literal command string to be transmitted.</param>
        /// <param name="Raw">
        /// if set to <c>true</c> the string is transmitted 'as-is'.
        /// If set to <c>false</c> then protocol framing characters may be added prior to transmission.
        /// </param>
       
        // CommandBlind, CommandBool and CommandString are not used
        public static void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            CommandString(command, raw);
            return;
        }

        public static bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            return (CommandString(command, raw) == "1");// return state of command from arduino
        }

        public static string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            
            try
            {
                return ("dummy");
            }
            catch (Exception)  // better luck next time :)
            {
                System.Windows.Forms.MessageBox.Show("Timed out, press OK to recover");
                return ("comms error");
            }
        }


        public static void Dispose()
        {
            try { LogMessage("Dispose", $"Disposing of assets and closing down."); } catch { }

            try
            {
                // Clean up the trace logger and utility objects
                tl.Enabled = false;
                tl.Dispose();
                tl = null;
            }
            catch { }

            try
            {
                utilities.Dispose();
                utilities = null;
            }
            catch { }

            try
            {
                astroUtilities.Dispose();
                astroUtilities = null;
            }
            catch { }
        }

        /// <summary>
        /// Set True to connect to the device hardware. Set False to disconnect from the device hardware.
        /// You can also read the property to check whether it is connected. This reports the current hardware state.
        /// </summary>
        /// <value><c>true</c> if connected to the hardware; otherwise, <c>false</c>.</value>
        public static bool Connected
        {
            get
            {
                LogMessage("Connected", $"Get {IsConnected}");
                return IsConnected;
            }
            set
            {
                LogMessage("Connected", $"Set {value}");
                if (value == IsConnected) // already the same state
                    return;

                if (value) // connect
                {
                    LogMessage("Connected Set", $"Connecting to port {comPort}");

                    if (!OpenArduino())  // try to connect to Arduino, if error, reset
                    {
                        Serial.Close();
                        Serial.Dispose();
                        connectedState = false;
                        LEDlevel = 0;
                        LogMessage("Connected Set", $"Failed to connect to port {comPort}");
                    }
                    else connectedState = true;
                }
                else  // disconnect
                {
                    LogMessage("Connected Set", $"Disconnecting from port {comPort}");
                    Serial.Close();  //disconnect to serial
                    Serial.Dispose();
                    connectedState = false;
                }
            }
        }

        public static string Description
        {
            get
            {
                LogMessage("Description Get", DriverDescription);
                return DriverDescription;
            }
        }
        public static string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverInfo = $"Version: {version.Major}.{version.Minor}";
                LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public static string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = $"{version.Major}.{version.Minor}";
                LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public static short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "1");
                return Convert.ToInt16("1");
            }
        }

        public static string Name
        {
            get
            {
                string name = "Flat Maestro";
                LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region ICoverCalibrator Implementation

        /// <summary>
        /// Returns the state of the device cover, if present, otherwise returns "NotPresent"
        /// </summary>
        internal static CoverStatus CoverState
        {
            get
            {
                LogMessage("CoverState Get", "Not implemented");
                return CoverStatus.NotPresent;
            }
        }

        /// <summary>
        /// Initiates cover opening if a cover is present
        /// </summary>
        internal static void OpenCover()
        {
            LogMessage("OpenCover", "Not implemented");
            throw new MethodNotImplementedException("OpenCover");
        }

        /// <summary>
        /// Initiates cover closing if a cover is present
        /// </summary>
        internal static void CloseCover()
        {
            LogMessage("CloseCover", "Not implemented");
            throw new MethodNotImplementedException("CloseCover");
        }

        /// <summary>
        /// Stops any cover movement that may be in progress if a cover is present and cover movement can be interrupted.
        /// </summary>
        internal static void HaltCover()
        {
            LogMessage("HaltCover", "Not implemented");
            throw new MethodNotImplementedException("HaltCover");
        }

        /// <summary>
        /// Returns the state of the calibration device, if present, otherwise returns "NotPresent"
        /// </summary>
        internal static CalibratorStatus CalibratorState
        {
            get
            {
                if (IsConnected)
                {
                    if (arduinoStatus) return (CalibratorStatus.Ready);
                    else return (CalibratorStatus.Off);
                }
                else return CalibratorStatus.NotReady;
            }
        }

        /// <summary>
        /// Returns the current calibrator brightness in the range 0 (completely off) to <see cref="MaxBrightness"/> (fully on)
        /// </summary>
        internal static int Brightness
        {
            get
            {
                if (IsConnected) return (LEDlevel);
                else throw new NotConnectedException("not connected");
            }
        }

        /// <summary>
        /// The Brightness value that makes the calibrator deliver its maximum illumination.
        /// </summary>
        internal static int MaxBrightness
        {
            get
            {
                return (maxlevel);
            }
        }

        /// <summary>
        /// Turns the calibrator on at the specified brightness if the device has calibration capability
        /// </summary>
        /// <param name="Brightness"></param>
        internal static void CalibratorOn(int Brightness)
        {
            if ((Brightness >= 0) && (Brightness <= maxlevel))  // check legitimate value
            {
                LEDlevel = (byte)Brightness; // update local value
                Serial.Write(LEDlevel.ToString() + endChar);
                arduinoStatus = true; // on
            }
            else throw new InvalidValueException("invalid brightness");
            return;
        }

        /// <summary>
        /// Turns the calibrator off if the device has calibration capability
        /// </summary>
        internal static void CalibratorOff()
        {
            CalibratorOn(0);
            arduinoStatus = false;  // off, but connected
            return;
        }

        #endregion

        #region Private properties and methods
        // Useful methods that can be used as required to help with driver development

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private static bool IsConnected
        {
            get
            {
                // check the actual serial connection (checks for unplugged)
                connectedState = Serial.IsOpen;
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private static void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal static void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "CoverCalibrator";
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(DriverProgId, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(DriverProgId, comPortProfileName, string.Empty, comPortDefault);
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal static void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "CoverCalibrator";
                driverProfile.WriteValue(DriverProgId, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(DriverProgId, comPortProfileName, comPort.ToString());
            }
        }

        /// <summary>
        /// Log helper function that takes identifier and message strings
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        internal static void LogMessage(string identifier, string message)
        {
            tl.LogMessageCrLf(identifier, message);
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            LogMessage(identifier, msg);
        }

        // OpenArduino initilises serial port and set up an event handler to suck in characters
        // it runs in the background, Arduino broadcasts every 4 seconds
        private static bool OpenArduino()
        {
            Serial.BaudRate = CoverCalibratorHardware.baud;  // note original was 9600
            Serial.PortName = comPort;
            Serial.Parity = Parity.None;
            Serial.DataBits = 8;
            Serial.Handshake = System.IO.Ports.Handshake.None;
            Serial.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(ReceiveData);
            Serial.ReceivedBytesThreshold = 1;
            try
            {
                Serial.Open();              // open port
                Serial.DiscardInBuffer();   // and clear it out just in case
                //Serial.Write("40#");  // diagnostic only
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
        // ReceiveData is based on a code fragment suggested by Per and reads characters as they arrive
        // it decodes the messages, looking for framing characters and then splits the CSV string into
        // component parts to represent the status flags from the Arduino 
        private static void ReceiveData(object sender, SerialDataReceivedEventArgs e)
        {
            if (e.EventType == System.IO.Ports.SerialData.Chars)
            {
                while (Serial.BytesToRead > 0)
                {
                    char c = (char)Serial.ReadChar();  // wait for start character
                    if (!startCharRx)
                    {
                        if (c == startChar)  // and then initialise the message
                        {
                            startCharRx = true;
                            buffer = "";  // clear buffer
                        }
                    }
                    else
                    {
                        if (c == endChar)
                        {
                            if ((buffer.Length < 4) && (buffer.Length > 0)) // nnn, nn or n
                            {
                                dataRx = true; // tell the world that data is available
                                LEDlevel = byte.Parse(buffer);  // transfer the buffer to the message and clear the buffer
                            }
                            else  // message was corrupted
                            {
                                dataRx = false;
                                tl.LogMessage("communications", "corrupted message length");
                            }
                            buffer = "";
                            startCharRx = false;
                        }
                        else
                        {
                            buffer += c;  // build up message string in buffer
                        }
                    }
                }
            }
        }

        #endregion
    }
}

