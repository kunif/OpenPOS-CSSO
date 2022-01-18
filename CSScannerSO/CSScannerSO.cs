
/*

   Copyright (C) 2022 Kunio Fukuchi

   This software is provided 'as-is', without any express or implied
   warranty. In no event will the authors be held liable for any damages
   arising from the use of this software.

   Permission is granted to anyone to use this software for any purpose,
   including commercial applications, and to alter it and redistribute it
   freely, subject to the following restrictions:

   1. The origin of this software must not be misrepresented; you must not
      claim that you wrote the original software. If you use this software
      in a product, an acknowledgment in the product documentation would be
      appreciated but is not required.

   2. Altered source versions must be plainly marked as such, and must not be
      misrepresented as being the original software.

   3. This notice may not be removed or altered from any source distribution.

   Kunio Fukuchi

 */

namespace OpenPOS.CSSO
{
    using Microsoft.Win32;
    using OpenPOS.Devices;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IO.Ports;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using System.Threading;

    [ComVisible(false)]
    internal class OposEventScanner : OposEvent
    {
        internal byte[] _scanData;
        internal int _scanDataType;
        internal byte[] _scanDataLabel;

        #region IDisposable Support

        public OposEventScanner()
        {
            _scanDataType = (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN;
            _scanData = new byte[0];
            _scanDataLabel = new byte[0];
        }

        private bool _disposedValue = false;

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (!this._disposedValue)
                {
                    if (disposing)
                    {
                        _scanDataType = (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN;
                        _scanData = null;
                        _scanDataLabel = null;
                    }
                    this._disposedValue = true;
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        #endregion IDisposable Support
    }

    [ComVisible(true)]
    [Guid("7FA68B58-327C-4E56-9882-F1704C64B6F7")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.Scanner.OpenPOS.CSSO.CSScannerSO.1")]
    public class CSScannerSO : IOPOSScannerSO, IDisposable
    {
        internal int _capPowerReporting;
        internal bool _capCompareFirmwareVersion;
        internal bool _capUpdateFirmware;
        internal bool _capStatisticsReporting;
        internal bool _capUpdateStatistics;

        internal volatile bool _autoDisable;
        internal int _binaryConversion;
        internal string _checkHealthText;
        internal bool _claimed;
        internal volatile int _dataCount;
        internal volatile bool _dataEventEnabled;
        internal volatile bool _deviceEnabled;
        internal volatile bool _freezeEvents;
        //internal int _outputID;
        internal int _powerNotify;
        internal int _powerState;
        internal int _resultCode;
        internal int _resultCodeExtended;
        internal int _state;
        internal string _serviceObjectDescription;
        internal int _serviceObjectVersion;
        internal string _deviceDescription;
        internal string _deviceName;

        internal volatile bool _coFreezeEvents;
        internal int _openResult;

        internal volatile bool _decodeData;
        internal volatile int _scanDataType;

        internal volatile byte[] _scanData;
        internal volatile byte[] _scanDataLabel;

        internal volatile bool _opened;
        internal dynamic _oposCO;
        internal string _deviceNameKey;

        internal enum SymbologyMode : int { None, AIM, Honeywell };

        internal int _minimumDataLength;
        internal int _prefixLength;
        internal byte[] _prefixData;
        internal SymbologyMode _symbologyMode;
        internal int _suffixLength;
        internal byte[] _suffixData;

        internal ClaimManager _claimManager;

        internal volatile bool _releaseClosing;

        internal volatile Object _dataAccessLock;
        internal volatile Mutex _eventControlMutex;
        internal volatile Mutex _eventListMutex;
        internal volatile EventWaitHandle _eventReactivateEvent;

        internal ThreadStart _eventThreadStart;
        internal Thread _eventThread;
        internal volatile List<OposEventScanner> _oposEvents;
        internal volatile bool _killEventThread;

        internal SerialPort _serialPort;
        internal byte[] _portBuffer;
        internal volatile Object _portBufferLock;
        internal volatile EventWaitHandle _portReceivedEvent;

        internal ThreadStart _parseThreadStart;
        internal Thread _parseThread;
        internal volatile bool _killParseThread;

        #region IDisposable Support Constructer / Destructer

        public CSScannerSO()
        {
            _capPowerReporting = (int)OPOS_Constants.OPOS_PR_STANDARD;
            _capCompareFirmwareVersion = false;
            _capUpdateFirmware = false;
            _capStatisticsReporting = true;
            _capUpdateStatistics = true;

            _autoDisable = false;
            _binaryConversion = (int)OPOS_Constants.OPOS_BC_NONE;
            _checkHealthText = string.Empty;
            _claimed = false;
            _dataCount = 0;
            _dataEventEnabled = false;
            _deviceEnabled = false;
            _freezeEvents = false;
            //_outputID = 0;
            _powerNotify = (int)OPOS_Constants.OPOS_PN_DISABLED;
            _powerState = (int)OPOS_Constants.OPOS_PS_UNKNOWN;
            _resultCode = (int)OPOS_Constants.OPOS_E_CLOSED;
            _resultCodeExtended = 0;
            _state = (int)OPOS_Constants.OPOS_S_CLOSED;
            _serviceObjectDescription = string.Empty;
            _serviceObjectVersion = 1016000;
            _deviceDescription = string.Empty;
            _deviceName = string.Empty;

            _coFreezeEvents = false;
            _openResult = (int)OPOS_Constants.OPOS_ORS_CONFIG;

            _decodeData = false;
            _scanDataType = 0;

            _scanData = new byte[0];
            _scanDataLabel = new byte[0];

            _opened = false;
            _deviceNameKey = string.Empty;

            _oposCO = null;

            _minimumDataLength = 1;
            _prefixLength = 0;
            _prefixData = new byte[0];
            _symbologyMode = SymbologyMode.None;
            _suffixLength = 0;
            _suffixData = new byte[0];

            _claimManager = null;

            _releaseClosing = false;

            _dataAccessLock = new object();
            _eventControlMutex = new Mutex(false);
            _eventListMutex = new Mutex(false);
            _eventReactivateEvent = new EventWaitHandle(false, EventResetMode.ManualReset);
            _oposEvents = new List<OposEventScanner>();
            _killEventThread = false;

            _serialPort = null;
            _portBuffer = new byte[0];
            _portBufferLock = new object();
            _portReceivedEvent = new EventWaitHandle(false, EventResetMode.ManualReset);

            _killParseThread = false;
        }

        private bool _disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    if (_opened)
                    {
                        CloseService();
                    }
                    _portReceivedEvent.Dispose();
                    _eventReactivateEvent.Dispose();
                    _eventListMutex.Dispose();
                    _eventControlMutex.Dispose();
                    _dataAccessLock = null;
                    _portBufferLock = null;
                }

                _disposedValue = true;
            }
        }

        ~CSScannerSO()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Support Constructer / Destructer

        #region Change ThreadingModel registry

        [ComRegisterFunction]
        private static void CSCOMRegister(Type registerType)
        {
            if (registerType != typeof(CSScannerSO))
            {
                return;
            }

            using (RegistryKey clsidKey = Registry.ClassesRoot.OpenSubKey("CLSID"))
            {
                using (RegistryKey guidKey = clsidKey.OpenSubKey(registerType.GUID.ToString("B"), true))
                {
                    using (RegistryKey inproc = guidKey.OpenSubKey("InprocServer32", true))
                    {
                        inproc.SetValue("ThreadingModel", "Apartment", RegistryValueKind.String);
                    }
                }
            }
        }

        [ComUnregisterFunction]
        private static void CSCOMUnregister(Type registerType)
        {
            if (registerType != typeof(CSScannerSO))
            {
                return;
            }
        }

        #endregion Change ThreadingModel registry

        #region OPOS ServiceObject Device Common Method

        public int GetPropertyNumber(int PropIndex)
        {
            int value = 0;

            switch (PropIndex)
            {
                case (int)OPOS_Internals.PIDX_CapPowerReporting:
                    value = _capPowerReporting;
                    break;

                case (int)OPOS_Internals.PIDX_CapCompareFirmwareVersion:
                    value = Convert.ToInt32(_capCompareFirmwareVersion);
                    break;

                case (int)OPOS_Internals.PIDX_CapUpdateFirmware:
                    value = Convert.ToInt32(_capUpdateFirmware);
                    break;

                case (int)OPOS_Internals.PIDX_CapStatisticsReporting:
                    value = Convert.ToInt32(_capStatisticsReporting);
                    break;

                case (int)OPOS_Internals.PIDX_CapUpdateStatistics:
                    value = Convert.ToInt32(_capUpdateStatistics);
                    break;

                case (int)OPOS_Internals.PIDX_AutoDisable:
                    value = Convert.ToInt32(_autoDisable);
                    break;

                case (int)OPOS_Internals.PIDX_BinaryConversion:
                    value = _binaryConversion;
                    break;

                case (int)OPOS_Internals.PIDX_Claimed:
                    value = Convert.ToInt32(_claimed);
                    break;

                case (int)OPOS_Internals.PIDX_DataCount:
                    value = _dataCount;
                    break;

                case (int)OPOS_Internals.PIDX_DataEventEnabled:
                    value = Convert.ToInt32(_dataEventEnabled);
                    break;

                case (int)OPOS_Internals.PIDX_DeviceEnabled:
                    value = Convert.ToInt32(_deviceEnabled);
                    break;

                case (int)OPOS_Internals.PIDX_FreezeEvents:
                    value = Convert.ToInt32(_freezeEvents);
                    break;

                //case (int)OPOS_Internals.PIDX_OutputID:
                //    value = _outputID;
                //    break;

                case (int)OPOS_Internals.PIDX_PowerNotify:
                    value = _powerNotify;
                    break;

                case (int)OPOS_Internals.PIDX_PowerState:
                    value = _powerState;
                    break;

                case (int)OPOS_Internals.PIDX_ResultCode:
                    value = _resultCode;
                    break;

                case (int)OPOS_Internals.PIDX_ResultCodeExtended:
                    value = _resultCodeExtended;
                    break;

                case (int)OPOS_Internals.PIDX_State:
                    value = _state;
                    break;

                case (int)OPOS_Internals.PIDX_ServiceObjectVersion:
                    value = _serviceObjectVersion;
                    break;

                case (int)OPOSScannerInternals.PIDXScan_DecodeData:
                    value = Convert.ToInt32(_decodeData);
                    break;

                case (int)OPOSScannerInternals.PIDXScan_ScanDataType:
                    value = _scanDataType;
                    break;
            }

            return value;
        }

        public void SetPropertyNumber(int PropIndex, int Number)
        {
            int result = (int)OPOS_Constants.OPOS_SUCCESS;
            _resultCodeExtended = 0;

            switch (PropIndex)
            {
                case (int)OPOS_Internals.PIDX_AutoDisable:
                    _autoDisable = Convert.ToBoolean(Number);
                    break;

                case (int)OPOS_Internals.PIDX_BinaryConversion:
                    _binaryConversion = Number;
                    break;

                case (int)OPOS_Internals.PIDX_DataEventEnabled:
                    try
                    {
                        _eventControlMutex.WaitOne();
                        _dataEventEnabled = Convert.ToBoolean(Number);
                        if (_dataEventEnabled && _deviceEnabled && !_freezeEvents && !_coFreezeEvents)
                        {
                            ReactivateEventLoop();
                        }
                    }
                    catch { }
                    finally
                    {
                        _eventControlMutex.ReleaseMutex();
                    }
                    break;

                case (int)OPOS_Internals.PIDX_DeviceEnabled:
                    if (!_claimed)
                    {
                        result = (int)OPOS_Constants.OPOS_E_NOTCLAIMED;
                        break;
                    }
                    try
                    {
                        _eventControlMutex.WaitOne();
                        _deviceEnabled = Convert.ToBoolean(Number);
                        if (_dataEventEnabled && _deviceEnabled && !_freezeEvents && !_coFreezeEvents)
                        {
                            ReactivateEventLoop();
                        }
                    }
                    catch { }
                    finally
                    {
                        _eventControlMutex.ReleaseMutex();
                    }
                    break;

                case (int)OPOS_Internals.PIDX_FreezeEvents:
                    try
                    {
                        _eventControlMutex.WaitOne();
                        _freezeEvents = Convert.ToBoolean(Number);
                        if (_dataEventEnabled && _deviceEnabled && !_freezeEvents && !_coFreezeEvents)
                        {
                            ReactivateEventLoop();
                        }
                    }
                    catch { }
                    finally
                    {
                        _eventControlMutex.ReleaseMutex();
                    }
                    break;

                case (int)OPOS_Internals.PIDX_PowerNotify:
                    if (_deviceEnabled || (_capPowerReporting == (int)OPOS_Constants.OPOS_PR_NONE))
                    {
                        result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
                        break;
                    }
                    _powerNotify = Number;
                    break;

                case (int)OPOSScannerInternals.PIDXScan_DecodeData:
                    _decodeData = Convert.ToBoolean(Number);
                    break;
            }

            _resultCode = result;
            return;
        }

        public string GetPropertyString(int PropIndex)
        {
            string value = string.Empty;

            switch (PropIndex)
            {
                case (int)OPOS_Internals.PIDX_CheckHealthText:
                    value = string.Copy(_checkHealthText);
                    break;

                case (int)OPOS_Internals.PIDX_ServiceObjectDescription:
                    value = string.Copy(_serviceObjectDescription);
                    break;

                case (int)OPOS_Internals.PIDX_DeviceDescription:
                    value = string.Copy(_deviceDescription);
                    break;

                case (int)OPOS_Internals.PIDX_DeviceName:
                    value = string.Copy(_deviceName);
                    break;

                case (int)OPOSScannerInternals.PIDXScan_ScanData:
                    lock (_dataAccessLock)
                    {
                        value = SOCommon.ToStringFromByteArray(_scanData, _binaryConversion);
                    }
                    break;

                case (int)OPOSScannerInternals.PIDXScan_ScanDataLabel:
                    lock (_dataAccessLock)
                    {
                        value = SOCommon.ToStringFromByteArray(_scanDataLabel, _binaryConversion);
                    }
                    break;
            }

            return string.Copy(value);
        }

        public void SetPropertyString(int PropIndex, string String)
        {
            int result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            _resultCode = result;
            return;
        }

        public int OpenService(string DeviceClass, string DeviceName, object pDispatch)
        {
            int result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;

            if (_opened)
            {
                _resultCode = result;    // Illegal=Multiple open error
                return result;
            }

            string registrypath = OposRegKey.OPOS_ROOTKEY + "\\" + DeviceClass + "\\" + DeviceName;
            _openResult = 0;

            result = InitializeFromRegistry(registrypath);
            if (result != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                _openResult = (int)OPOS_Constants.OPOS_ORS_CONFIG;
                _resultCode = result;    // NoExist/Failure/Illegal error
                return result;
            }

            _portBuffer = new byte[0];

            _serialPort = new SerialPort();
            _serialPort.PortName = "COM3";
            _serialPort.BaudRate = 115200;
            _serialPort.DataBits = 8;
            _serialPort.StopBits = StopBits.One;
            _serialPort.Parity = Parity.None;
            _serialPort.DiscardNull = false;
            _serialPort.DtrEnable = true;
            _serialPort.RtsEnable = true;
            _serialPort.Handshake = Handshake.None;
            _serialPort.NewLine = "\r";
            _serialPort.ParityReplace = 0;
            _serialPort.ReadBufferSize = 8192;
            _serialPort.ReadTimeout = 100;
            _serialPort.ReceivedBytesThreshold = 1;
            _serialPort.WriteBufferSize = 8192;
            _serialPort.WriteTimeout = 100;
            result = SOCommon.GetSerialConfigFromRegistry(ref _serialPort, registrypath + @"\PortConfig");
            if (result != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                _serialPort.Dispose();
                _serialPort = null;
                _portBuffer = null;
                _openResult = (int)OPOS_Constants.OPOS_ORS_CONFIG;
                _resultCode = result;    // NoExist/Failure/Illegal error
                return result;
            }
            _serialPort.DataReceived += new SerialDataReceivedEventHandler(_serialPort_DataReceived);
            _serialPort.ErrorReceived += new SerialErrorReceivedEventHandler(_serialPort_ErrorReceived);
            _serialPort.PinChanged += new SerialPinChangedEventHandler(_serialPort_PinChanged);

            _deviceName = DeviceName;
            _claimManager = new ClaimManager(DeviceClass, DeviceName);
            _oposCO = pDispatch;

            _killEventThread = false;
            _eventThreadStart = this.DoEventThread;
            _eventThread = new Thread(_eventThreadStart);
            _eventThread.IsBackground = true;

            _killParseThread = false;
            _parseThreadStart = this.DoParseThread;
            _parseThread = new Thread(_parseThreadStart);
            _parseThread.IsBackground = true;


            _opened = true;
            _resultCode = (int)OPOS_Constants.OPOS_SUCCESS;
            return _resultCode;
        }

        [ComVisible(false)]
        private int InitializeFromRegistry(string key)
        {
            int result = (int)OPOS_Constants.OPOS_E_NOEXIST;

            using (RegistryKey propKey = Registry.LocalMachine.OpenSubKey(key + @"\PropertyCommon"))
            {
                _capCompareFirmwareVersion = SOCommon.GetBoolFromRegistry(propKey, "CapCompareFirmwareVersion", _capCompareFirmwareVersion);
                _capUpdateFirmware = SOCommon.GetBoolFromRegistry(propKey, "CapUpdateFirmware", _capUpdateFirmware);
                _capPowerReporting = SOCommon.GetIntegerFromRegistry(propKey, "CapPowerReporting", _capPowerReporting);
                _capStatisticsReporting = SOCommon.GetBoolFromRegistry(propKey, "CapStatisticsReporting", _capStatisticsReporting);
                _capUpdateStatistics = SOCommon.GetBoolFromRegistry(propKey, "CapUpdateStatistics", _capUpdateStatistics);
                _serviceObjectDescription = SOCommon.GetStringFromRegistry(propKey, "ServiceObjectDescription", _serviceObjectDescription);
                _serviceObjectVersion = SOCommon.GetIntegerFromRegistry(propKey, "ServiceObjectVersion", _serviceObjectVersion);
                _deviceDescription = SOCommon.GetStringFromRegistry(propKey, "DeviceDescription", _deviceDescription);
                result = (int)OPOS_Constants.OPOS_SUCCESS;
            }

            using (RegistryKey propKey = Registry.LocalMachine.OpenSubKey(key + @"\PropertySpecific"))
            {
                string strWork = "None";
                strWork = SOCommon.GetStringFromRegistry(propKey, "SymbologyMode", strWork);
                switch (strWork.ToUpper())
                {
                    case "NONE":
                        _symbologyMode = SymbologyMode.None;
                        break;
                    case "AIM":
                        _symbologyMode = SymbologyMode.AIM;
                        break;
                    case "HONEYWELL":
                        _symbologyMode = SymbologyMode.Honeywell;
                        break;
                }
                _prefixData = SOCommon.GetByteArrayFromRegistry(propKey, "PrefixData", _prefixData);
                _prefixLength = _prefixData.Length;
                if (_prefixLength > 0)
                {
                    _suffixData = SOCommon.GetByteArrayFromRegistry(propKey, "SuffixData", _suffixData);
                }
                _suffixLength = _suffixData.Length;
                result = (int)OPOS_Constants.OPOS_SUCCESS;
            }

            _minimumDataLength = 1 + _prefixLength + _suffixLength;
            switch (_symbologyMode)
            {
                case SymbologyMode.AIM:
                    _minimumDataLength += 3;
                    break;
                case SymbologyMode.Honeywell:
                    _minimumDataLength += 1;
                    break;
            }
            return result;
        }

        public int CloseService()
        {
            if (!_opened)
            {
                _resultCode = (int)OPOS_Constants.OPOS_E_CLOSED;
                return _resultCode;
            }

            _serialPort.DataReceived -= new SerialDataReceivedEventHandler(_serialPort_DataReceived);
            _serialPort.ErrorReceived -= new SerialErrorReceivedEventHandler(_serialPort_ErrorReceived);
            _serialPort.PinChanged -= new SerialPinChangedEventHandler(_serialPort_PinChanged);

            _portBuffer = null;

            _serialPort.Dispose();
            _serialPort = null;

            _eventThreadStart = null;
            _claimManager.Dispose();
            _claimManager = null;

            _opened = false;
            _resultCode = (int)OPOS_Constants.OPOS_SUCCESS;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ClaimDevice(int Timeout)
        {
            _resultCodeExtended = 0;

            if (_claimed)
            {
                _resultCode = (int)OPOS_Constants.OPOS_SUCCESS;
                return _resultCode;
            }

            _resultCode = _claimManager.ClaimDevice(Timeout);
            if (_resultCode == (int)OPOS_Constants.OPOS_SUCCESS)
            {
                _claimed = true;

                _killEventThread = false;
                _eventThread.Start();
                _killParseThread = false;
                _parseThread.Start();
            }

            try
            {
                _serialPort.Open();
            }
            catch { }

            return _resultCode;
        }

        public int ReleaseDevice()
        {
            _resultCodeExtended = 0;

            if (!_claimed)
            {
                _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
                return _resultCode;
            }

            if (_serialPort.IsOpen)
            {
                try
                {
                    _serialPort.Close();
                }
                catch { }
            }

            _killParseThread = true;
            _parseThread.Join();
            _killEventThread = true;
            _eventThread.Join();
            int result = _claimManager.ReleaseDevice();
            _claimed = false;
            _resultCode = (int)OPOS_Constants.OPOS_SUCCESS;
            return _resultCode;
        }

        public int CheckHealth(int Level)
        {
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed, _deviceEnabled)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }

            switch (Level)
            {
                case (int)OPOS_Constants.OPOS_CH_INTERNAL:
                    _checkHealthText = "Internal CheckHealth successfull";
                    break;

                case (int)OPOS_Constants.OPOS_CH_EXTERNAL:
                    _checkHealthText = "External CheckHealth successfull";
                    break;

                case (int)OPOS_Constants.OPOS_CH_INTERACTIVE:
                    _checkHealthText = "Interactive CheckHealth successfull";
                    break;

                default:
                    _checkHealthText = "The Level parameter is illegal value.";
                    _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
                    break;
            }

            return _resultCode;
        }

        private const int clearInputEvent = (int)OposEvent.EVENT_DATA | (int)OposEvent.EVENT_ERROR;

        public int ClearInput()
        {
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }
            ClearEventList(clearInputEvent);
            return _resultCode;
        }

        public int ClearInputProperties()
        {
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }
            lock (_dataAccessLock)
            {
                _scanData = new byte[0];
                _scanDataLabel = new byte[0];
                _scanDataType = (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN;
            }
            return _resultCode;
        }

        public int DirectIO(int Command, ref int pData, ref string pStrIng)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CompareFirmwareVersion(string FirmwareFileName, out int pResult)
        {
            pResult = 0;
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed, _deviceEnabled)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }
            return _resultCode;
        }

        public int UpdateFirmware(string FirmwareFileName)
        {
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed, _deviceEnabled, _capCompareFirmwareVersion)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }
            return _resultCode;
        }

        public int ResetStatistics(string StatisticsBuffer)
        {
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed, _deviceEnabled, _capUpdateStatistics)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }
            return _resultCode;
        }

        public int RetrieveStatistics(ref string pStatisticsBuffer)
        {
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed, _deviceEnabled, _capStatisticsReporting)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }
            return _resultCode;
        }

        public int UpdateStatistics(string StatisticsBuffer)
        {
            _resultCodeExtended = 0;
            if ((_resultCode = SOCommon.VerifyStateCapability(_claimed, _deviceEnabled, _capUpdateStatistics)) != (int)OPOS_Constants.OPOS_SUCCESS)
            {
                return _resultCode;
            }
            return _resultCode;
        }

        public int COFreezeEvents(bool Freeze)
        {
            _resultCode = (int)OPOS_Constants.OPOS_SUCCESS;
            _resultCodeExtended = 0;
            try
            {
                _eventControlMutex.WaitOne();
                _coFreezeEvents = Freeze;
                if (_dataEventEnabled && _deviceEnabled && !_freezeEvents && !_coFreezeEvents)
                {
                    ReactivateEventLoop();
                }
            }
            catch { }
            finally
            {
                _eventControlMutex.ReleaseMutex();
            }
            return _resultCode;
        }

        public int GetOpenResult()
        {
            return _openResult;
        }

        #endregion OPOS ServiceObject Device Common Method

        #region SerialPort I/O

        [ComVisible(false)]
        private void _serialPort_PinChanged(object sender, SerialPinChangedEventArgs e)
        {
            SerialPort portObject = (SerialPort)sender;
            if (portObject.IsOpen)
            {
            }
        }

        private const int inputErrors = (int)SerialError.Frame | (int)SerialError.Overrun | (int)SerialError.RXOver | (int)SerialError.RXParity;
        [ComVisible(false)]
        private void _serialPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            SerialPort portObject = (SerialPort)sender;
            if (!portObject.IsOpen)
            {
                return;
            }
            if (((int)e.EventType & inputErrors) != 0)
            {
                _serialPort_ReadData(ref portObject);
                portObject.DiscardInBuffer();
            }
            if (((int)e.EventType & (int)SerialError.TXFull) != 0)
            {
                portObject.DiscardOutBuffer();
            }
        }

        [ComVisible(false)]
        private void _serialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort portObject = (SerialPort)sender;
            if (!portObject.IsOpen)
            {
                return;
            }
            _serialPort_ReadData(ref portObject);
        }

        [ComVisible(false)]
        private void _serialPort_ReadData(ref SerialPort portObject)
        {
            int bytes = portObject.BytesToRead;
            if (bytes > 0)
            {
                byte[] tempbuffer = new byte[bytes];
                portObject.Read(tempbuffer, 0, bytes);
                lock (_portBufferLock)
                {
                    int currentsize = _portBuffer.Length;
                    Array.Resize<byte>(ref _portBuffer, (currentsize + bytes));
                    Buffer.BlockCopy(tempbuffer, 0, _portBuffer, currentsize, bytes);
                    _portReceivedEvent.Set();
                }
            }
        }

        [ComVisible(false)]
        internal void DoParseThread()
        {
            while (_killParseThread == false)
            {
                if (_portReceivedEvent.WaitOne(OposEvent.EVENT_CYCLE_TIMEOUT))
                {
                    _portReceivedEvent.Reset();
                }
                if (_killParseThread)
                {
                    break;
                }
                OposEventScanner pEvent = null;
                lock (_portBufferLock)
                {
                    if (_portBuffer.Length >= _minimumDataLength)
                    {
                        pEvent = new OposEventScanner();
                        pEvent.eventType = OposEvent.EVENT_DATA;
                        pEvent._scanData = (byte[])_portBuffer.Clone();
                        _portBuffer = new byte[0];
                        try
                        {
                            _eventListMutex.WaitOne(OposEvent.EVENT_MUTEX_TIMEOUT);
                            _oposEvents.Add(pEvent);
                            _eventReactivateEvent.Set();
                        }
                        finally
                        {
                            _eventListMutex.ReleaseMutex();
                        }
                    }
                }
            }
        }

        [ComVisible(false)]
        internal OposEventScanner CheckExtractBarcode()
        {
            OposEventScanner pEvent = null;
            int indexPrefix = 0;
            int indexSymbology = 0;
            int indexBarcode = 0;
            int indexSuffix = 0;
            int barcodeSymbology = (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN;
            int bufferLength = _portBuffer.Length;
            int startIndex = 0;
            if (_prefixLength > 0)
            {
                for (; ; )
                {
                    indexPrefix = Array.IndexOf(_portBuffer, _prefixData[0], startIndex);
                    if ((indexPrefix < 0)
                            || (bufferLength - indexPrefix < _prefixLength))
                    {
                        break;
                    }
                    bool foundPrefix = true;
                    for (int i = 1; i < _prefixLength; i++)
                    {
                        if (_portBuffer[indexPrefix + i] != _prefixData[i])
                        {
                            foundPrefix = false;
                            break;
                        }
                    }
                    if (!foundPrefix)
                    {
                    }
                    else if (indexPrefix + _minimumDataLength > bufferLength)
                    {
                    }
                    else
                    {
                        indexSymbology = indexBarcode = indexPrefix + _prefixLength;
                    }
                }
            }
            if (_symbologyMode == SymbologyMode.AIM)
            {
                indexSymbology = Array.IndexOf(_portBuffer, ']', indexSymbology);
                barcodeSymbology = s_SymbologyAIM[_portBuffer[indexSymbology + 1]][_portBuffer[indexSymbology + 2]];
                indexBarcode = indexSymbology + 3;
            }
            else if (_symbologyMode == SymbologyMode.Honeywell)
            {
                barcodeSymbology = s_SymbologyHoneywell[_portBuffer[indexSymbology]];
                indexBarcode = indexSymbology + 1;
            }

            if (_suffixLength > 0)
            {
                for (; ; )
                {
                    indexSuffix = Array.IndexOf(_portBuffer, _suffixData[0], indexBarcode);
                    if ((indexSuffix < 0)
                            || (bufferLength - indexSuffix < _suffixLength))
                    {
                        break;
                    }
                    bool foundSuffix = true;
                    for (int i = 1; i < _suffixLength; i++)
                    {
                        if (_portBuffer[indexPrefix + i] != _suffixData[i])
                        {
                            foundSuffix = false;
                            break;
                        }
                    }
                    if (!foundSuffix || (indexSuffix + _minimumDataLength > bufferLength))
                    {
                    }
                    else
                    {
                        indexSymbology = indexBarcode = indexPrefix + _prefixLength;
                    }
                }
            }
            return pEvent;
        }

        private static Dictionary<byte, Dictionary<byte, int>> s_SymbologyAIM = new Dictionary<byte, Dictionary<byte, int>>()
        {
            {
                /*'A'*/0x41, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_Code39 },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_Code39_CK },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_Code39 },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_Code39 },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_Code39_CK },
                    { /*'7'*/0x37, (int)OPOSScannerConstants.SCAN_SDT_Code39 }
                }
            },
            {
                /*'B'*/0x42, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_TELEPEN },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_TELEPEN },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_TELEPEN },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_TELEPEN }
                }
            },
            {
                /*'C'*/0x43, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_Code128 },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_EAN128 },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_Code128 },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_Code128 }
                }
            },
            {
                /*'D'*/0x44, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // Code One  undefined UPOS
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }
                }
            },
            {
                /*'E'*/0x45, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_EAN13 },  // EAN/JAN-13, UPC-A/E
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_EAN13_S },  // AddOn 2 only
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_EAN13_S },  // AddOn 5 only
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_EAN13_S },  // EAN/JAN-13 AddOn, UPC-A/E AddOn
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_EAN8 }
                }
            },
            {
                /*'F'*/0x46, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_Codabar },  //
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_Codabar },  // American Blood Commission
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_Codabar },  // included check digit
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_Codabar }   // stripped check digit
                }
            },
            {
                /*'G'*/0x47, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_Code93 },  // Code93 only
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_Code93 },  // Code93i optional
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'7'*/0x37, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'8'*/0x38, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'9'*/0x39, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'A'*/0x41, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'B'*/0x42, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'C'*/0x43, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'D'*/0x44, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'E'*/0x45, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'F'*/0x46, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'G'*/0x47, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'H'*/0x48, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'I'*/0x49, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'J'*/0x4A, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'K'*/0x4B, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'L'*/0x4C, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'M'*/0x4D, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'N'*/0x4E, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'O'*/0x4F, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'P'*/0x50, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'Q'*/0x51, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'R'*/0x52, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'S'*/0x53, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'T'*/0x54, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'U'*/0x55, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'V'*/0x56, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'W'*/0x57, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'X'*/0x58, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'Y'*/0x59, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'Z'*/0x5A, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'a'*/0x61, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'b'*/0x62, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'c'*/0x63, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'d'*/0x64, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'e'*/0x65, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'f'*/0x66, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'g'*/0x67, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'h'*/0x68, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'i'*/0x69, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'j'*/0x6A, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'k'*/0x6B, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'l'*/0x6C, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
                    { /*'m'*/0x6D, (int)OPOSScannerConstants.SCAN_SDT_Code93 }
                }
            },
            {
                /*'H'*/0x48, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_Code11 },  // single modulo 11 check character
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_Code11 },  // two modulo 11 check characters
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_Code11 }   // stripped check character(s)
                }
            },
            {
                /*'I'*/0x49, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_ITF },  // no check digit
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_ITF_CK },  // included check digit
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_ITF }   // stripped check digit
                }
            },
            // { /*'J'*/0x4A, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            {
                /*'K'*/0x4B, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_Code16k },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_Code16k },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_Code16k },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_Code16k }
                }
            },
            {
                /*'L'*/0x4C, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_PDF417 },  // PDF417, MicroPDF417
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_PDF417 },  // PDF417, MicroPDF417
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_PDF417 },  // PDF417, MicroPDF417
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_UPDF417 },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_UPDF417 },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_UPDF417 }
                }
            },
            {
                /*'M'*/0x4D, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_MSI },  // included check character
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_MSI }   // stripped check character
                }
            },
            {
                /*'N'*/0x4E, new Dictionary<byte, int>()
                {
                    {
                        /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN
                    }   // Anker Code
                }
            },
            {
                /*'O'*/0x4F, new Dictionary<byte, int>()
                {
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_CodablockF },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_CodablockF },
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_CodablockA }
                }
            },
            {
                /*'P'*/0x50, new Dictionary<byte, int>()
                {
                    {
                        /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_PLESSEY
                    }
                }
            },
            {
                /*'Q'*/0x51, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_QRCODE },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_QRCODE },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_QRCODE },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_QRCODE },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_QRCODE },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_QRCODE },
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_QRCODE }
                }
            },
            {
                /*'R'*/0x52, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_TF },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_TF },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_TF }
                }
            },
            {
                /*'S'*/0x53, new Dictionary<byte, int>()
                {
                    {
                        /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_TF
                    }
                }
            },
            {
                /*'T'*/0x54, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_Code49 },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_Code49 },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_Code49 },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_Code49 }
                }
            },
            {
                /*'U'*/0x55, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_MAXICODE },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_MAXICODE },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_MAXICODE },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_MAXICODE }
                }
            },
            // { /*'V'*/0x56, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'W'*/0x57, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            {
                /*'X'*/0x58, new Dictionary<byte, int>()  // manufacturer difined
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // Honeywell: Code 32 Pharmaceutical (PARAF), China Post, Matrix 2 of 5, NEC 2 of 5,
                    // UPC-E1, Chinese Sensible Code, Australian Post, British Post, Canadian Post, China Post, InfoMail, Intelligent Mail Bar Code, Japanese Post, KIX (Netherlands) Post, Korea Post, Planet Code, Postal-4i, Postnet

                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'7'*/0x37, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'8'*/0x38, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'9'*/0x39, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'A'*/0x41, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'B'*/0x42, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'C'*/0x43, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'D'*/0x44, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'E'*/0x45, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'F'*/0x46, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }
                }
            },
            // { /*'Y'*/0x59, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            {
                /*'Z'*/0x5A, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // keyboard
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // magnetic stripe
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // radio frequency tag
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // manufactureer defined
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'7'*/0x37, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'8'*/0x38, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'9'*/0x39, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'A'*/0x41, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'B'*/0x42, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'C'*/0x43, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'D'*/0x44, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'E'*/0x45, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'F'*/0x46, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }
                }
            },
            // { /*'a'*/0x61, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'b'*/0x62, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            {
                /*'c'*/0x63, new Dictionary<byte, int>()
                {
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // Channel Code  undefined UPOS
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // Channel Code
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // Channel Code
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // Channel Code
                    { /*'7'*/0x37, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // Channel Code
                    { /*'8'*/0x38, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // Channel Code
                    { /*'9'*/0x39, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // Channel Code
                }
            },
            {
                /*'d'*/0x64, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
                }
            },
            {
                /*'e'*/0x65, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_GS1DATABAR },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_GS1DATABAR },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_GS1DATABAR },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_GS1DATABAR },
                }
            },
            // { /*'f'*/0x66, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'g'*/0x67, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'h'*/0x68, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'i'*/0x69, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'j'*/0x6A, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'k'*/0x6B, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'l'*/0x6C, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'m'*/0x6D, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'n'*/0x6E, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            {
                /*'o'*/0x6F, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // OCR Undefined Font  undefined UPOS
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_OCRA }, // OCR-A Font
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_OCRB }, // OCR-B Font
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN }, // OCR Other Font  undefined UPOS
                }
            },
            // { /*'p'*/0x70, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'q'*/0x71, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'r'*/0x72, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            {
                /*'s'*/0x73, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // Super Code  undefined UPOS
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },
                }
            },
            // { /*'t'*/0x74, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'u'*/0x75, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'v'*/0x76, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'w'*/0x77, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'x'*/0x78, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            // { /*'y'*/0x79, new Dictionary<byte, int>(){ { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN } } }, // reserve
            {
                /*'z'*/0x7A, new Dictionary<byte, int>()
                {
                    { /*'0'*/0x30, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'1'*/0x31, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'2'*/0x32, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'3'*/0x33, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'4'*/0x34, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'5'*/0x35, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'6'*/0x36, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'7'*/0x37, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'8'*/0x38, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'9'*/0x39, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'A'*/0x41, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'B'*/0x42, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
                    { /*'C'*/0x43, (int)OPOSScannerConstants.SCAN_SDT_AZTEC }
                }
            }
        };

        private static byte[] a_SpecialHoneywell = { 0x44, 0x45, 0x63, 0x64 };

        private static Dictionary<byte, int> s_SymbologyHoneywell = new Dictionary<byte, int>()
        {
            { /*','*/0x2C, (int)OPOSScannerConstants.SCAN_SDT_InfoMail },
            { /*'<'*/0x3C, (int)OPOSScannerConstants.SCAN_SDT_Code32 },
            { /*'?'*/0x3F, (int)OPOSScannerConstants.SCAN_SDT_KoreanPost },
            { /*'A'*/0x41, (int)OPOSScannerConstants.SCAN_SDT_AusPost },
            { /*'B'*/0x42, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // British Post  undefined UPOS
            { /*'C'*/0x43, (int)OPOSScannerConstants.SCAN_SDT_CanPost },
            { /*'D'*/0x44, (int)OPOSScannerConstants.SCAN_SDT_EAN8 },  // EAN/JAN-8, EAN/JAN-8 AddOn
            { /*'E'*/0x45, (int)OPOSScannerConstants.SCAN_SDT_UPCE },  // UPC-E, UPC-E AddOn
            // { /*'F'*/0x46, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            // { /*'G'*/0x47, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'H'*/0x48, (int)OPOSScannerConstants.SCAN_SDT_HANXIN },  // Chinese Sensible Code
            { /*'I'*/0x49, (int)OPOSScannerConstants.SCAN_SDT_EAN128 },
            { /*'J'*/0x4A, (int)OPOSScannerConstants.SCAN_SDT_JapanPost },
            { /*'K'*/0x4B, (int)OPOSScannerConstants.SCAN_SDT_DutchKix },
            { /*'L'*/0x4C, (int)OPOSScannerConstants.SCAN_SDT_UsPlanet },
            { /*'M'*/0x4D, (int)OPOSScannerConstants.SCAN_SDT_UsIntelligent },
            { /*'N'*/0x4E, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // Postal-4i
            // { /*'O'*/0x4F, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'P'*/0x50, (int)OPOSScannerConstants.SCAN_SDT_PostNet },
            { /*'Q'*/0x51, (int)OPOSScannerConstants.SCAN_SDT_ChinaPost },
            { /*'R'*/0x52, (int)OPOSScannerConstants.SCAN_SDT_UPDF417 },
            // { /*'S'*/0x53, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'T'*/0x54, (int)OPOSScannerConstants.SCAN_SDT_TLC39 },
            // { /*'U'*/0x55, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'V'*/0x56, (int)OPOSScannerConstants.SCAN_SDT_CodablockA },
            // { /*'W'*/0x57, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            // { /*'X'*/0x58, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'Y'*/0x59, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // NEC 2 of 5 = alternatively Matrix 2 of 5  undefined UPOS
            // { /*'Z'*/0x5A, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'a'*/0x61, (int)OPOSScannerConstants.SCAN_SDT_Codabar },
            { /*'b'*/0x62, (int)OPOSScannerConstants.SCAN_SDT_Code39 },
            { /*'c'*/0x63, (int)OPOSScannerConstants.SCAN_SDT_UPCA },  // UPC-A, UPC-A AddOn
            { /*'d'*/0x64, (int)OPOSScannerConstants.SCAN_SDT_EAN13 },  // EAN/JAN-13, EAN/JAN-13 AddOn
            { /*'e'*/0x65, (int)OPOSScannerConstants.SCAN_SDT_ITF },
            { /*'f'*/0x66, (int)OPOSScannerConstants.SCAN_SDT_TF },
            { /*'g'*/0x67, (int)OPOSScannerConstants.SCAN_SDT_MSI },
            { /*'h'*/0x68, (int)OPOSScannerConstants.SCAN_SDT_Code11 },
            { /*'i'*/0x69, (int)OPOSScannerConstants.SCAN_SDT_Code93 },
            { /*'j'*/0x6A, (int)OPOSScannerConstants.SCAN_SDT_Code128 },
            // { /*'k'*/0x6B, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'l'*/0x6C, (int)OPOSScannerConstants.SCAN_SDT_Code49 },
            { /*'m'*/0x6D, (int)OPOSScannerConstants.SCAN_SDT_TFMAT },
            // { /*'n'*/0x6E, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            // { /*'o'*/0x6F, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            // { /*'p'*/0x70, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'q'*/0x71, (int)OPOSScannerConstants.SCAN_SDT_CodablockF },
            { /*'r'*/0x72, (int)OPOSScannerConstants.SCAN_SDT_PDF417 },
            { /*'s'*/0x73, (int)OPOSScannerConstants.SCAN_SDT_QRCODE },  // QRCODE, MicroQRCODE
            { /*'t'*/0x74, (int)OPOSScannerConstants.SCAN_SDT_TELEPEN },
            // { /*'u'*/0x75, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            // { /*'v'*/0x76, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'w'*/0x77, (int)OPOSScannerConstants.SCAN_SDT_DATAMATRIX },
            // { /*'x'*/0x78, (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN },  // nothing
            { /*'y'*/0x79, (int)OPOSScannerConstants.SCAN_SDT_GS1DATABAR },  // GS1 DataBar, GS1 Composite
            { /*'z'*/0x7A, (int)OPOSScannerConstants.SCAN_SDT_AZTEC },
            { /*'{'*/0x7B, (int)OPOSScannerConstants.SCAN_SDT_GS1DATABAR_TYPE2 },  // GS1 DataBar Limited
            { /*'}'*/0x7D, (int)OPOSScannerConstants.SCAN_SDT_GS1DATABAR_E }
        };

        #endregion SerialPort I/O

        #region EventThread

        [ComVisible(false)]
        internal void DoEventThread()
        {
            while (_killEventThread == false)
            {
                _eventListMutex.WaitOne(OposEvent.EVENT_MUTEX_TIMEOUT);
                _eventControlMutex.WaitOne(OposEvent.EVENT_MUTEX_TIMEOUT);

                if ((_releaseClosing == false)
                        && (_freezeEvents == false)
                        && (_coFreezeEvents == false)
                        && (FindNotifiableEvent()))
                {
                    _eventControlMutex.ReleaseMutex();

                    OposEventScanner scannerEvent = _oposEvents[0];
                    _oposEvents.RemoveAt(0);

                    _eventListMutex.ReleaseMutex();

                    switch (scannerEvent.eventType)
                    {
                        case OposEvent.EVENT_DATA:

                            lock (_dataAccessLock)
                            {
                                _scanData = (byte[])scannerEvent._scanData.Clone();
                                if (_decodeData)
                                {
                                    _scanDataType = scannerEvent._scanDataType;
                                    _scanDataLabel = (byte[])scannerEvent._scanDataLabel.Clone();
                                }
                                else
                                {
                                    _scanDataType = (int)OPOSScannerConstants.SCAN_SDT_UNKNOWN;
                                    _scanDataLabel = new byte[0];
                                }
                                _dataEventEnabled = false;
                            }
                            _oposCO.SOData(scannerEvent.status);
                            break;
                        case OposEvent.EVENT_DIRECTIO:
                            _oposCO.SODirectIO(scannerEvent.eventNumber, ref scannerEvent.pData, ref scannerEvent.pString);
                            // todo:somethingelse
                            break;
                        case OposEvent.EVENT_ERROR:
                            _oposCO.SOError(scannerEvent.resultCode, scannerEvent.resultCodeExtended, scannerEvent.errorLocus, ref scannerEvent.pErrorResponse);
                            if (scannerEvent.pErrorResponse == (int)OPOS_Constants.OPOS_ER_CLEAR)
                            {
                                ClearEventList((OposEvent.EVENT_DATA | OposEvent.EVENT_ERROR));
                            }
                            else
                            {
                                // todo:retry or continueinput
                            }
                            break;
                        case OposEvent.EVENT_STATUSUPDATE:
                            _oposCO.SOStatusUpdate(scannerEvent.data);
                            break;
                    }
                }
                else
                {
                    _eventControlMutex.ReleaseMutex();
                    _eventListMutex.ReleaseMutex();
                    if (_killEventThread == false)
                    {
                        if (_eventReactivateEvent.WaitOne(OposEvent.EVENT_CYCLE_TIMEOUT))
                        {
                            _eventReactivateEvent.Reset();
                        }
                    }
                }
            }
            ClearEventList(OposEvent.EVENT_ALL);
        }

        [ComVisible(false)]
        private bool FindNotifiableEvent()
        {
            bool result = false;
            int eventCount = _oposEvents.Count;
            if (eventCount <= 0)
            {
                return result;
            }
            int eventType = _oposEvents[0].eventType;
            switch (eventType)
            {
                case OposEvent.EVENT_DATA:
                    if (_dataEventEnabled)
                    {
                        result = true;
                    }
                    break;
                case OposEvent.EVENT_DIRECTIO:
                case OposEvent.EVENT_ERROR:
                case OposEvent.EVENT_STATUSUPDATE:
                    result = true;
                    break;
            }
            return result;
        }

        [ComVisible(false)]
        internal void ClearEventList(int flags)
        {
            if (_eventListMutex.WaitOne(OposEvent.EVENT_MUTEX_TIMEOUT))
            {
                for (int i = _oposEvents.Count - 1; i >= 0; i--)
                {
                    if ((_oposEvents[i].eventType & flags) != 0)
                    {
                        OposEventScanner e = _oposEvents[i];
                        _oposEvents.RemoveAt(i);
                        e.Dispose();
                    }
                }
                _eventListMutex.ReleaseMutex();
            }
            return;
        }

        [ComVisible(false)]
        internal void ReactivateEventLoop()
        {
            _eventReactivateEvent.Set();
        }

        #endregion EventThread
    }
}
