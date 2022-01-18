
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
    internal class OposEventSoundRecorder : OposEvent
    {
        internal byte[] _soundData;

        #region IDisposable Support

        public OposEventSoundRecorder()
        {
            _soundData = new byte[0];
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
                        _soundData = null;
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
    [Guid("CABF342C-F23C-4BAC-9945-7F6401806FA2")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.SoundRecorder.OpenPOS.CSSO.CSSoundRecorderSO.1")]
    public class CSSoundRecorderSO : IOPOSSoundRecorderSO, IDisposable
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

        internal string _capAssociatedHardTotalsDevice;
        internal bool _capChannel;
        internal bool _capRecordingLevel;
        internal bool _capSamplingRate;
        internal bool _capSoundType;
        internal int _capStorage;
        internal string _channel;
        internal string _channelList;
        internal int _recordingLevel;
        internal int _remainingRecordingTimeInSec;
        internal string _samplingRate;
        internal string _samplingRateList;
        internal volatile byte[] _soundData;
        internal string _soundType;
        internal string _soundTypeList;
        internal int _storage;

        internal volatile bool _opened;
        internal string _deviceNameKey;

        internal dynamic _oposCO;

        internal ClaimManager _claimManager;

        internal volatile bool _releaseClosing;

        internal volatile Object _dataAccessLock;
        internal volatile Mutex _eventControlMutex;
        internal volatile Mutex _eventListMutex;
        internal volatile EventWaitHandle _eventReactivateEvent;

        internal ThreadStart _eventThreadStart;
        internal Thread _eventThread;
        internal volatile List<OposEventSoundRecorder> _oposEvents;
        internal volatile bool _killEventThread;

        #region IDisposable Support Constructer / Destructer

        public CSSoundRecorderSO()
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

            _capAssociatedHardTotalsDevice = string.Empty;
            _capChannel = false;
            _capRecordingLevel = false;
            _capSamplingRate = false;
            _capSoundType = false;
            _capStorage = 0;
            _channel = string.Empty;
            _channelList = string.Empty;
            _recordingLevel = 0;
            _remainingRecordingTimeInSec = 0;
            _samplingRate = string.Empty;
            _samplingRateList = string.Empty;
            _soundData = new byte[0];
            _soundType = string.Empty;
            _soundTypeList = string.Empty;
            _storage = 0;

            _opened = false;
            _deviceNameKey = string.Empty;

            _oposCO = null;

            _claimManager = null;

            _releaseClosing = false;

            _dataAccessLock = new object();
            _eventControlMutex = new Mutex(false);
            _eventListMutex = new Mutex(false);
            _eventReactivateEvent = new EventWaitHandle(false, EventResetMode.ManualReset);
            _oposEvents = new List<OposEventSoundRecorder>();
            _killEventThread = false;
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
                    _eventReactivateEvent.Dispose();
                    _eventListMutex.Dispose();
                    _eventControlMutex.Dispose();
                    _dataAccessLock = null;
                }

                _disposedValue = true;
            }
        }

        ~CSSoundRecorderSO()
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
            if (registerType != typeof(CSSoundRecorderSO))
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
            if (registerType != typeof(CSSoundRecorderSO))
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

                case (int)OPOSSoundRecorderInternals.PIDXSrec_CapChannel:
                    value = Convert.ToInt32(_capChannel);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_CapRecordingLevel:
                    value = Convert.ToInt32(_capRecordingLevel);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_CapSamplingRate:
                    value = Convert.ToInt32(_capSamplingRate);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_CapSoundType:
                    value = Convert.ToInt32(_capSoundType);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_CapStorage:
                    value = _capStorage;
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_RecordingLevel:
                    value = _recordingLevel;
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_RemainingRecordingTimeInSec:
                    value = _remainingRecordingTimeInSec;
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_Storage:
                    value = _storage;
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

                case (int)OPOSSoundRecorderInternals.PIDXSrec_RecordingLevel:
                    _recordingLevel = Number;
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_RemainingRecordingTimeInSec:
                    _remainingRecordingTimeInSec = Number;
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_Storage:
                    _storage = Number;
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

                case (int)OPOSSoundRecorderInternals.PIDXSrec_CapAssociatedHardTotalsDevice:
                    value = string.Copy(_capAssociatedHardTotalsDevice);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_Channel:
                    value = string.Copy(_channel);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_ChannelList:
                    value = string.Copy(_channelList);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_SamplingRate:
                    value = string.Copy(_samplingRate);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_SamplingRateList:
                    value = string.Copy(_samplingRateList);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_SoundData:
                    lock (_dataAccessLock)
                    {
                        value = SOCommon.ToStringFromByteArray(_soundData, _binaryConversion);
                    }
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_SoundType:
                    value = string.Copy(_soundType);
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_SoundTypeList:
                    value = string.Copy(_soundTypeList);
                    break;

            }

            return string.Copy(value);
        }

        public void SetPropertyString(int PropIndex, string String)
        {
            int result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;

            switch (PropIndex)
            {
                case (int)OPOSSoundRecorderInternals.PIDXSrec_Channel:
                    _channel = String;
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_SamplingRate:
                    _samplingRate = String;
                    break;

                case (int)OPOSSoundRecorderInternals.PIDXSrec_SoundType:
                    _soundType = String;
                    break;

            }

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

            _deviceName = DeviceName;
            _claimManager = new ClaimManager(DeviceClass, DeviceName);
            _oposCO = pDispatch;

            _killEventThread = false;
            _eventThreadStart = this.DoEventThread;
            _eventThread = new Thread(_eventThreadStart);
            _eventThread.IsBackground = true;

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

            return result;
        }

        public int CloseService()
        {
            if (!_opened)
            {
                _resultCode = (int)OPOS_Constants.OPOS_E_CLOSED;
                return _resultCode;
            }

            _eventThreadStart = null;
            _claimManager.Dispose();
            _claimManager = null;

            _oposCO = null;
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
            }

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
                _soundData = new byte[0];
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

        #region OPOS ServiceObject Device Specific Method

        public int StartRecording(string FileName, bool OverWrite, int RecordingTime)
        {
            throw new NotImplementedException();
        }

        public int StopRecording()
        {
            throw new NotImplementedException();
        }

        #endregion OPOS ServiceObject Device Specific Method

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

                    OposEventSoundRecorder soundRecorderEvent = _oposEvents[0];
                    _oposEvents.RemoveAt(0);

                    _eventListMutex.ReleaseMutex();

                    switch (soundRecorderEvent.eventType)
                    {
                        case OposEvent.EVENT_DATA:
                            lock (_dataAccessLock)
                            {
                                _soundData = soundRecorderEvent._soundData;
                                _dataEventEnabled = false;
                            }
                            _oposCO.SOData(soundRecorderEvent.status);
                            break;
                        case OposEvent.EVENT_DIRECTIO:
                            _oposCO.SODirectIO(soundRecorderEvent.eventNumber, ref soundRecorderEvent.pData, ref soundRecorderEvent.pString);
                            // todo:somethingelse
                            break;
                        case OposEvent.EVENT_ERROR:
                            _oposCO.SOError(soundRecorderEvent.resultCode, soundRecorderEvent.resultCodeExtended, soundRecorderEvent.errorLocus, ref soundRecorderEvent.pErrorResponse);
                            if (soundRecorderEvent.pErrorResponse == (int)OPOS_Constants.OPOS_ER_CLEAR)
                            {
                                ClearEventList((OposEvent.EVENT_DATA | OposEvent.EVENT_ERROR));
                            }
                            else
                            {
                                // todo:retry or continueinput
                            }
                            break;
                        case OposEvent.EVENT_STATUSUPDATE:
                            _oposCO.SOStatusUpdate(soundRecorderEvent.data);
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
                        OposEventSoundRecorder e = _oposEvents[i];
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
