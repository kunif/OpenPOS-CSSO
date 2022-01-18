
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

    [ComVisible(true)]
    [Guid("B4D123AB-24A9-4C22-94A4-7ABF92C89CA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.PointCardRW.OpenPOS.CSSO.CSPointCardRWSO.1")]
    public class CSPointCardRWSO : IOPOSPointCardRWSO, IDisposable
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
        internal int _outputID;
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

        internal int _cardState;
        internal int _characterSet;
        internal int _lineChars;
        internal int _lineHeight;
        internal int _lineSpacing;
        internal int _lineWidth;
        internal int _mapMode;
        internal int _maxLine;
        internal int _printHeight;
        internal int _readState1;
        internal int _readState2;
        internal int _recvLength1;
        internal int _recvLength2;
        internal int _sidewaysMaxChars;
        internal int _sidewaysMaxLines;
        internal int _tracksToRead;
        internal int _tracksToWrite;
        internal int _writeState1;
        internal int _writeState2;
        internal bool _mapCharacterSet;

        internal bool _capBold;
        internal int _capCardEntranceSensor;
        internal int _capCharacterSet;
        internal bool _capCleanCard;
        internal bool _capClearPrint;
        internal bool _capDhigh;
        internal bool _capDwide;
        internal bool _capDwideDhigh;
        internal bool _capItalic;
        internal bool _capLeft90;
        internal bool _capPrint;
        internal bool _capPrintMode;
        internal bool _capRight90;
        internal bool _capRotate180;
        internal int _capTracksToRead;
        internal int _capTracksToWrite;
        internal bool _capMapCharacterSet;

        internal string _characterSetList;
        internal string _fontTypeFaceList;
        internal string _lineCharsList;
        internal string _track1Data;
        internal string _track2Data;
        internal string _track3Data;
        internal string _track4Data;
        internal string _track5Data;
        internal string _track6Data;
        internal string _write1Data;
        internal string _write2Data;
        internal string _write3Data;
        internal string _write4Data;
        internal string _write5Data;
        internal string _write6Data;

        internal volatile bool _opened;
        internal dynamic _oposCO;
        internal string _deviceNameKey;

        internal ClaimManager _claimManager;

        internal volatile bool _releaseClosing;

        internal volatile Object _dataAccessLock;
        internal volatile Mutex _eventControlMutex;
        internal volatile Mutex _eventListMutex;
        internal volatile EventWaitHandle _eventReactivateEvent;

        internal ThreadStart _eventThreadStart;
        internal Thread _eventThread;
        internal volatile List<OposEvent> _oposEvents;
        internal volatile bool _killEventThread;

        #region IDisposable Support Constructer / Destructer

        public CSPointCardRWSO()
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
            _outputID = 0;
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

            _cardState = 0;
            _characterSet = 0;
            _lineChars = 0;
            _lineHeight = 0;
            _lineSpacing = 0;
            _lineWidth = 0;
            _mapMode = 0;
            _maxLine = 0;
            _printHeight = 0;
            _readState1 = 0;
            _readState2 = 0;
            _recvLength1 = 0;
            _recvLength2 = 0;
            _sidewaysMaxChars = 0;
            _sidewaysMaxLines = 0;
            _tracksToRead = 0;
            _tracksToWrite = 0;
            _writeState1 = 0;
            _writeState2 = 0;
            _mapCharacterSet = false;

            _capBold = false;
            _capCardEntranceSensor = 0;
            _capCharacterSet = 0;
            _capCleanCard = false;
            _capClearPrint = false;
            _capDhigh = false;
            _capDwide = false;
            _capDwideDhigh = false;
            _capItalic = false;
            _capLeft90 = false;
            _capPrint = false;
            _capPrintMode = false;
            _capRight90 = false;
            _capRotate180 = false;
            _capTracksToRead = 0;
            _capTracksToWrite = 0;
            _capMapCharacterSet = false;

            _characterSetList = string.Empty;
            _fontTypeFaceList = string.Empty;
            _lineCharsList = string.Empty;
            _track1Data = string.Empty;
            _track2Data = string.Empty;
            _track3Data = string.Empty;
            _track4Data = string.Empty;
            _track5Data = string.Empty;
            _track6Data = string.Empty;
            _write1Data = string.Empty;
            _write2Data = string.Empty;
            _write3Data = string.Empty;
            _write4Data = string.Empty;
            _write5Data = string.Empty;
            _write6Data = string.Empty;

            _opened = false;
            _deviceNameKey = string.Empty;

            _oposCO = null;

            _claimManager = null;

            _releaseClosing = false;

            _dataAccessLock = new object();
            _eventControlMutex = new Mutex(false);
            _eventListMutex = new Mutex(false);
            _eventReactivateEvent = new EventWaitHandle(false, EventResetMode.ManualReset);
            _oposEvents = new List<OposEvent>();
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

        ~CSPointCardRWSO()
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
            if (registerType != typeof(CSPointCardRWSO))
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
            if (registerType != typeof(CSPointCardRWSO))
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

                case (int)OPOS_Internals.PIDX_OutputID:
                    value = _outputID;
                    break;

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

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CardState:
                    value = _cardState;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CharacterSet:
                    value = _characterSet;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_LineChars:
                    value = _lineChars;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_LineHeight:
                    value = _lineHeight;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_LineSpacing:
                    value = _lineSpacing;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_LineWidth:
                    value = _lineWidth;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_MapMode:
                    value = _mapMode;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_MaxLine:
                    value = _maxLine;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_PrintHeight:
                    value = _printHeight;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_ReadState1:
                    value = _readState1;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_ReadState2:
                    value = _readState2;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_RecvLength1:
                    value = _recvLength1;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_RecvLength2:
                    value = _recvLength2;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_SidewaysMaxChars:
                    value = _sidewaysMaxChars;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_SidewaysMaxLines:
                    value = _sidewaysMaxLines;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_TracksToRead:
                    value = _tracksToRead;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_TracksToWrite:
                    value = _tracksToWrite;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_WriteState1:
                    value = _writeState1;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_WriteState2:
                    value = _writeState2;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_MapCharacterSet:
                    value = Convert.ToInt32(_mapCharacterSet);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapBold:
                    value = Convert.ToInt32(_capBold);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapCardEntranceSensor:
                    value = _capCardEntranceSensor;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapCharacterSet:
                    value = _capCharacterSet;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapCleanCard:
                    value = Convert.ToInt32(_capCleanCard);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapClearPrint:
                    value = Convert.ToInt32(_capClearPrint);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapDhigh:
                    value = Convert.ToInt32(_capDhigh);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapDwide:
                    value = Convert.ToInt32(_capDwide);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapDwideDhigh:
                    value = Convert.ToInt32(_capDwideDhigh);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapItalic:
                    value = Convert.ToInt32(_capItalic);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapLeft90:
                    value = _capLeft90 ? 1 : 0;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapPrint:
                    value = Convert.ToInt32(_capPrint);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapPrintMode:
                    value = Convert.ToInt32(_capPrintMode);
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapRight90:
                    value = _capRight90 ? 1 : 0;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapRotate180:
                    value = _capRotate180 ? 1 : 0;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapTracksToRead:
                    value = _capTracksToRead;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapTracksToWrite:
                    value = _capTracksToWrite;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CapMapCharacterSet:
                    value = Convert.ToInt32(_capMapCharacterSet);
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
                    _dataEventEnabled = Convert.ToBoolean(Number);
                    break;

                case (int)OPOS_Internals.PIDX_DeviceEnabled:
                    _deviceEnabled = Convert.ToBoolean(Number);
                    break;

                case (int)OPOS_Internals.PIDX_FreezeEvents:
                    _freezeEvents = Convert.ToBoolean(Number);
                    break;

                case (int)OPOS_Internals.PIDX_PowerNotify:
                    _powerNotify = Number;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CharacterSet:
                    _characterSet = Number;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_TracksToRead:
                    _tracksToRead = Number;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_TracksToWrite:
                    _tracksToWrite = Number;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_MapCharacterSet:
                    _mapCharacterSet = Convert.ToBoolean(Number);
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

                case (int)OPOSPointCardRWInternals.PIDXPcrw_CharacterSetList:
                    value = _characterSetList;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_FontTypeFaceList:
                    value = _fontTypeFaceList;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_LineCharsList:
                    value = _lineCharsList;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Track1Data:
                    value = _track1Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Track2Data:
                    value = _track2Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Track3Data:
                    value = _track3Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Track4Data:
                    value = _track4Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Track5Data:
                    value = _track5Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Track6Data:
                    value = _track6Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write1Data:
                    value = _write1Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write2Data:
                    value = _write2Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write3Data:
                    value = _write3Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write4Data:
                    value = _write4Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write5Data:
                    value = _write5Data;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write6Data:
                    value = _write6Data;
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
                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write1Data:
                    _write1Data = String;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write2Data:
                    _write2Data = String;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write3Data:
                    _write3Data = String;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write4Data:
                    _write4Data = String;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write5Data:
                    _write5Data = String;
                    break;

                case (int)OPOSPointCardRWInternals.PIDXPcrw_Write6Data:
                    _write6Data = String;
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

            using (RegistryKey propKey = Registry.LocalMachine.OpenSubKey(key + @"\PropertySpecific"))
            {
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

        public int ClearInput()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ClearInputProperties()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ClearOutput()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
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

        public int BeginInsertion(int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int BeginRemoval(int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CleanCard()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ClearPrintWrite(int Kind, int Hposition, int Vposition, int Width, int Height)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int EndInsertion()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int EndRemoval()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PrintWrite(int Kind, int Hposition, int Vposition, string Data)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int RotatePrint(int Rotation)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ValidateData(string Data)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
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

                    OposEvent oposEvent = _oposEvents[0];
                    _oposEvents.RemoveAt(0);

                    _eventListMutex.ReleaseMutex();

                    switch (oposEvent.eventType)
                    {
                        case OposEvent.EVENT_DATA:

                            lock (_dataAccessLock)
                            {
                                _dataEventEnabled = false;
                            }
                            _oposCO.SOData(oposEvent.status);
                            break;
                        case OposEvent.EVENT_DIRECTIO:
                            _oposCO.SODirectIO(oposEvent.eventNumber, ref oposEvent.pData, ref oposEvent.pString);
                            // todo:somethingelse
                            break;
                        case OposEvent.EVENT_ERROR:
                            _oposCO.SOError(oposEvent.resultCode, oposEvent.resultCodeExtended, oposEvent.errorLocus, ref oposEvent.pErrorResponse);
                            if (oposEvent.pErrorResponse == (int)OPOS_Constants.OPOS_ER_CLEAR)
                            {
                                ClearEventList((OposEvent.EVENT_DATA | OposEvent.EVENT_ERROR));
                            }
                            else
                            {
                                // todo:retry or continueinput
                            }
                            break;
                        case OposEvent.EVENT_STATUSUPDATE:
                            _oposCO.SOStatusUpdate(oposEvent.data);
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
                        OposEvent e = _oposEvents[i];
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
