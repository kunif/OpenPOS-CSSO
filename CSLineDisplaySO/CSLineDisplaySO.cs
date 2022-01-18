
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
    [Guid("28461964-0A8C-49E1-912F-8A1DF3886EB5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.LineDisplay.OpenPOS.CSSO.CSLineDisplaySO.1")]
    public class CSLineDisplaySO : IOPOSLineDisplaySO, IDisposable
    {
        internal int _capPowerReporting;
        internal bool _capCompareFirmwareVersion;
        internal bool _capUpdateFirmware;
        internal bool _capStatisticsReporting;
        internal bool _capUpdateStatistics;

        //internal volatile bool _autoDisable;
        internal int _binaryConversion;
        internal string _checkHealthText;
        internal bool _claimed;
        //internal volatile int _dataCount;
        //internal volatile bool _dataEventEnabled;
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

        internal int _characterSet;
        internal int _columns;
        internal int _currentWindow;
        internal int _cursorColumn;
        internal int _cursorRow;
        internal bool _cursorUpdate;
        internal int _deviceBrightness;
        internal int _deviceColumns;
        internal int _deviceDescriptors;
        internal int _deviceRows;
        internal int _deviceWindows;
        internal int _interCharacterWait;
        internal int _marqueeRepeatWait;
        internal int _marqueeType;
        internal int _marqueeUnitWait;
        internal int _rows;
        internal int _marqueeFormat;
        internal int _blinkRate;
        internal int _cursorType;
        internal int _glyphHeight;
        internal int _glyphWidth;
        internal bool _mapCharacterSet;
        internal int _maximumX;
        internal int _maximumY;
        internal int _screenMode;

        internal int _capBlink;
        internal bool _capBrightness;
        internal int _capCharacterSet;
        internal bool _capDescriptors;
        internal bool _capHMarquee;
        internal bool _capICharWait;
        internal bool _capVMarquee;
        internal bool _capBlinkRate;
        internal int _capCursorType;
        internal bool _capCustomGlyph;
        internal int _capReadBack;
        internal int _capReverse;
        internal bool _capBitmap;
        internal bool _capMapCharacterSet;
        internal bool _capScreenMode;

        internal string _characterSetList;
        internal string _customGlyphList;
        internal string _screenModeList;

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

        public CSLineDisplaySO()
        {
            _capPowerReporting = (int)OPOS_Constants.OPOS_PR_STANDARD;
            _capCompareFirmwareVersion = false;
            _capUpdateFirmware = false;
            _capStatisticsReporting = true;
            _capUpdateStatistics = true;

            //_autoDisable = false;
            _binaryConversion = (int)OPOS_Constants.OPOS_BC_NONE;
            _checkHealthText = string.Empty;
            _claimed = false;
            //_dataCount = 0;
            //_dataEventEnabled = false;
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

            _characterSet = 0;
            _columns = 0;
            _currentWindow = 0;
            _cursorColumn = 0;
            _cursorRow = 0;
            _cursorUpdate = false;
            _deviceBrightness = 0;
            _deviceColumns = 0;
            _deviceDescriptors = 0;
            _deviceRows = 0;
            _deviceWindows = 0;
            _interCharacterWait = 0;
            _marqueeRepeatWait = 0;
            _marqueeType = 0;
            _marqueeUnitWait = 0;
            _rows = 0;
            _marqueeFormat = 0;
            _blinkRate = 0;
            _cursorType = 0;
            _glyphHeight = 0;
            _glyphWidth = 0;
            _mapCharacterSet = false;
            _maximumX = 0;
            _maximumY = 0;
            _screenMode = 0;

            _capBlink = 0;
            _capBrightness = false;
            _capCharacterSet = 0;
            _capDescriptors = false;
            _capHMarquee = false;
            _capICharWait = false;
            _capVMarquee = false;
            _capBlinkRate = false;
            _capCursorType = 0;
            _capCustomGlyph = false;
            _capReadBack = 0;
            _capReverse = 0;
            _capBitmap = false;
            _capMapCharacterSet = false;
            _capScreenMode = false;

            _characterSetList = string.Empty;
            _customGlyphList = string.Empty;
            _screenModeList = string.Empty;

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

        ~CSLineDisplaySO()
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
            if (registerType != typeof(CSLineDisplaySO))
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
            if (registerType != typeof(CSLineDisplaySO))
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

                //case (int)OPOS_Internals.PIDX_AutoDisable:
                //    value = Convert.ToInt32(_autoDisable);
                //    break;
                case (int)OPOS_Internals.PIDX_BinaryConversion:
                    value = _binaryConversion;
                    break;

                case (int)OPOS_Internals.PIDX_Claimed:
                    value = Convert.ToInt32(_claimed);
                    break;

                //case (int)OPOS_Internals.PIDX_DataCount:
                //    value = _dataCount;
                //    break;
                //case (int)OPOS_Internals.PIDX_DataEventEnabled:
                //    value = Convert.ToInt32(_dataEventEnabled);
                //    break;
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

                case (int)OPOSLineDisplayInternals.PIDXDisp_CharacterSet:
                    value = _characterSet;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_Columns:
                    value = _columns;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CurrentWindow:
                    value = _currentWindow;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorColumn:
                    value = _cursorColumn;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorRow:
                    value = _cursorRow;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorUpdate:
                    value = Convert.ToInt32(_cursorUpdate);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_DeviceBrightness:
                    value = _deviceBrightness;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_DeviceColumns:
                    value = _deviceColumns;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_DeviceDescriptors:
                    value = _deviceDescriptors;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_DeviceRows:
                    value = _deviceRows;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_DeviceWindows:
                    value = _deviceWindows;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_InterCharacterWait:
                    value = _interCharacterWait;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeRepeatWait:
                    value = _marqueeRepeatWait;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeType:
                    value = _marqueeType;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeUnitWait:
                    value = _marqueeUnitWait;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_Rows:
                    value = _rows;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeFormat:
                    value = _marqueeFormat;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_BlinkRate:
                    value = _blinkRate;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorType:
                    value = _cursorType;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_GlyphHeight:
                    value = _glyphHeight;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_GlyphWidth:
                    value = _glyphWidth;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MapCharacterSet:
                    value = Convert.ToInt32(_mapCharacterSet);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MaximumX:
                    value = _maximumX;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MaximumY:
                    value = _maximumY;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_ScreenMode:
                    value = _screenMode;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapBlink:
                    value = _capBlink;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapBrightness:
                    value = Convert.ToInt32(_capBrightness);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapCharacterSet:
                    value = _capCharacterSet;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapDescriptors:
                    value = Convert.ToInt32(_capDescriptors);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapHMarquee:
                    value = Convert.ToInt32(_capHMarquee);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapICharWait:
                    value = Convert.ToInt32(_capICharWait);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapVMarquee:
                    value = Convert.ToInt32(_capVMarquee);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapBlinkRate:
                    value = Convert.ToInt32(_capBlinkRate);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapCursorType:
                    value = _capCursorType;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapCustomGlyph:
                    value = Convert.ToInt32(_capCustomGlyph);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapReadBack:
                    value = _capReadBack;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapReverse:
                    value = _capReverse;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapBitmap:
                    value = Convert.ToInt32(_capBitmap);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapMapCharacterSet:
                    value = Convert.ToInt32(_capMapCharacterSet);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CapScreenMode:
                    value = Convert.ToInt32(_capScreenMode);
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
                //case (int)OPOS_Internals.PIDX_AutoDisable:
                //    _autoDisable = Convert.ToBoolean(Number);
                //    break;
                case (int)OPOS_Internals.PIDX_BinaryConversion:
                    _binaryConversion = Number;
                    break;

                //case (int)OPOS_Internals.PIDX_DataEventEnabled:
                //    _dataEventEnabled = Convert.ToBoolean(Number);
                //    break;
                case (int)OPOS_Internals.PIDX_DeviceEnabled:
                    _deviceEnabled = Convert.ToBoolean(Number);
                    break;

                case (int)OPOS_Internals.PIDX_FreezeEvents:
                    _freezeEvents = Convert.ToBoolean(Number);
                    break;

                case (int)OPOS_Internals.PIDX_PowerNotify:
                    _powerNotify = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CharacterSet:
                    _characterSet = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CurrentWindow:
                    _currentWindow = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorColumn:
                    _cursorColumn = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorRow:
                    _cursorRow = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorUpdate:
                    _cursorUpdate = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_DeviceBrightness:
                    _deviceBrightness = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_InterCharacterWait:
                    _interCharacterWait = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeRepeatWait:
                    _marqueeRepeatWait = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeType:
                    _marqueeType = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeUnitWait:
                    _marqueeUnitWait = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MarqueeFormat:
                    _marqueeFormat = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_BlinkRate:
                    _blinkRate = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CursorType:
                    _cursorType = Number;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_MapCharacterSet:
                    _mapCharacterSet = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_ScreenMode:
                    _screenMode = Number;
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

                case (int)OPOSLineDisplayInternals.PIDXDisp_CharacterSetList:
                    value = _characterSetList;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_CustomGlyphList:
                    value = _customGlyphList;
                    break;

                case (int)OPOSLineDisplayInternals.PIDXDisp_ScreenModeList:
                    value = _screenModeList;
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
                if (_deviceEnabled && !_freezeEvents && !_coFreezeEvents)
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

        public int ClearDescriptors()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ClearText()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CreateWindow(int ViewportRow, int ViewportColumn, int ViewportHeight, int ViewportWidth, int WindowHeight, int WindowWidth)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DefineGlyph(int GlyphCode, string Glyph)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DestroyWindow()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DisplayBitmap(string FileName, int Width, int AlignmentX, int AlignmentY)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DisplayText(string Data, int Attribute)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DisplayTextAt(int Row, int Column, string Data, int Attribute)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ReadCharacterAtCursor(out int pChar)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            pChar = 0;
            return _resultCode;
        }

        public int RefreshWindow(int Window)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ScrollText(int Direction, int Units)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int SetBitmap(int BitmapNumber, string FileName, int Width, int AlignmentX, int AlignmentY)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int SetDescriptor(int Descriptor, int Attribute)
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
