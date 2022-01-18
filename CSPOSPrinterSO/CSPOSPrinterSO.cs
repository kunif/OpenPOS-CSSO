
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
    [Guid("A3B6F37F-36CF-4E1B-9DEF-F5EF9039AE81")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.POSPrinter.OpenPOS.CSSO.CSPOSPrinterSO.1")]
    public class CSPOSPrinterSO : IOPOSPOSPrinterSO, IDisposable
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

        internal bool _asyncMode;
        internal int _characterSet;
        internal bool _coverOpen;
        internal int _errorStation;
        internal bool _flagWhenIdle;
        internal bool _jrnEmpty;
        internal bool _jrnLetterQuality;
        internal int _jrnLineChars;
        internal int _jrnLineHeight;
        internal int _jrnLineSpacing;
        internal int _jrnLineWidth;
        internal bool _jrnNearEnd;
        internal int _mapMode;
        internal bool _recEmpty;
        internal bool _recLetterQuality;
        internal int _recLineChars;
        internal int _recLineHeight;
        internal int _recLineSpacing;
        internal int _recLinesToPaperCut;
        internal int _recLineWidth;
        internal bool _recNearEnd;
        internal int _recSidewaysMaxChars;
        internal int _recSidewaysMaxLines;
        internal bool _slpEmpty;
        internal bool _slpLetterQuality;
        internal int _slpLineChars;
        internal int _slpLineHeight;
        internal int _slpLinesNearEndToEnd;
        internal int _slpLineSpacing;
        internal int _slpLineWidth;
        internal int _slpMaxLines;
        internal bool _slpNearEnd;
        internal int _slpSidewaysMaxChars;
        internal int _slpSidewaysMaxLines;
        internal int _errorLevel;
        internal int _rotateSpecial;
        internal int _cartridgeNotify;
        internal int _jrnCartridgeState;
        internal int _jrnCurrentCartridge;
        internal int _recCartridgeState;
        internal int _recCurrentCartridge;
        internal int _slpPrintSide;
        internal int _slpCartridgeState;
        internal int _slpCurrentCartridge;
        internal bool _mapCharacterSet;
        internal int _pageModeDescriptor;
        internal int _pageModeHorizontalPosition;
        internal int _pageModePrintDirection;
        internal int _pageModeStation;
        internal int _pageModeVerticalPosition;

        internal bool _capConcurrentJrnRec;
        internal bool _capConcurrentJrnSlp;
        internal bool _capConcurrentRecSlp;
        internal bool _capCoverSensor;
        internal bool _capJrn2Color;
        internal bool _capJrnBold;
        internal bool _capJrnDhigh;
        internal bool _capJrnDwide;
        internal bool _capJrnDwideDhigh;
        internal bool _capJrnEmptySensor;
        internal bool _capJrnItalic;
        internal bool _capJrnNearEndSensor;
        internal bool _capJrnPresent;
        internal bool _capJrnUnderline;
        internal bool _capRec2Color;
        internal bool _capRecBarCode;
        internal bool _capRecBitmap;
        internal bool _capRecBold;
        internal bool _capRecDhigh;
        internal bool _capRecDwide;
        internal bool _capRecDwideDhigh;
        internal bool _capRecEmptySensor;
        internal bool _capRecItalic;
        internal bool _capRecLeft90;
        internal bool _capRecNearEndSensor;
        internal bool _capRecPapercut;
        internal bool _capRecPresent;
        internal bool _capRecRight90;
        internal bool _capRecRotate180;
        internal bool _capRecStamp;
        internal bool _capRecUnderline;
        internal bool _capSlp2Color;
        internal bool _capSlpBarCode;
        internal bool _capSlpBitmap;
        internal bool _capSlpBold;
        internal bool _capSlpDhigh;
        internal bool _capSlpDwide;
        internal bool _capSlpDwideDhigh;
        internal bool _capSlpEmptySensor;
        internal bool _capSlpFullslip;
        internal bool _capSlpItalic;
        internal bool _capSlpLeft90;
        internal bool _capSlpNearEndSensor;
        internal bool _capSlpPresent;
        internal bool _capSlpRight90;
        internal bool _capSlpRotate180;
        internal bool _capSlpUnderline;
        internal int _capCharacterSet;
        internal bool _capTransaction;
        internal int _capJrnCartridgeSensor;
        internal int _capJrnColor;
        internal int _capRecCartridgeSensor;
        internal int _capRecColor;
        internal int _capRecMarkFeed;
        internal bool _capSlpBothSidesPrint;
        internal int _capSlpCartridgeSensor;
        internal int _capSlpColor;
        internal bool _capMapCharacterSet;
        internal bool _capConcurrentPageMode;
        internal bool _capRecPageMode;
        internal bool _capSlpPageMode;
        internal int _capRecRuledLine;
        internal int _capSlpRuledLine;

        internal string _characterSetList;
        internal string _jrnLineCharsList;
        internal string _recLineCharsList;
        internal string _slpLineCharsList;
        internal string _errorString;
        internal string _fontTypefaceList;
        internal string _recBarCodeRotationList;
        internal string _slpBarCodeRotationList;
        internal string _recBitmapRotationList;
        internal string _slpBitmapRotationList;
        internal string _pageModeArea;
        internal string _pageModePrintArea;

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

        public CSPOSPrinterSO()
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

            _asyncMode = false;
            _characterSet = 0;
            _coverOpen = false;
            _errorStation = 0;
            _flagWhenIdle = false;
            _jrnEmpty = false;
            _jrnLetterQuality = false;
            _jrnLineChars = 0;
            _jrnLineHeight = 0;
            _jrnLineSpacing = 0;
            _jrnLineWidth = 0;
            _jrnNearEnd = false;
            _mapMode = 0;
            _recEmpty = false;
            _recLetterQuality = false;
            _recLineChars = 0;
            _recLineHeight = 0;
            _recLineSpacing = 0;
            _recLinesToPaperCut = 0;
            _recLineWidth = 0;
            _recNearEnd = false;
            _recSidewaysMaxChars = 0;
            _recSidewaysMaxLines = 0;
            _slpEmpty = false;
            _slpLetterQuality = false;
            _slpLineChars = 0;
            _slpLineHeight = 0;
            _slpLinesNearEndToEnd = 0;
            _slpLineSpacing = 0;
            _slpLineWidth = 0;
            _slpMaxLines = 0;
            _slpNearEnd = false;
            _slpSidewaysMaxChars = 0;
            _slpSidewaysMaxLines = 0;
            _errorLevel = 0;
            _rotateSpecial = 0;
            _cartridgeNotify = 0;
            _jrnCartridgeState = 0;
            _jrnCurrentCartridge = 0;
            _recCartridgeState = 0;
            _recCurrentCartridge = 0;
            _slpPrintSide = 0;
            _slpCartridgeState = 0;
            _slpCurrentCartridge = 0;
            _mapCharacterSet = false;
            _pageModeDescriptor = 0;
            _pageModeHorizontalPosition = 0;
            _pageModePrintDirection = 0;
            _pageModeStation = 0;
            _pageModeVerticalPosition = 0;

            _capConcurrentJrnRec = false;
            _capConcurrentJrnSlp = false;
            _capConcurrentRecSlp = false;
            _capCoverSensor = false;
            _capJrn2Color = false;
            _capJrnBold = false;
            _capJrnDhigh = false;
            _capJrnDwide = false;
            _capJrnDwideDhigh = false;
            _capJrnEmptySensor = false;
            _capJrnItalic = false;
            _capJrnNearEndSensor = false;
            _capJrnPresent = false;
            _capJrnUnderline = false;
            _capRec2Color = false;
            _capRecBarCode = false;
            _capRecBitmap = false;
            _capRecBold = false;
            _capRecDhigh = false;
            _capRecDwide = false;
            _capRecDwideDhigh = false;
            _capRecEmptySensor = false;
            _capRecItalic = false;
            _capRecLeft90 = false;
            _capRecNearEndSensor = false;
            _capRecPapercut = false;
            _capRecPresent = false;
            _capRecRight90 = false;
            _capRecRotate180 = false;
            _capRecStamp = false;
            _capRecUnderline = false;
            _capSlp2Color = false;
            _capSlpBarCode = false;
            _capSlpBitmap = false;
            _capSlpBold = false;
            _capSlpDhigh = false;
            _capSlpDwide = false;
            _capSlpDwideDhigh = false;
            _capSlpEmptySensor = false;
            _capSlpFullslip = false;
            _capSlpItalic = false;
            _capSlpLeft90 = false;
            _capSlpNearEndSensor = false;
            _capSlpPresent = false;
            _capSlpRight90 = false;
            _capSlpRotate180 = false;
            _capSlpUnderline = false;
            _capCharacterSet = 0;
            _capTransaction = false;
            _capJrnCartridgeSensor = 0;
            _capJrnColor = 0;
            _capRecCartridgeSensor = 0;
            _capRecColor = 0;
            _capRecMarkFeed = 0;
            _capSlpBothSidesPrint = false;
            _capSlpCartridgeSensor = 0;
            _capSlpColor = 0;
            _capMapCharacterSet = false;
            _capConcurrentPageMode = false;
            _capRecPageMode = false;
            _capSlpPageMode = false;
            _capRecRuledLine = 0;
            _capSlpRuledLine = 0;

            _characterSetList = string.Empty;
            _jrnLineCharsList = string.Empty;
            _recLineCharsList = string.Empty;
            _slpLineCharsList = string.Empty;
            _errorString = string.Empty;
            _fontTypefaceList = string.Empty;
            _recBarCodeRotationList = string.Empty;
            _slpBarCodeRotationList = string.Empty;
            _recBitmapRotationList = string.Empty;
            _slpBitmapRotationList = string.Empty;
            _pageModeArea = string.Empty;
            _pageModePrintArea = string.Empty;

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

        ~CSPOSPrinterSO()
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
            if (registerType != typeof(CSPOSPrinterSO))
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
            if (registerType != typeof(CSPOSPrinterSO))
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

                case (int)OPOSPOSPrinterInternals.PIDXPtr_AsyncMode:
                    value = Convert.ToInt32(_asyncMode);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CharacterSet:
                    value = _characterSet;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CoverOpen:
                    value = Convert.ToInt32(_coverOpen);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_ErrorStation:
                    value = _errorStation;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_FlagWhenIdle:
                    value = Convert.ToInt32(_flagWhenIdle);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnEmpty:
                    value = Convert.ToInt32(_jrnEmpty);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLetterQuality:
                    value = Convert.ToInt32(_jrnLetterQuality);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineChars:
                    value = _jrnLineChars;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineHeight:
                    value = _jrnLineHeight;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineSpacing:
                    value = _jrnLineSpacing;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineWidth:
                    value = _jrnLineWidth;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnNearEnd:
                    value = Convert.ToInt32(_jrnNearEnd);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_MapMode:
                    value = _mapMode;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecEmpty:
                    value = Convert.ToInt32(_recEmpty);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLetterQuality:
                    value = Convert.ToInt32(_recLetterQuality);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineChars:
                    value = _recLineChars;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineHeight:
                    value = _recLineHeight;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineSpacing:
                    value = _recLineSpacing;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLinesToPaperCut:
                    value = _recLinesToPaperCut;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineWidth:
                    value = _recLineWidth;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecNearEnd:
                    value = Convert.ToInt32(_recNearEnd);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecSidewaysMaxChars:
                    value = _recSidewaysMaxChars;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecSidewaysMaxLines:
                    value = _recSidewaysMaxLines;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpEmpty:
                    value = Convert.ToInt32(_slpEmpty);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLetterQuality:
                    value = Convert.ToInt32(_slpLetterQuality);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineChars:
                    value = _slpLineChars;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineHeight:
                    value = _slpLineHeight;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLinesNearEndToEnd:
                    value = _slpLinesNearEndToEnd;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineSpacing:
                    value = _slpLineSpacing;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineWidth:
                    value = _slpLineWidth;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpMaxLines:
                    value = _slpMaxLines;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpNearEnd:
                    value = Convert.ToInt32(_slpNearEnd);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpSidewaysMaxChars:
                    value = _slpSidewaysMaxChars;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpSidewaysMaxLines:
                    value = _slpSidewaysMaxLines;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_ErrorLevel:
                    value = _errorLevel;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RotateSpecial:
                    value = _rotateSpecial;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CartridgeNotify:
                    value = _cartridgeNotify;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnCartridgeState:
                    value = _jrnCartridgeState;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnCurrentCartridge:
                    value = _jrnCurrentCartridge;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecCartridgeState:
                    value = _recCartridgeState;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecCurrentCartridge:
                    value = _recCurrentCartridge;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpPrintSide:
                    value = _slpPrintSide;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpCartridgeState:
                    value = _slpCartridgeState;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpCurrentCartridge:
                    value = _slpCurrentCartridge;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_MapCharacterSet:
                    value = Convert.ToInt32(_mapCharacterSet);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeDescriptor:
                    value = _pageModeDescriptor;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeHorizontalPosition:
                    value = _pageModeHorizontalPosition;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModePrintDirection:
                    value = _pageModePrintDirection;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeStation:
                    value = _pageModeStation;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeVerticalPosition:
                    value = _pageModeVerticalPosition;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapConcurrentJrnRec:
                    value = Convert.ToInt32(_capConcurrentJrnRec);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapConcurrentJrnSlp:
                    value = Convert.ToInt32(_capConcurrentJrnSlp);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapConcurrentRecSlp:
                    value = Convert.ToInt32(_capConcurrentRecSlp);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapCoverSensor:
                    value = Convert.ToInt32(_capCoverSensor);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrn2Color:
                    value = _capJrn2Color ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnBold:
                    value = Convert.ToInt32(_capJrnBold);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnDhigh:
                    value = Convert.ToInt32(_capJrnDhigh);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnDwide:
                    value = Convert.ToInt32(_capJrnDwide);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnDwideDhigh:
                    value = Convert.ToInt32(_capJrnDwideDhigh);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnEmptySensor:
                    value = Convert.ToInt32(_capJrnEmptySensor);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnItalic:
                    value = Convert.ToInt32(_capJrnItalic);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnNearEndSensor:
                    value = Convert.ToInt32(_capJrnNearEndSensor);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnPresent:
                    value = Convert.ToInt32(_capJrnPresent);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnUnderline:
                    value = Convert.ToInt32(_capJrnUnderline);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRec2Color:
                    value = _capRec2Color ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecBarCode:
                    value = Convert.ToInt32(_capRecBarCode);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecBitmap:
                    value = Convert.ToInt32(_capRecBitmap);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecBold:
                    value = Convert.ToInt32(_capRecBold);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecDhigh:
                    value = Convert.ToInt32(_capRecDhigh);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecDwide:
                    value = Convert.ToInt32(_capRecDwide);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecDwideDhigh:
                    value = Convert.ToInt32(_capRecDwideDhigh);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecEmptySensor:
                    value = Convert.ToInt32(_capRecEmptySensor);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecItalic:
                    value = Convert.ToInt32(_capRecItalic);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecLeft90:
                    value = _capRecLeft90 ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecNearEndSensor:
                    value = Convert.ToInt32(_capRecNearEndSensor);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecPapercut:
                    value = Convert.ToInt32(_capRecPapercut);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecPresent:
                    value = Convert.ToInt32(_capRecPresent);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecRight90:
                    value = _capRecRight90 ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecRotate180:
                    value = _capRecRotate180 ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecStamp:
                    value = Convert.ToInt32(_capRecStamp);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecUnderline:
                    value = Convert.ToInt32(_capRecUnderline);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlp2Color:
                    value = _capSlp2Color ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpBarCode:
                    value = Convert.ToInt32(_capSlpBarCode);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpBitmap:
                    value = Convert.ToInt32(_capSlpBitmap);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpBold:
                    value = Convert.ToInt32(_capSlpBold);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpDhigh:
                    value = Convert.ToInt32(_capSlpDhigh);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpDwide:
                    value = Convert.ToInt32(_capSlpDwide);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpDwideDhigh:
                    value = Convert.ToInt32(_capSlpDwideDhigh);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpEmptySensor:
                    value = Convert.ToInt32(_capSlpEmptySensor);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpFullslip:
                    value = Convert.ToInt32(_capSlpFullslip);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpItalic:
                    value = Convert.ToInt32(_capSlpItalic);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpLeft90:
                    value = _capSlpLeft90 ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpNearEndSensor:
                    value = Convert.ToInt32(_capSlpNearEndSensor);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpPresent:
                    value = Convert.ToInt32(_capSlpPresent);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpRight90:
                    value = _capSlpRight90 ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpRotate180:
                    value = _capSlpRotate180 ? 1 : 0;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpUnderline:
                    value = Convert.ToInt32(_capSlpUnderline);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapCharacterSet:
                    value = _capCharacterSet;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapTransaction:
                    value = Convert.ToInt32(_capTransaction);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnCartridgeSensor:
                    value = _capJrnCartridgeSensor;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapJrnColor:
                    value = _capJrnColor;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecCartridgeSensor:
                    value = _capRecCartridgeSensor;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecColor:
                    value = _capRecColor;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecMarkFeed:
                    value = _capRecMarkFeed;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpBothSidesPrint:
                    value = Convert.ToInt32(_capSlpBothSidesPrint);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpCartridgeSensor:
                    value = _capSlpCartridgeSensor;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpColor:
                    value = _capSlpColor;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapMapCharacterSet:
                    value = Convert.ToInt32(_capMapCharacterSet);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapConcurrentPageMode:
                    value = Convert.ToInt32(_capConcurrentPageMode);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecPageMode:
                    value = Convert.ToInt32(_capRecPageMode);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpPageMode:
                    value = Convert.ToInt32(_capSlpPageMode);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapRecRuledLine:
                    value = _capRecRuledLine;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CapSlpRuledLine:
                    value = _capSlpRuledLine;
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

                case (int)OPOSPOSPrinterInternals.PIDXPtr_AsyncMode:
                    _asyncMode = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CharacterSet:
                    _characterSet = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_FlagWhenIdle:
                    _flagWhenIdle = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLetterQuality:
                    _jrnLetterQuality = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineChars:
                    _jrnLineChars = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineHeight:
                    _jrnLineHeight = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineSpacing:
                    _jrnLineSpacing = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_MapMode:
                    _mapMode = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLetterQuality:
                    _recLetterQuality = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineChars:
                    _recLineChars = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineHeight:
                    _recLineHeight = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineSpacing:
                    _recLineSpacing = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLetterQuality:
                    _slpLetterQuality = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineChars:
                    _slpLineChars = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineHeight:
                    _slpLineHeight = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineSpacing:
                    _slpLineSpacing = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RotateSpecial:
                    _rotateSpecial = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CartridgeNotify:
                    _cartridgeNotify = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnCurrentCartridge:
                    _jrnCurrentCartridge = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecCurrentCartridge:
                    _recCurrentCartridge = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpCurrentCartridge:
                    _slpCurrentCartridge = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_MapCharacterSet:
                    _mapCharacterSet = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeDescriptor:
                    _pageModeDescriptor = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeHorizontalPosition:
                    _pageModeHorizontalPosition = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModePrintDirection:
                    _pageModePrintDirection = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeStation:
                    _pageModeStation = Number;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeVerticalPosition:
                    _pageModeVerticalPosition = Number;
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

                case (int)OPOSPOSPrinterInternals.PIDXPtr_CharacterSetList:
                    value = _characterSetList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_JrnLineCharsList:
                    value = _jrnLineCharsList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecLineCharsList:
                    value = _recLineCharsList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpLineCharsList:
                    value = _slpLineCharsList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_ErrorString:
                    value = _errorString;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_FontTypefaceList:
                    value = _fontTypefaceList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecBarCodeRotationList:
                    value = _recBarCodeRotationList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpBarCodeRotationList:
                    value = _slpBarCodeRotationList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_RecBitmapRotationList:
                    value = _recBitmapRotationList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_SlpBitmapRotationList:
                    value = _slpBitmapRotationList;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModeArea:
                    value = _pageModeArea;
                    break;

                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModePrintArea:
                    value = _pageModePrintArea;
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
                case (int)OPOSPOSPrinterInternals.PIDXPtr_PageModePrintArea:
                    _pageModePrintArea = String;
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

        public int ChangePrintSide(int Side)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ClearPrintArea()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CutPaper(int Percentage)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DrawRuledLine(int Station, string PositionList, int LineDirection, int LineWidth, int LineStyle, int LineColor)
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

        public int MarkFeed(int Type)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PageModePrint(int Control)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PrintBarCode(int Station, string Data, int Symbology, int Height, int Width, int Alignment, int TextPosition)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PrintBitmap(int Station, string FileName, int Width, int Alignment)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PrintImmediate(int Station, string Data)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PrintMemoryBitmap(int Station, string Data, int Type, int Width, int Alignment)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PrintNormal(int Station, string Data)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int PrintTwoNormal(int Stations, string Data1, string Data2)
        {
            throw new NotImplementedException();
        }

        public int RotatePrint(int Station, int Rotation)
        {
            throw new NotImplementedException();
        }

        public int SetBitmap(int BitmapNumber, int Station, string FileName, int Width, int Alignment)
        {
            throw new NotImplementedException();
        }

        public int SetLogo(int Location, string Data)
        {
            throw new NotImplementedException();
        }

        public int TransactionPrint(int Station, int Control)
        {
            throw new NotImplementedException();
        }

        public int ValidateData(int Station, string Data)
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
