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
    [Guid("AE35C023-E507-41CB-AF32-833B26E34E88")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.FiscalPrinter.OpenPOS.CSSO.CSFiscalPrinterSO.1")]
    public class CSFiscalPrinterSO : IOPOSFiscalPrinterSO, IDisposable
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

        internal int _amountDecimalPlaces;
        internal bool _asyncMode;
        internal bool _checkTotal;
        internal int _countryCode;
        internal bool _coverOpen;
        internal bool _dayOpened;
        internal int _descriptionLength;
        internal bool _duplicateReceipt;
        internal int _errorLevel;
        internal int _errorOutID;
        internal int _errorState;
        internal int _errorStation;
        internal bool _flagWhenIdle;
        internal bool _jrnEmpty;
        internal bool _jrnNearEnd;
        internal int _messageLength;
        internal int _numHeaderLines;
        internal int _numTrailerLines;
        internal int _numVatRates;
        internal int _printerState;
        internal int _quantityDecimalPlaces;
        internal int _quantityLength;
        internal bool _recEmpty;
        internal bool _recNearEnd;
        internal int _remainingFiscalMemory;
        internal bool _slpEmpty;
        internal bool _slpNearEnd;
        internal int _slipSelection;
        internal bool _trainingModeActive;
        internal int _actualCurrency;
        internal int _contractorId;
        internal int _dateType;
        internal int _fiscalReceiptStation;
        internal int _fiscalReceiptType;
        internal int _messageType;
        internal int _totalizerType;

        internal bool _capAdditionalLines;
        internal bool _capAmountAdjustment;
        internal bool _capAmountNotPaid;
        internal bool _capCheckTotal;
        internal bool _capCoverSensor;
        internal bool _capDoubleWidth;
        internal bool _capDuplicateReceipt;
        internal bool _capFixedOutput;
        internal bool _capHasVatTable;
        internal bool _capIndependentHeader;
        internal bool _capItemList;
        internal bool _capJrnEmptySensor;
        internal bool _capJrnNearEndSensor;
        internal bool _capJrnPresent;
        internal bool _capNonFiscalMode;
        internal bool _capOrderAdjustmentFirst;
        internal bool _capPercentAdjustment;
        internal bool _capPositiveAdjustment;
        internal bool _capPowerLossReport;
        internal bool _capPredefinedPaymentLines;
        internal bool _capReceiptNotPaid;
        internal bool _capRecEmptySensor;
        internal bool _capRecNearEndSensor;
        internal bool _capRecPresent;
        internal bool _capRemainingFiscalMemory;
        internal bool _capReservedWord;
        internal bool _capSetHeader;
        internal bool _capSetPOSID;
        internal bool _capSetStoreFiscalID;
        internal bool _capSetTrailer;
        internal bool _capSetVatTable;
        internal bool _capSlpEmptySensor;
        internal bool _capSlpFiscalDocument;
        internal bool _capSlpFullSlip;
        internal bool _capSlpNearEndSensor;
        internal bool _capSlpPresent;
        internal bool _capSlpValidation;
        internal bool _capSubAmountAdjustment;
        internal bool _capSubPercentAdjustment;
        internal bool _capSubtotal;
        internal bool _capTrainingMode;
        internal bool _capValidateJournal;
        internal bool _capXReport;
        internal bool _capAdditionalHeader;
        internal bool _capAdditionalTrailer;
        internal bool _capChangeDue;
        internal bool _capEmptyReceiptIsVoidable;
        internal bool _capFiscalReceiptStation;
        internal bool _capFiscalReceiptType;
        internal bool _capMultiContractor;
        internal bool _capOnlyVoidLastItem;
        internal bool _capPackageAdjustment;
        internal bool _capPostPreLine;
        internal bool _capSetCurrency;
        internal bool _capTotalizerType;
        internal bool _capPositiveSubtotalAdjustment;

        internal string _errorString;
        internal string _predefinedPaymentLines;
        internal string _reservedWord;
        internal string _additionalHeader;
        internal string _additionalTrailer;
        internal string _changeDue;
        internal string _postLine;
        internal string _preLine;

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

        public CSFiscalPrinterSO()
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

            _amountDecimalPlaces = 0;
            _asyncMode = false;
            _checkTotal = false;
            _countryCode = 0;
            _coverOpen = false;
            _dayOpened = false;
            _descriptionLength = 0;
            _duplicateReceipt = false;
            _errorLevel = 0;
            _errorOutID = 0;
            _errorState = 0;
            _errorStation = 0;
            _flagWhenIdle = false;
            _jrnEmpty = false;
            _jrnNearEnd = false;
            _messageLength = 0;
            _numHeaderLines = 0;
            _numTrailerLines = 0;
            _numVatRates = 0;
            _printerState = 0;
            _quantityDecimalPlaces = 0;
            _quantityLength = 0;
            _recEmpty = false;
            _recNearEnd = false;
            _remainingFiscalMemory = 0;
            _slpEmpty = false;
            _slpNearEnd = false;
            _slipSelection = 0;
            _trainingModeActive = false;
            _actualCurrency = 0;
            _contractorId = 0;
            _dateType = 0;
            _fiscalReceiptStation = 0;
            _fiscalReceiptType = 0;
            _messageType = 0;
            _totalizerType = 0;

            _capAdditionalLines = false;
            _capAmountAdjustment = false;
            _capAmountNotPaid = false;
            _capCheckTotal = false;
            _capCoverSensor = false;
            _capDoubleWidth = false;
            _capDuplicateReceipt = false;
            _capFixedOutput = false;
            _capHasVatTable = false;
            _capIndependentHeader = false;
            _capItemList = false;
            _capJrnEmptySensor = false;
            _capJrnNearEndSensor = false;
            _capJrnPresent = false;
            _capNonFiscalMode = false;
            _capOrderAdjustmentFirst = false;
            _capPercentAdjustment = false;
            _capPositiveAdjustment = false;
            _capPowerLossReport = false;
            _capPredefinedPaymentLines = false;
            _capReceiptNotPaid = false;
            _capRecEmptySensor = false;
            _capRecNearEndSensor = false;
            _capRecPresent = false;
            _capRemainingFiscalMemory = false;
            _capReservedWord = false;
            _capSetHeader = false;
            _capSetPOSID = false;
            _capSetStoreFiscalID = false;
            _capSetTrailer = false;
            _capSetVatTable = false;
            _capSlpEmptySensor = false;
            _capSlpFiscalDocument = false;
            _capSlpFullSlip = false;
            _capSlpNearEndSensor = false;
            _capSlpPresent = false;
            _capSlpValidation = false;
            _capSubAmountAdjustment = false;
            _capSubPercentAdjustment = false;
            _capSubtotal = false;
            _capTrainingMode = false;
            _capValidateJournal = false;
            _capXReport = false;
            _capAdditionalHeader = false;
            _capAdditionalTrailer = false;
            _capChangeDue = false;
            _capEmptyReceiptIsVoidable = false;
            _capFiscalReceiptStation = false;
            _capFiscalReceiptType = false;
            _capMultiContractor = false;
            _capOnlyVoidLastItem = false;
            _capPackageAdjustment = false;
            _capPostPreLine = false;
            _capSetCurrency = false;
            _capTotalizerType = false;
            _capPositiveSubtotalAdjustment = false;

            _errorString = string.Empty;
            _predefinedPaymentLines = string.Empty;
            _reservedWord = string.Empty;
            _additionalHeader = string.Empty;
            _additionalTrailer = string.Empty;
            _changeDue = string.Empty;
            _postLine = string.Empty;
            _preLine = string.Empty;

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

        ~CSFiscalPrinterSO()
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
            if (registerType != typeof(CSFiscalPrinterSO))
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
            if (registerType != typeof(CSFiscalPrinterSO))
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

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_AmountDecimalPlaces:
                    value = _amountDecimalPlaces;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_AsyncMode:
                    value = Convert.ToInt32(_asyncMode);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CheckTotal:
                    value = Convert.ToInt32(_checkTotal);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CountryCode:
                    value = _countryCode;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CoverOpen:
                    value = Convert.ToInt32(_coverOpen);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_DayOpened:
                    value = Convert.ToInt32(_dayOpened);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_DescriptionLength:
                    value = _descriptionLength;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_DuplicateReceipt:
                    value = Convert.ToInt32(_duplicateReceipt);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ErrorLevel:
                    value = _errorLevel;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ErrorOutID:
                    value = _errorOutID;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ErrorState:
                    value = _errorState;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ErrorStation:
                    value = _errorStation;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_FlagWhenIdle:
                    value = Convert.ToInt32(_flagWhenIdle);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_JrnEmpty:
                    value = Convert.ToInt32(_jrnEmpty);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_JrnNearEnd:
                    value = Convert.ToInt32(_jrnNearEnd);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_MessageLength:
                    value = _messageLength;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_NumHeaderLines:
                    value = _numHeaderLines;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_NumTrailerLines:
                    value = _numTrailerLines;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_NumVatRates:
                    value = _numVatRates;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_PrinterState:
                    value = _printerState;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_QuantityDecimalPlaces:
                    value = _quantityDecimalPlaces;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_QuantityLength:
                    value = _quantityLength;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_RecEmpty:
                    value = Convert.ToInt32(_recEmpty);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_RecNearEnd:
                    value = Convert.ToInt32(_recNearEnd);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_RemainingFiscalMemory:
                    value = _remainingFiscalMemory;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_SlpEmpty:
                    value = Convert.ToInt32(_slpEmpty);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_SlpNearEnd:
                    value = Convert.ToInt32(_slpNearEnd);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_SlipSelection:
                    value = _slipSelection;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_TrainingModeActive:
                    value = Convert.ToInt32(_trainingModeActive);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ActualCurrency:
                    value = _actualCurrency;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ContractorId:
                    value = _contractorId;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_DateType:
                    value = _dateType;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_FiscalReceiptStation:
                    value = _fiscalReceiptStation;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_FiscalReceiptType:
                    value = _fiscalReceiptType;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_MessageType:
                    value = _messageType;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_TotalizerType:
                    value = _totalizerType;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapAdditionalLines:
                    value = Convert.ToInt32(_capAdditionalLines);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapAmountAdjustment:
                    value = Convert.ToInt32(_capAmountAdjustment);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapAmountNotPaid:
                    value = Convert.ToInt32(_capAmountNotPaid);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapCheckTotal:
                    value = Convert.ToInt32(_capCheckTotal);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapCoverSensor:
                    value = Convert.ToInt32(_capCoverSensor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapDoubleWidth:
                    value = Convert.ToInt32(_capDoubleWidth);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapDuplicateReceipt:
                    value = Convert.ToInt32(_capDuplicateReceipt);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapFixedOutput:
                    value = Convert.ToInt32(_capFixedOutput);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapHasVatTable:
                    value = Convert.ToInt32(_capHasVatTable);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapIndependentHeader:
                    value = Convert.ToInt32(_capIndependentHeader);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapItemList:
                    value = Convert.ToInt32(_capItemList);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapJrnEmptySensor:
                    value = Convert.ToInt32(_capJrnEmptySensor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapJrnNearEndSensor:
                    value = Convert.ToInt32(_capJrnNearEndSensor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapJrnPresent:
                    value = Convert.ToInt32(_capJrnPresent);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapNonFiscalMode:
                    value = Convert.ToInt32(_capNonFiscalMode);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapOrderAdjustmentFirst:
                    value = Convert.ToInt32(_capOrderAdjustmentFirst);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapPercentAdjustment:
                    value = Convert.ToInt32(_capPercentAdjustment);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapPositiveAdjustment:
                    value = Convert.ToInt32(_capPositiveAdjustment);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapPowerLossReport:
                    value = Convert.ToInt32(_capPowerLossReport);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapPredefinedPaymentLines:
                    value = Convert.ToInt32(_capPredefinedPaymentLines);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapReceiptNotPaid:
                    value = Convert.ToInt32(_capReceiptNotPaid);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapRecEmptySensor:
                    value = Convert.ToInt32(_capRecEmptySensor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapRecNearEndSensor:
                    value = Convert.ToInt32(_capRecNearEndSensor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapRecPresent:
                    value = Convert.ToInt32(_capRecPresent);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapRemainingFiscalMemory:
                    value = Convert.ToInt32(_capRemainingFiscalMemory);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapReservedWord:
                    value = Convert.ToInt32(_capReservedWord);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSetHeader:
                    value = Convert.ToInt32(_capSetHeader);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSetPOSID:
                    value = Convert.ToInt32(_capSetPOSID);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSetStoreFiscalID:
                    value = Convert.ToInt32(_capSetStoreFiscalID);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSetTrailer:
                    value = Convert.ToInt32(_capSetTrailer);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSetVatTable:
                    value = Convert.ToInt32(_capSetVatTable);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSlpEmptySensor:
                    value = Convert.ToInt32(_capSlpEmptySensor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSlpFiscalDocument:
                    value = Convert.ToInt32(_capSlpFiscalDocument);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSlpFullSlip:
                    value = Convert.ToInt32(_capSlpFullSlip);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSlpNearEndSensor:
                    value = Convert.ToInt32(_capSlpNearEndSensor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSlpPresent:
                    value = Convert.ToInt32(_capSlpPresent);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSlpValidation:
                    value = Convert.ToInt32(_capSlpValidation);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSubAmountAdjustment:
                    value = Convert.ToInt32(_capSubAmountAdjustment);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSubPercentAdjustment:
                    value = Convert.ToInt32(_capSubPercentAdjustment);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSubtotal:
                    value = Convert.ToInt32(_capSubtotal);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapTrainingMode:
                    value = Convert.ToInt32(_capTrainingMode);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapValidateJournal:
                    value = Convert.ToInt32(_capValidateJournal);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapXReport:
                    value = Convert.ToInt32(_capXReport);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapAdditionalHeader:
                    value = Convert.ToInt32(_capAdditionalHeader);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapAdditionalTrailer:
                    value = Convert.ToInt32(_capAdditionalTrailer);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapChangeDue:
                    value = Convert.ToInt32(_capChangeDue);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapEmptyReceiptIsVoidable:
                    value = Convert.ToInt32(_capEmptyReceiptIsVoidable);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapFiscalReceiptStation:
                    value = Convert.ToInt32(_capFiscalReceiptStation);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapFiscalReceiptType:
                    value = Convert.ToInt32(_capFiscalReceiptType);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapMultiContractor:
                    value = Convert.ToInt32(_capMultiContractor);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapOnlyVoidLastItem:
                    value = Convert.ToInt32(_capOnlyVoidLastItem);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapPackageAdjustment:
                    value = Convert.ToInt32(_capPackageAdjustment);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapPostPreLine:
                    value = Convert.ToInt32(_capPostPreLine);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapSetCurrency:
                    value = Convert.ToInt32(_capSetCurrency);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapTotalizerType:
                    value = Convert.ToInt32(_capTotalizerType);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CapPositiveSubtotalAdjustment:
                    value = Convert.ToInt32(_capPositiveSubtotalAdjustment);
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

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_AsyncMode:
                    _asyncMode = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_CheckTotal:
                    _checkTotal = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_DuplicateReceipt:
                    _duplicateReceipt = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_FlagWhenIdle:
                    _flagWhenIdle = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_SlipSelection:
                    _slipSelection = Number;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ContractorId:
                    _contractorId = Number;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_DateType:
                    _dateType = Number;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_FiscalReceiptStation:
                    _fiscalReceiptStation = Number;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_FiscalReceiptType:
                    _fiscalReceiptType = Number;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_MessageType:
                    _messageType = Number;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_TotalizerType:
                    _totalizerType = Number;
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

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ErrorString:
                    value = _errorString;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_PredefinedPaymentLines:
                    value = _predefinedPaymentLines;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ReservedWord:
                    value = _reservedWord;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_AdditionalHeader:
                    value = _additionalHeader;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_AdditionalTrailer:
                    value = _additionalTrailer;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ChangeDue:
                    value = _changeDue;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_PostLine:
                    value = _postLine;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_PreLine:
                    value = _preLine;
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
                case (int)OPOSFiscalPrinterInternals.PIDXFptr_AdditionalHeader:
                    _additionalHeader = String;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_AdditionalTrailer:
                    _additionalTrailer = String;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_ChangeDue:
                    _changeDue = String;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_PostLine:
                    _postLine = String;
                    break;

                case (int)OPOSFiscalPrinterInternals.PIDXFptr_PreLine:
                    _preLine = String;
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

        public int BeginFiscalDocument(int DocumentAmount)
        {
            throw new NotImplementedException();
        }

        public int BeginFiscalReceipt(bool PrintHeader)
        {
            throw new NotImplementedException();
        }

        public int BeginFixedOutput(int Station, int DocumentType)
        {
            throw new NotImplementedException();
        }

        public int BeginInsertion(int Timeout)
        {
            throw new NotImplementedException();
        }

        public int BeginItemList(int VatID)
        {
            throw new NotImplementedException();
        }

        public int BeginNonFiscal()
        {
            throw new NotImplementedException();
        }

        public int BeginRemoval(int Timeout)
        {
            throw new NotImplementedException();
        }

        public int BeginTraining()
        {
            throw new NotImplementedException();
        }

        public int ClearError()
        {
            throw new NotImplementedException();
        }

        public int EndFiscalDocument()
        {
            throw new NotImplementedException();
        }

        public int EndFiscalReceipt(bool PrintHeader)
        {
            throw new NotImplementedException();
        }

        public int EndFixedOutput()
        {
            throw new NotImplementedException();
        }

        public int EndInsertion()
        {
            throw new NotImplementedException();
        }

        public int EndItemList()
        {
            throw new NotImplementedException();
        }

        public int EndNonFiscal()
        {
            throw new NotImplementedException();
        }

        public int EndRemoval()
        {
            throw new NotImplementedException();
        }

        public int EndTraining()
        {
            throw new NotImplementedException();
        }

        public int GetData(int DataItem, out int OptArgs, out string Data)
        {
            throw new NotImplementedException();
        }

        public int GetDate(out string Date)
        {
            throw new NotImplementedException();
        }

        public int GetTotalizer(int VatID, int OptArgs, out string Data)
        {
            throw new NotImplementedException();
        }

        public int GetVatEntry(int VatID, int OptArgs, out int VatRate)
        {
            throw new NotImplementedException();
        }

        public int PrintDuplicateReceipt()
        {
            throw new NotImplementedException();
        }

        public int PrintFiscalDocumentLine(string DocumentLine)
        {
            throw new NotImplementedException();
        }

        public int PrintFixedOutput(int DocumentType, int LineNumber, string Data)
        {
            throw new NotImplementedException();
        }

        public int PrintNormal(int Station, string Data)
        {
            throw new NotImplementedException();
        }

        public int PrintPeriodicTotalsReport(string Date1, string Date2)
        {
            throw new NotImplementedException();
        }

        public int PrintPowerLossReport()
        {
            throw new NotImplementedException();
        }

        public int PrintRecCash(decimal Amount)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItem(string Description, decimal Price, int Quantity, int VatInfo, decimal UnitPrice, string UnitName)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItemAdjustment(int AdjustmentType, string Description, decimal Amount, int VatInfo)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItemAdjustmentVoid(int AdjustmentType, string Description, decimal Amount, int VatInfo)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItemFuel(string Description, decimal Price, int Quantity, int VatInfo, decimal UnitPrice, string UnitName, decimal SpecialTax, string SpecialTaxName)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItemFuelVoid(string Description, decimal Price, int VatInfo, decimal SpecialTax)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItemRefund(string Description, decimal Amount, int Quantity, int VatInfo, decimal UnitAmount, string UnitName)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItemRefundVoid(string Description, decimal Amount, int Quantity, int VatInfo, decimal UnitAmount, string UnitName)
        {
            throw new NotImplementedException();
        }

        public int PrintRecItemVoid(string Description, decimal Price, int Quantity, int VatInfo, decimal UnitPrice, string UnitName)
        {
            throw new NotImplementedException();
        }

        public int PrintRecMessage(string Message)
        {
            throw new NotImplementedException();
        }

        public int PrintRecNotPaid(string Description, decimal Amount)
        {
            throw new NotImplementedException();
        }

        public int PrintRecPackageAdjustment(int AdjustmentType, string Description, string VatAdjustment)
        {
            throw new NotImplementedException();
        }

        public int PrintRecPackageAdjustVoid(int AdjustmentType, string VatAdjustment)
        {
            throw new NotImplementedException();
        }

        public int PrintRecRefund(string Description, decimal Amount, int VatInfo)
        {
            throw new NotImplementedException();
        }

        public int PrintRecRefundVoid(string Description, decimal Amount, int VatInfo)
        {
            throw new NotImplementedException();
        }

        public int PrintRecSubtotal(decimal Amount)
        {
            throw new NotImplementedException();
        }

        public int PrintRecSubtotalAdjustment(int AdjustmentType, string Description, decimal Amount)
        {
            throw new NotImplementedException();
        }

        public int PrintRecSubtotalAdjustVoid(int AdjustmentType, decimal Amount)
        {
            throw new NotImplementedException();
        }

        public int PrintRecTaxID(string TaxID)
        {
            throw new NotImplementedException();
        }

        public int PrintRecTotal(decimal Total, decimal Payment, string Description)
        {
            throw new NotImplementedException();
        }

        public int PrintRecVoid(string Description)
        {
            throw new NotImplementedException();
        }

        public int PrintRecVoidItem(string Description, decimal Amount, int Quantity, int AdjustmentType, decimal Adjustment, int VatInfo)
        {
            throw new NotImplementedException();
        }

        public int PrintReport(int ReportType, string StartNum, string EndNum)
        {
            throw new NotImplementedException();
        }

        public int PrintXReport()
        {
            throw new NotImplementedException();
        }

        public int PrintZReport()
        {
            throw new NotImplementedException();
        }

        public int ResetPrinter()
        {
            throw new NotImplementedException();
        }

        public int SetCurrency(int NewCurrency)
        {
            throw new NotImplementedException();
        }

        public int SetDate(string Date)
        {
            throw new NotImplementedException();
        }

        public int SetHeaderLine(int LineNumber, string Text, bool DoubleWidth)
        {
            throw new NotImplementedException();
        }

        public int SetPOSID(string POSID, string CashierID)
        {
            throw new NotImplementedException();
        }

        public int SetStoreFiscalID(string ID)
        {
            throw new NotImplementedException();
        }

        public int SetTrailerLine(int LineNumber, string Text, bool DoubleWidth)
        {
            throw new NotImplementedException();
        }

        public int SetVatTable()
        {
            throw new NotImplementedException();
        }

        public int SetVatValue(int VatID, string VatValue)
        {
            throw new NotImplementedException();
        }

        public int VerifyItem(string ItemName, int VatID)
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