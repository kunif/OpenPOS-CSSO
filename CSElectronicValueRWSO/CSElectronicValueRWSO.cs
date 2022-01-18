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
    [Guid("F62C63BD-E9D7-4349-B717-9BEFCF925A08")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.ElectronicValueRW.OpenPOS.CSSO.CSElectronicValueRWSO.1")]
    public class CSElectronicValueRWSO : IOPOSElectronicValueRWSO_1_15, IDisposable
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

        internal bool _asyncMode;
        internal bool _detectionControl;
        internal int _detectionStatus;
        internal int _logStatus;
        internal int _sequenceNumber;
        internal int _pINEntry;
        internal int _trainingModeState;
        internal int _serviceType;

        internal int _paymentCondition;
        internal int _paymentMedia;
        internal int _transactionType;

        internal bool _capActivateService;
        internal bool _capAddValue;
        internal bool _capCancelValue;
        internal int _capCardSensor;
        internal int _capDetectionControl;
        internal bool _capElectronicMoney;
        internal bool _capEnumerateCardServices;
        internal bool _capIndirectTransactionLog;
        internal bool _capLockTerminal;
        internal bool _capLogStatus;
        internal bool _capMediumID;
        internal bool _capPoint;
        internal bool _capSubtractValue;
        internal bool _capTransaction;
        internal bool _capTransactionLog;
        internal bool _capUnlockTerminal;
        internal bool _capUpdateKey;
        internal bool _capVoucher;
        internal bool _capWriteValue;
        internal bool _capPINDevice;
        internal bool _capTrainingMode;
        internal bool _capMembershipCertificate;

        internal bool _capAdditionalSecurityInformation;
        internal bool _capAuthorizeCompletion;
        internal bool _capAuthorizePreSales;
        internal bool _capAuthorizeRefund;
        internal bool _capAuthorizeVoid;
        internal bool _capAuthorizeVoidPreSales;
        internal bool _capCashDeposit;
        internal bool _capCenterResultCode;
        internal bool _capCheckCard;
        internal int _capDailyLog;
        internal bool _capInstallments;
        internal bool _capPaymentDetail;
        internal bool _capTaxOthers;
        internal bool _capTransactionNumber;

        internal string _accountNumber;
        internal byte[] _additionalSecurityInformation;
        internal string _approvalCode;
        internal string _cardServiceList;
        internal string _currentService;
        internal string _expirationDate;
        internal string _lastUsedDate;
        internal string _mediumID;
        internal string _readerWriterServiceList;
        internal string _transactionLog;
        internal string _voucherID;
        internal string _voucherIDList;

        internal string _cardCompanyID;
        internal string _centerResultCode;
        internal byte[] _dailyLog;
        internal string _paymentDetail;
        internal string _slipNumber;
        internal string _transactionNumber;

        internal decimal _amount;
        internal decimal _balance;
        internal decimal _balanceOfPoint;
        internal decimal _point;
        internal decimal _settledAmount;
        internal decimal _settledPoint;

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

        public CSElectronicValueRWSO()
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

            _asyncMode = false;
            _detectionControl = false;
            _detectionStatus = 0;
            _logStatus = 0;
            _sequenceNumber = 0;
            _pINEntry = 0;
            _trainingModeState = 0;
            _serviceType = 0;

            _paymentCondition = 0;
            _sequenceNumber = 0;
            _transactionType = 0;

            _capActivateService = false;
            _capAddValue = false;
            _capCancelValue = false;
            _capCardSensor = 0;
            _capDetectionControl = 0;
            _capElectronicMoney = false;
            _capEnumerateCardServices = false;
            _capIndirectTransactionLog = false;
            _capLockTerminal = false;
            _capLogStatus = false;
            _capMediumID = false;
            _capPoint = false;
            _capSubtractValue = false;
            _capTransaction = false;
            _capTransactionLog = false;
            _capUnlockTerminal = false;
            _capUpdateKey = false;
            _capVoucher = false;
            _capWriteValue = false;
            _capPINDevice = false;
            _capTrainingMode = false;
            _capMembershipCertificate = false;

            _capAdditionalSecurityInformation = false;
            _capAuthorizeCompletion = false;
            _capAuthorizePreSales = false;
            _capAuthorizeRefund = false;
            _capAuthorizeVoid = false;
            _capAuthorizeVoidPreSales = false;
            _capCenterResultCode = false;
            _capCheckCard = false;
            _capDailyLog = 0;
            _capInstallments = false;
            _capPaymentDetail = false;
            _capTaxOthers = false;
            _capTransactionNumber = false;
            _capTrainingMode = false;
            _capCashDeposit = false;
            _capLockTerminal = false;
            _capLogStatus = false;
            _capUnlockTerminal = false;

            _accountNumber = string.Empty;
            _additionalSecurityInformation = new byte[0];
            _approvalCode = string.Empty;
            _cardServiceList = string.Empty;
            _currentService = string.Empty;
            _expirationDate = string.Empty;
            _lastUsedDate = string.Empty;
            _mediumID = string.Empty;
            _readerWriterServiceList = string.Empty;
            _transactionLog = string.Empty;
            _voucherID = string.Empty;
            _voucherIDList = string.Empty;

            _cardCompanyID = string.Empty;
            _centerResultCode = string.Empty;
            _dailyLog = new byte[0];
            _paymentDetail = string.Empty;
            _slipNumber = string.Empty;
            _transactionNumber = string.Empty;

            _amount = 0;
            _balance = 0;
            _balanceOfPoint = 0;
            _point = 0;
            _settledAmount = 0;
            _settledPoint = 0;

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

        ~CSElectronicValueRWSO()
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
            if (registerType != typeof(CSElectronicValueRWSO))
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
            if (registerType != typeof(CSElectronicValueRWSO))
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

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_AsyncMode:
                    value = Convert.ToInt32(_asyncMode);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_DetectionControl:
                    value = Convert.ToInt32(_detectionControl);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_DetectionStatus:
                    value = _detectionStatus;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_LogStatus:
                    value = _logStatus;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_SequenceNumber:
                    value = _sequenceNumber;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_PINEntry:
                    value = _pINEntry;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_TrainingModeState:
                    value = _trainingModeState;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_ServiceType:
                    value = _serviceType;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_PaymentCondition:
                    value = _paymentCondition;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_PaymentMedia:
                    value = _paymentMedia;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_TransactionType:
                    value = _transactionType;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapActivateService:
                    value = Convert.ToInt32(_capActivateService);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapAddValue:
                    value = Convert.ToInt32(_capAddValue);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapCancelValue:
                    value = Convert.ToInt32(_capCancelValue);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapCardSensor:
                    value = _capCardSensor;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapDetectionControl:
                    value = _capDetectionControl;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapElectronicMoney:
                    value = Convert.ToInt32(_capElectronicMoney);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapEnumerateCardServices:
                    value = Convert.ToInt32(_capEnumerateCardServices);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapIndirectTransactionLog:
                    value = Convert.ToInt32(_capIndirectTransactionLog);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapLockTerminal:
                    value = Convert.ToInt32(_capLockTerminal);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapLogStatus:
                    value = Convert.ToInt32(_capLogStatus);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapMediumID:
                    value = Convert.ToInt32(_capMediumID);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapPoint:
                    value = Convert.ToInt32(_capPoint);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapSubtractValue:
                    value = Convert.ToInt32(_capSubtractValue);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapTransaction:
                    value = Convert.ToInt32(_capTransaction);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapTransactionLog:
                    value = Convert.ToInt32(_capTransactionLog);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapUnlockTerminal:
                    value = Convert.ToInt32(_capUnlockTerminal);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapUpdateKey:
                    value = Convert.ToInt32(_capUpdateKey);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapVoucher:
                    value = Convert.ToInt32(_capVoucher);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapWriteValue:
                    value = Convert.ToInt32(_capWriteValue);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapPINDevice:
                    value = Convert.ToInt32(_capPINDevice);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapTrainingMode:
                    value = Convert.ToInt32(_capTrainingMode);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapMembershipCertificate:
                    value = Convert.ToInt32(_capMembershipCertificate);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapAdditionalSecurityInformation:
                    value = Convert.ToInt32(_capAdditionalSecurityInformation);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapAuthorizeCompletion:
                    value = Convert.ToInt32(_capAuthorizeCompletion);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapAuthorizePreSales:
                    value = Convert.ToInt32(_capAuthorizePreSales);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapAuthorizeRefund:
                    value = Convert.ToInt32(_capAuthorizeRefund);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapAuthorizeVoid:
                    value = Convert.ToInt32(_capAuthorizeVoid);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapAuthorizeVoidPreSales:
                    value = Convert.ToInt32(_capAuthorizeVoidPreSales);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapCashDeposit:
                    value = Convert.ToInt32(_capCashDeposit);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapCenterResultCode:
                    value = Convert.ToInt32(_capCenterResultCode);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapCheckCard:
                    value = Convert.ToInt32(_capCheckCard);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapDailyLog:
                    value = _capDailyLog;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapInstallments:
                    value = Convert.ToInt32(_capInstallments);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapPaymentDetail:
                    value = Convert.ToInt32(_capPaymentDetail);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapTaxOthers:
                    value = Convert.ToInt32(_capTaxOthers);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CapTransactionNumber:
                    value = Convert.ToInt32(_capTransactionNumber);
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

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_AsyncMode:
                    _asyncMode = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_DetectionControl:
                    _detectionControl = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_PINEntry:
                    _pINEntry = Number;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_TrainingModeState:
                    _trainingModeState = Number;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_PaymentMedia:
                    _paymentMedia = Number;
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

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_AccountNumber:
                    value = _accountNumber;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_AdditionalSecurityInformation:
                    value = SOCommon.ToStringFromByteArray(_additionalSecurityInformation, _binaryConversion);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_ApprovalCode:
                    value = _approvalCode;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CardServiceList:
                    value = _cardServiceList;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CurrentService:
                    value = _currentService;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_ExpirationDate:
                    value = _expirationDate;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_LastUsedDate:
                    value = _lastUsedDate;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_MediumID:
                    value = _mediumID;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_ReaderWriterServiceList:
                    value = _readerWriterServiceList;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_TransactionLog:
                    value = _transactionLog;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_VoucherID:
                    value = _voucherID;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_VoucherIDList:
                    value = _voucherIDList;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CardCompanyID:
                    value = _cardCompanyID;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CenterResultCode:
                    value = _centerResultCode;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_DailyLog:
                    value = SOCommon.ToStringFromByteArray(_dailyLog, _binaryConversion);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_PaymentDetail:
                    value = _paymentDetail;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_SlipNumber:
                    value = _slipNumber;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_TransactionNumber:
                    value = _transactionNumber;
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
                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_AdditionalSecurityInformation:
                    _additionalSecurityInformation = SOCommon.ToByteArrayFromString(String, _binaryConversion);
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_ApprovalCode:
                    _approvalCode = String;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_CurrentService:
                    _currentService = String;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_MediumID:
                    _mediumID = String;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_VoucherID:
                    _voucherID = String;
                    break;

                case (int)OPOSElectronicValueRWInternals.PIDXEvrw_VoucherIDList:
                    _voucherIDList = String;
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

        #region OPOS ServiceObject Device Specific Properties

        public decimal GetAmount()
        {
            return _amount;
        }

        public void SetAmount(decimal Amount)
        {
            _amount = Amount;
            _resultCode = (int)OPOS_Constants.OPOS_SUCCESS;
            _resultCodeExtended = 0;
            return;
        }

        public decimal GetBalance()
        {
            return _balance;
        }

        public decimal GetBalanceOfPoint()
        {
            return _balanceOfPoint;
        }

        public decimal GetPoint()
        {
            return _point;
        }

        public void SetPoint(decimal Point)
        {
            _point = Point;
            _resultCode = (int)OPOS_Constants.OPOS_SUCCESS;
            _resultCodeExtended = 0;
            return;
        }

        public decimal GetSettledAmount()
        {
            return _settledAmount;
        }

        public decimal GetSettledPoint()
        {
            return _settledPoint;
        }

        #endregion OPOS ServiceObject Device Specific Properties

        #region OPOS ServiceObject Device Specific Method

        public int AccessData(int DataType, ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AccessLog(int SequenceNumber, int Type, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ActivateEVService(ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ActivateService(ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AddValue(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int BeginDetection(int Type, int Timeout)
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

        public int CancelValue(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CaptureCard()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CheckServiceRegistrationToMedium(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ClearParameterInformation()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CloseDailyEVService(ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DeactivateEVService(ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int EndDetection()
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

        public int EnumerateCardServices()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int LockTerminal()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int OpenDailyEVService(ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int QueryLastSuccessfulTransactionResult()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int ReadValue(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int RegisterServiceToMedium(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int RetrieveResultInformation(string Name, ref string pValue)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int SetParameterInformation(string Name, string Value)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int SubtractValue(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int TransactionAccess(int Control)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int UnlockTerminal()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int UnregisterServiceToMedium(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int UpdateData(int DataType, ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int UpdateKey(ref int pData, ref string pObj)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int WriteValue(int SequenceNumber, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AccessDailyLog(int SequenceNumber, int Type, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AuthorizeCompletion(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AuthorizePreSales(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AuthorizeRefund(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AuthorizeSales(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AuthorizeVoid(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int AuthorizeVoidPreSales(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CashDeposit(int SequenceNumber, decimal Amount, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int CheckCard(int SequenceNumber, int Timeout)
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