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
    [Guid("9278B471-BC3D-4485-8040-B358A633227C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.CAT.OpenPOS.CSSO.CSCATSO.1")]
    public class CSCATSO : IOPOSCATSO, IDisposable
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
        internal bool _trainingMode;
        internal int _transactionType;
        internal int _paymentMedia;
        internal int _logStatus;

        internal bool _capAdditionalSecurityInformation;
        internal bool _capAuthorizeCompletion;
        internal bool _capAuthorizePreSales;
        internal bool _capAuthorizeRefund;
        internal bool _capAuthorizeVoid;
        internal bool _capAuthorizeVoidPreSales;
        internal bool _capCenterResultCode;
        internal bool _capCheckCard;
        internal int _capDailyLog;
        internal bool _capInstallments;
        internal bool _capPaymentDetail;
        internal bool _capTaxOthers;
        internal bool _capTransactionNumber;
        internal bool _capTrainingMode;
        internal bool _capCashDeposit;
        internal bool _capLockTerminal;
        internal bool _capLogStatus;
        internal bool _capUnlockTerminal;

        internal string _accountNumber;
        internal byte[] _additionalSecurityInformation;
        internal string _approvalCode;
        internal string _cardCompanyID;
        internal string _centerResultCode;
        internal byte[] _dailyLog;
        internal string _paymentCondition;
        internal string _paymentDetail;
        internal string _sequenceNumber;
        internal string _slipNumber;
        internal string _transactionNumber;

        internal decimal _balance;
        internal decimal _settledAmount;

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

        public CSCATSO()
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
            _trainingMode = false;
            _transactionType = 0;
            _paymentMedia = 0;
            _logStatus = 0;

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
            _cardCompanyID = string.Empty;
            _centerResultCode = string.Empty;
            _dailyLog = new byte[0];
            _paymentCondition = string.Empty;
            _paymentDetail = string.Empty;
            _sequenceNumber = string.Empty;
            _slipNumber = string.Empty;
            _transactionNumber = string.Empty;

            _balance = 0;
            _settledAmount = 0;

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

        ~CSCATSO()
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
            if (registerType != typeof(CSCATSO))
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
            if (registerType != typeof(CSCATSO))
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

                case (int)OPOSCATInternals.PIDXCat_AsyncMode:
                    value = Convert.ToInt32(_asyncMode);
                    break;

                case (int)OPOSCATInternals.PIDXCat_TrainingMode:
                    value = Convert.ToInt32(_trainingMode);
                    break;

                case (int)OPOSCATInternals.PIDXCat_TransactionType:
                    value = _transactionType;
                    break;

                case (int)OPOSCATInternals.PIDXCat_PaymentMedia:
                    value = _paymentMedia;
                    break;

                case (int)OPOSCATInternals.PIDXCat_LogStatus:
                    value = _logStatus;
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapAdditionalSecurityInformation:
                    value = Convert.ToInt32(_capAdditionalSecurityInformation);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapAuthorizeCompletion:
                    value = Convert.ToInt32(_capAuthorizeCompletion);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapAuthorizePreSales:
                    value = Convert.ToInt32(_capAuthorizePreSales);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapAuthorizeRefund:
                    value = Convert.ToInt32(_capAuthorizeRefund);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapAuthorizeVoid:
                    value = Convert.ToInt32(_capAuthorizeVoid);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapAuthorizeVoidPreSales:
                    value = Convert.ToInt32(_capAuthorizeVoidPreSales);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapCenterResultCode:
                    value = Convert.ToInt32(_capCenterResultCode);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapCheckCard:
                    value = Convert.ToInt32(_capCheckCard);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapDailyLog:
                    value = _capDailyLog;
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapInstallments:
                    value = Convert.ToInt32(_capInstallments);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapPaymentDetail:
                    value = Convert.ToInt32(_capPaymentDetail);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapTaxOthers:
                    value = Convert.ToInt32(_capTaxOthers);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapTransactionNumber:
                    value = Convert.ToInt32(_capTransactionNumber);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapTrainingMode:
                    value = Convert.ToInt32(_capTrainingMode);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapCashDeposit:
                    value = Convert.ToInt32(_capCashDeposit);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapLockTerminal:
                    value = Convert.ToInt32(_capLockTerminal);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapLogStatus:
                    value = Convert.ToInt32(_capLogStatus);
                    break;

                case (int)OPOSCATInternals.PIDXCat_CapUnlockTerminal:
                    value = Convert.ToInt32(_capUnlockTerminal);
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

                case (int)OPOSCATInternals.PIDXCat_AsyncMode:
                    _asyncMode = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSCATInternals.PIDXCat_TrainingMode:
                    _trainingMode = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSCATInternals.PIDXCat_PaymentMedia:
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

                case (int)OPOSCATInternals.PIDXCat_AccountNumber:
                    value = _accountNumber;
                    break;

                case (int)OPOSCATInternals.PIDXCat_AdditionalSecurityInformation:
                    value = SOCommon.ToStringFromByteArray(_additionalSecurityInformation, _binaryConversion);
                    break;

                case (int)OPOSCATInternals.PIDXCat_ApprovalCode:
                    value = _approvalCode;
                    break;

                case (int)OPOSCATInternals.PIDXCat_CardCompanyID:
                    value = _cardCompanyID;
                    break;

                case (int)OPOSCATInternals.PIDXCat_CenterResultCode:
                    value = _centerResultCode;
                    break;

                case (int)OPOSCATInternals.PIDXCat_DailyLog:
                    value = SOCommon.ToStringFromByteArray(_dailyLog, _binaryConversion);
                    break;

                case (int)OPOSCATInternals.PIDXCat_PaymentCondition:
                    value = _paymentCondition;
                    break;

                case (int)OPOSCATInternals.PIDXCat_PaymentDetail:
                    value = _paymentDetail;
                    break;

                case (int)OPOSCATInternals.PIDXCat_SequenceNumber:
                    value = _sequenceNumber;
                    break;

                case (int)OPOSCATInternals.PIDXCat_SlipNumber:
                    value = _slipNumber;
                    break;

                case (int)OPOSCATInternals.PIDXCat_TransactionNumber:
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
                case (int)OPOSCATInternals.PIDXCat_AdditionalSecurityInformation:
                    _additionalSecurityInformation = SOCommon.ToByteArrayFromString(String, _binaryConversion);
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

        #region OPOS ServiceObject Device Specific Property Access

        public decimal GetBalance()
        {
            return _balance;
        }

        public decimal GetSettledAmount()
        {
            return _settledAmount;
        }

        #endregion OPOS ServiceObject Device Specific Property Access

        #region OPOS ServiceObject Device Specific Method

        public int AccessDailyLog(int SequenceNumber, int Type, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int AuthorizeCompletion(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int AuthorizePreSales(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int AuthorizeRefund(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int AuthorizeSales(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int AuthorizeVoid(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int AuthorizeVoidPreSales(int SequenceNumber, decimal Amount, decimal TaxOthers, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int CashDeposit(int SequenceNumber, decimal Amount, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int CheckCard(int SequenceNumber, int Timeout)
        {
            throw new NotImplementedException();
        }

        public int LockTerminal()
        {
            throw new NotImplementedException();
        }

        public int UnlockTerminal()
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