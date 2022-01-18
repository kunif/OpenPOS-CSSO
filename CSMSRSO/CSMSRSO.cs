
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
    [Guid("6F04EE05-9BEF-42CF-BF3D-7ED4BADA73DC")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.MSR.OpenPOS.CSSO.CSMSRSO.1")]
    public class CSMSRSO : IOPOSMSRSO, IDisposable
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

        internal bool _decodeData;
        internal bool _parseDecodeData;
        internal int _tracksToRead;
        internal int _errorReportingType;
        internal bool _transmitSentinels;
        internal int _encodingMaxLength;
        internal int _tracksToWrite;
        internal int _cardAuthenticationDataLength;
        internal int _dataEncryptionAlgorithm;
        internal bool _deviceAuthenticated;
        internal int _deviceAuthenticationProtocol;
        internal int _track1EncryptedDataLength;
        internal int _track2EncryptedDataLength;
        internal int _track3EncryptedDataLength;
        internal int _track4EncryptedDataLength;

        internal bool _capISO;
        internal bool _capJISOne;
        internal bool _capJISTwo;
        internal bool _capTransmitSentinels;
        internal int _capWritableTracks;
        internal int _capDataEncryption;
        internal int _capDeviceAuthentication;
        internal bool _capTrackDataMasking;

        internal string _accountNumber;
        internal string _expirationDate;
        internal string _firstName;
        internal string _middleInitial;
        internal string _serviceCode;
        internal string _suffix;
        internal string _surname;
        internal string _title;
        internal byte[] _track1Data;
        internal byte[] _track1DiscretionaryData;
        internal byte[] _track2Data;
        internal byte[] _track2DiscretionaryData;
        internal byte[] _track3Data;
        internal byte[] _track4Data;
        internal byte[] _additionalSecurityInformation;
        internal string _capCardAuthentication;
        internal byte[] _cardAuthenticationData;
        internal string _cardPropertyList;
        internal string _cardType;
        internal string _cardTypeList;
        internal byte[] _track1EncryptedData;
        internal byte[] _track2EncryptedData;
        internal byte[] _track3EncryptedData;
        internal byte[] _track4EncryptedData;
        internal string _writeCardType;

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

        public CSMSRSO()
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
            _parseDecodeData = false;
            _tracksToRead = 0;
            _errorReportingType = 0;
            _transmitSentinels = false;
            _encodingMaxLength = 0;
            _tracksToWrite = 0;
            _cardAuthenticationDataLength = 0;
            _dataEncryptionAlgorithm = 0;
            _deviceAuthenticated = false;
            _deviceAuthenticationProtocol = 0;
            _track1EncryptedDataLength = 0;
            _track2EncryptedDataLength = 0;
            _track3EncryptedDataLength = 0;
            _track4EncryptedDataLength = 0;

            _capISO = false;
            _capJISOne = false;
            _capJISTwo = false;
            _capTransmitSentinels = false;
            _capWritableTracks = 0;
            _capDataEncryption = 0;
            _capDeviceAuthentication = 0;
            _capTrackDataMasking = false;

            _accountNumber = string.Empty;
            _expirationDate = string.Empty;
            _firstName = string.Empty;
            _middleInitial = string.Empty;
            _serviceCode = string.Empty;
            _suffix = string.Empty;
            _surname = string.Empty;
            _title = string.Empty;
            _track1Data = new byte[0];
            _track1DiscretionaryData = new byte[0];
            _track2Data = new byte[0];
            _track2DiscretionaryData = new byte[0];
            _track3Data = new byte[0];
            _track4Data = new byte[0];
            _additionalSecurityInformation = new byte[0];
            _capCardAuthentication = string.Empty;
            _cardAuthenticationData = new byte[0];
            _cardPropertyList = string.Empty;
            _cardType = string.Empty;
            _cardTypeList = string.Empty;
            _track1EncryptedData = new byte[0];
            _track2EncryptedData = new byte[0];
            _track3EncryptedData = new byte[0];
            _track4EncryptedData = new byte[0];
            _writeCardType = string.Empty;

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

        ~CSMSRSO()
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
            if (registerType != typeof(CSMSRSO))
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
            if (registerType != typeof(CSMSRSO))
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

                case (int)OPOSMSRInternals.PIDXMsr_DecodeData:
                    value = Convert.ToInt32(_decodeData);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_ParseDecodeData:
                    value = Convert.ToInt32(_parseDecodeData);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_TracksToRead:
                    value = _tracksToRead;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_ErrorReportingType:
                    value = _errorReportingType;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_TransmitSentinels:
                    value = Convert.ToInt32(_transmitSentinels);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_EncodingMaxLength:
                    value = _encodingMaxLength;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_TracksToWrite:
                    value = _tracksToWrite;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CardAuthenticationDataLength:
                    value = _cardAuthenticationDataLength;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_DataEncryptionAlgorithm:
                    value = _dataEncryptionAlgorithm;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_DeviceAuthenticated:
                    value = Convert.ToInt32(_deviceAuthenticated);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_DeviceAuthenticationProtocol:
                    value = _deviceAuthenticationProtocol;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track1EncryptedDataLength:
                    value = _track1EncryptedDataLength;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track2EncryptedDataLength:
                    value = _track2EncryptedDataLength;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track3EncryptedDataLength:
                    value = _track3EncryptedDataLength;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track4EncryptedDataLength:
                    value = _track4EncryptedDataLength;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapISO:
                    value = Convert.ToInt32(_capISO);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapJISOne:
                    value = Convert.ToInt32(_capJISOne);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapJISTwo:
                    value = Convert.ToInt32(_capJISTwo);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapTransmitSentinels:
                    value = Convert.ToInt32(_capTransmitSentinels);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapWritableTracks:
                    value = _capWritableTracks;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapDataEncryption:
                    value = _capDataEncryption;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapDeviceAuthentication:
                    value = _capDeviceAuthentication;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapTrackDataMasking:
                    value = Convert.ToInt32(_capTrackDataMasking);
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

                case (int)OPOSMSRInternals.PIDXMsr_DecodeData:
                    _decodeData = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_ParseDecodeData:
                    _parseDecodeData = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_TracksToRead:
                    _tracksToRead = Number;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_ErrorReportingType:
                    _errorReportingType = Number;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_TransmitSentinels:
                    _transmitSentinels = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_TracksToWrite:
                    _tracksToWrite = Number;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_DataEncryptionAlgorithm:
                    _dataEncryptionAlgorithm = Number;
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

                case (int)OPOSMSRInternals.PIDXMsr_AccountNumber:
                    value = _accountNumber;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_ExpirationDate:
                    value = _expirationDate;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_FirstName:
                    value = _firstName;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_MiddleInitial:
                    value = _middleInitial;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_ServiceCode:
                    value = _serviceCode;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Suffix:
                    value = _suffix;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Surname:
                    value = _surname;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Title:
                    value = _title;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track1Data:
                    value = SOCommon.ToStringFromByteArray(_track1Data, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track1DiscretionaryData:
                    value = SOCommon.ToStringFromByteArray(_track1DiscretionaryData, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track2Data:
                    value = SOCommon.ToStringFromByteArray(_track2Data, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track2DiscretionaryData:
                    value = SOCommon.ToStringFromByteArray(_track2DiscretionaryData, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track3Data:
                    value = SOCommon.ToStringFromByteArray(_track3Data, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track4Data:
                    value = SOCommon.ToStringFromByteArray(_track4Data, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_AdditionalSecurityInformation:
                    value = SOCommon.ToStringFromByteArray(_additionalSecurityInformation, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CapCardAuthentication:
                    value = _capCardAuthentication;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CardAuthenticationData:
                    value = SOCommon.ToStringFromByteArray(_cardAuthenticationData, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CardPropertyList:
                    value = _cardPropertyList;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CardType:
                    value = _cardType;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_CardTypeList:
                    value = _cardTypeList;
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track1EncryptedData:
                    value = SOCommon.ToStringFromByteArray(_track1EncryptedData, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track2EncryptedData:
                    value = SOCommon.ToStringFromByteArray(_track2EncryptedData, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track3EncryptedData:
                    value = SOCommon.ToStringFromByteArray(_track3EncryptedData, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_Track4EncryptedData:
                    value = SOCommon.ToStringFromByteArray(_track4EncryptedData, _binaryConversion);
                    break;

                case (int)OPOSMSRInternals.PIDXMsr_WriteCardType:
                    value = _writeCardType;
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
                case (int)OPOSMSRInternals.PIDXMsr_WriteCardType:
                    _writeCardType = String;
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

        public int AuthenticateDevice(string Response)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int DeauthenticateDevice(string Response)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int RetrieveCardProperty(string Name, out string pValue)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            pValue = string.Empty;
            return _resultCode;
        }

        public int RetrieveDeviceAuthenticationData(ref string pChallenge)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int UpdateKey(string Key, string KeyName)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int WriteTracks(object Data, int Timeout)
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
