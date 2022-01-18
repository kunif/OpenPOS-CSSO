
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
    [Guid("0D565F07-CAD0-4EE0-8DF1-FC59026959FF")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("OPOS.VideoCapture.OpenPOS.CSSO.CSVideoCaptureSO.1")]
    public class CSVideoCaptureSO : IOPOSVideoCaptureSO, IDisposable
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

        internal bool _autoExposure;
        internal bool _autoFocus;
        internal bool _autoGain;
        internal bool _autoWhiteBalance;
        internal int _brightness;
        internal int _contrast;
        internal int _exposure;
        internal int _gain;
        internal bool _horizontalFlip;
        internal int _hue;
        internal int _photoFrameRate;
        internal int _photoMaxFrameRate;
        internal int _remainingRecordingTimeInSec;
        internal int _saturation;
        internal int _storage;
        internal bool _verticalFlip;
        internal int _videoCaptureMode;
        internal int _videoFrameRate;
        internal int _videoMaxFrameRate;

        internal bool _capAutoExposure;
        internal bool _capAutoFocus;
        internal bool _capAutoGain;
        internal bool _capAutoWhiteBalance;
        internal bool _capBrightness;
        internal bool _capContrast;
        internal bool _capExposure;
        internal bool _capGain;
        internal bool _capHorizontalFlip;
        internal bool _capHue;
        internal bool _capPhoto;
        internal bool _capPhotoColorSpace;
        internal bool _capPhotoFrameRate;
        internal bool _capPhotoResolution;
        internal bool _capPhotoType;
        internal bool _capSaturation;
        internal int _capStorage;
        internal bool _capVerticalFlip;
        internal bool _capVideo;
        internal bool _capVideoColorSpace;
        internal bool _capVideoFrameRate;
        internal bool _capVideoResolution;
        internal bool _capVideoType;

        internal string _photoColorSpace;
        internal string _photoColorSpaceList;
        internal string _photoResolution;
        internal string _photoResolutionList;
        internal string _photoType;
        internal string _photoTypeList;
        internal string _videoColorSpace;
        internal string _videoColorSpaceList;
        internal string _videoResolution;
        internal string _videoResolutionList;
        internal string _videoType;
        internal string _videoTypeList;
        internal string _capAssociatedHardTotalsDevice;

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
        internal volatile List<OposEvent> _oposEvents;
        internal volatile bool _killEventThread;

        #region IDisposable Support Constructer / Destructer

        public CSVideoCaptureSO()
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

            _autoExposure = false;
            _autoFocus = false;
            _autoGain = false;
            _autoWhiteBalance = false;
            _brightness = 0;
            _contrast = 0;
            _exposure = 0;
            _gain = 0;
            _horizontalFlip = false;
            _hue = 0;
            _photoFrameRate = 0;
            _photoMaxFrameRate = 0;
            _remainingRecordingTimeInSec = 0;
            _saturation = 0;
            _storage = 0;
            _verticalFlip = false;
            _videoCaptureMode = 0;
            _videoFrameRate = 0;
            _videoMaxFrameRate = 0;

            _capAutoExposure = false;
            _capAutoFocus = false;
            _capAutoGain = false;
            _capAutoWhiteBalance = false;
            _capBrightness = false;
            _capContrast = false;
            _capExposure = false;
            _capGain = false;
            _capHorizontalFlip = false;
            _capHue = false;
            _capPhoto = false;
            _capPhotoColorSpace = false;
            _capPhotoFrameRate = false;
            _capPhotoResolution = false;
            _capPhotoType = false;
            _capSaturation = false;
            _capStorage = 0;
            _capVerticalFlip = false;
            _capVideo = false;
            _capVideoColorSpace = false;
            _capVideoFrameRate = false;
            _capVideoResolution = false;
            _capVideoType = false;

            _photoColorSpace = string.Empty;
            _photoColorSpaceList = string.Empty;
            _photoResolution = string.Empty;
            _photoResolutionList = string.Empty;
            _photoType = string.Empty;
            _photoTypeList = string.Empty;
            _videoColorSpace = string.Empty;
            _videoColorSpaceList = string.Empty;
            _videoResolution = string.Empty;
            _videoResolutionList = string.Empty;
            _videoType = string.Empty;
            _videoTypeList = string.Empty;
            _capAssociatedHardTotalsDevice = string.Empty;

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

        ~CSVideoCaptureSO()
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
            if (registerType != typeof(CSVideoCaptureSO))
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
            if (registerType != typeof(CSVideoCaptureSO))
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

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoExposure:
                    value = Convert.ToInt32(_autoExposure);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoFocus:
                    value = Convert.ToInt32(_autoFocus);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoGain:
                    value = Convert.ToInt32(_autoGain);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoWhiteBalance:
                    value = Convert.ToInt32(_autoWhiteBalance);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Brightness:
                    value = _brightness;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Contrast:
                    value = _contrast;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Exposure:
                    value = _exposure;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Gain:
                    value = _gain;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_HorizontalFlip:
                    value = Convert.ToInt32(_horizontalFlip);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Hue:
                    value = _hue;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoFrameRate:
                    value = _photoFrameRate;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoMaxFrameRate:
                    value = _photoMaxFrameRate;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_RemainingRecordingTimeInSec:
                    value = _remainingRecordingTimeInSec;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Saturation:
                    value = _saturation;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Storage:
                    value = _storage;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VerticalFlip:
                    value = Convert.ToInt32(_verticalFlip);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoCaptureMode:
                    value = _videoCaptureMode;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoFrameRate:
                    value = _videoFrameRate;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoMaxFrameRate:
                    value = _videoMaxFrameRate;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapAutoExposure:
                    value = Convert.ToInt32(_capAutoExposure);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapAutoFocus:
                    value = Convert.ToInt32(_capAutoFocus);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapAutoGain:
                    value = Convert.ToInt32(_capAutoGain);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapAutoWhiteBalance:
                    value = Convert.ToInt32(_capAutoWhiteBalance);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapBrightness:
                    value = Convert.ToInt32(_capBrightness);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapContrast:
                    value = Convert.ToInt32(_capContrast);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapExposure:
                    value = Convert.ToInt32(_capExposure);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapGain:
                    value = Convert.ToInt32(_capGain);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapHorizontalFlip:
                    value = Convert.ToInt32(_capHorizontalFlip);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapHue:
                    value = Convert.ToInt32(_capHue);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapPhotoColorSpace:
                    value = Convert.ToInt32(_capPhotoColorSpace);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapPhotoFrameRate:
                    value = Convert.ToInt32(_capPhotoFrameRate);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapPhotoResolution:
                    value = Convert.ToInt32(_capPhotoResolution);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapPhotoType:
                    value = Convert.ToInt32(_capPhotoType);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapStorage:
                    value = _capStorage;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapSaturation:
                    value = Convert.ToInt32(_capSaturation);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapVerticalFlip:
                    value = Convert.ToInt32(_capVerticalFlip);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapVideo:
                    value = Convert.ToInt32(_capVideo);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapVideoColorSpace:
                    value = Convert.ToInt32(_capVideoColorSpace);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapVideoFrameRate:
                    value = Convert.ToInt32(_capVideoFrameRate);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapVideoResolution:
                    value = Convert.ToInt32(_capVideoResolution);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapVideoType:
                    value = Convert.ToInt32(_capVideoType);
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

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoExposure:
                    _autoExposure = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoFocus:
                    _autoFocus = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoGain:
                    _autoGain = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_AutoWhiteBalance:
                    _autoWhiteBalance = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Brightness:
                    _brightness = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Contrast:
                    _contrast = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Exposure:
                    _exposure = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Gain:
                    _gain = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_HorizontalFlip:
                    _horizontalFlip = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Hue:
                    _hue = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoFrameRate:
                    _photoFrameRate = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Saturation:
                    _saturation = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_Storage:
                    _storage = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VerticalFlip:
                    _verticalFlip = Convert.ToBoolean(Number);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoCaptureMode:
                    _videoCaptureMode = Number;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoFrameRate:
                    _videoFrameRate = Number;
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

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoColorSpace:
                    value = string.Copy(_photoColorSpace);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoColorSpaceList:
                    value = string.Copy(_photoColorSpaceList);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoResolution:
                    value = string.Copy(_photoResolution);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoResolutionList:
                    value = string.Copy(_photoResolutionList);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoType:
                    value = string.Copy(_photoType);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoTypeList:
                    value = string.Copy(_photoTypeList);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoColorSpace:
                    value = string.Copy(_videoColorSpace);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoColorSpaceList:
                    value = string.Copy(_videoColorSpaceList);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoResolution:
                    value = string.Copy(_videoResolution);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoResolutionList:
                    value = string.Copy(_videoResolutionList);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoType:
                    value = string.Copy(_videoType);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoTypeList:
                    value = string.Copy(_videoTypeList);
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_CapAssociatedHardTotalsDevice:
                    value = string.Copy(_capAssociatedHardTotalsDevice);
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
                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoColorSpace:
                    _photoColorSpace = String;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoResolution:
                    _photoResolution = String;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_PhotoType:
                    _photoType = String;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoColorSpace:
                    _videoColorSpace = String;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoResolution:
                    _videoResolution = String;
                    break;

                case (int)OPOSVideoCaptureInternals.PIDXVcap_VideoType:
                    _videoType = String;
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

        #region Device Specific Methods

        public int StartVideo(string FileName, bool OverWrite, int RecordingTime)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int StopVideo()
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        public int TakePhoto(string FileName, bool OverWrite, int Timeout)
        {
            _resultCode = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            _resultCodeExtended = 0;
            return _resultCode;
        }

        #endregion Device Specific Methods

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

                    OposEvent videoCaptureEvent = _oposEvents[0];
                    _oposEvents.RemoveAt(0);

                    _eventListMutex.ReleaseMutex();

                    switch (videoCaptureEvent.eventType)
                    {
                        case OposEvent.EVENT_DIRECTIO:
                            _oposCO.SODirectIO(videoCaptureEvent.eventNumber, ref videoCaptureEvent.pData, ref videoCaptureEvent.pString);
                            // todo:somethingelse
                            break;
                        case OposEvent.EVENT_ERROR:
                            _oposCO.SOError(videoCaptureEvent.resultCode, videoCaptureEvent.resultCodeExtended, videoCaptureEvent.errorLocus, ref videoCaptureEvent.pErrorResponse);
                            if (videoCaptureEvent.pErrorResponse == (int)OPOS_Constants.OPOS_ER_CLEAR)
                            {
                                ClearEventList(OposEvent.EVENT_ERROR);
                            }
                            else
                            {
                                // todo:retry or continueinput
                            }
                            break;
                        case OposEvent.EVENT_OUTPUTCOMPLETE:
                            _oposCO.SOOutputComplete(videoCaptureEvent.outputID);
                            break;
                        case OposEvent.EVENT_STATUSUPDATE:
                            _oposCO.SOStatusUpdate(videoCaptureEvent.data);
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
