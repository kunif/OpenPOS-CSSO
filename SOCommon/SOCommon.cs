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

    //using System.Security.AccessControl;
    //using System.Security.Principal;
    using OpenPOS.Devices;
    using System;
    using System.Collections.Generic;
    using System.IO.MemoryMappedFiles;
    using System.IO.Ports;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;

    [ComVisible(false)]
    internal static class SOCommon
    {
        #region Device Common Utility function

        internal static int VerifyStateCapability(bool claimed, bool enabled = true, bool capable = true)
        {
            int result = (int)OPOS_Constants.OPOS_SUCCESS;

            if (!claimed)
            {
                result = (int)OPOS_Constants.OPOS_E_NOTCLAIMED;
            }
            else if (!enabled)
            {
                result = (int)OPOS_Constants.OPOS_E_DISABLED;
            }
            else if (!capable)
            {
                result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            }

            return result;
        }

        internal static int VerifyPowerReporting(bool enabled, int capable, int value)
        {
            int result = (int)OPOS_Constants.OPOS_SUCCESS;

            if (enabled)
            {
                result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            }
            else if (capable == (int)OPOS_Constants.OPOS_PR_NONE)
            {
                result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            }
            else if ((value != (int)OPOS_Constants.OPOS_PN_DISABLED) && (value != (int)OPOS_Constants.OPOS_PN_ENABLED))
            {
                result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            }

            return result;
        }

        internal static int[] ToIntegerArray(string integerListString, Char separator)
        {
            List<int> Result = new List<int>();

            if (!string.IsNullOrWhiteSpace(integerListString))
            {
                foreach (string s in integerListString.Split(separator))
                {
                    if (string.IsNullOrWhiteSpace(s))
                    {
                        continue;
                    }

                    int value = 0;

                    if (Int32.TryParse(s.Trim(), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value))
                    {
                        Result.Add(value);
                    }
                }
            }

            return Result.ToArray();
        }

        internal static string[] ToStringArray(string stringListString, Char separator)
        {
            List<string> Result = new List<string>();

            if (!string.IsNullOrWhiteSpace(stringListString))
            {
                foreach (string s in stringListString.Split(separator))
                {
                    if (string.IsNullOrWhiteSpace(s))
                    {
                        continue;
                    }

                    Result.Add(s.Trim());
                }
            }

            return Result.ToArray();
        }

        internal static int CompareBytes(byte[] bytes1, byte[] bytes2, int size)
        {
            int result = 0; // assume all 0
            for (int i = 0; i < size; i++)
            {
                if (bytes1[i] != 0)
                {
                    result = 1; // assume same bytes
                }
            }
            if (result == 0)
            {
                return result;
            }
            for (int i = 0; i < size; i++)
            {
                if (bytes1[i] != bytes2[i])
                {
                    result = -1;    // found different byte
                }
            }
            return result;
        }

        internal static List<string> trueSet = new List<string> { "true", "t", "-1", "1", "yes", "y" };
        internal static List<string> falseSet = new List<string> { "false", "f", "0", "no", "n" };

        internal static bool GetBoolFromRegistry(RegistryKey registryKey, string entryName, bool defaultValue)
        {
            bool result = defaultValue;
            try
            {
                RegistryValueKind kind = registryKey.GetValueKind(entryName);
                if ((kind == RegistryValueKind.DWord) || (kind == RegistryValueKind.QWord))
                {
                    result = Convert.ToBoolean((int)registryKey.GetValue(entryName, defaultValue, RegistryValueOptions.DoNotExpandEnvironmentNames));
                }
                else if (kind == RegistryValueKind.String)
                {
                    string value = ((string)registryKey.GetValue(entryName, string.Empty, RegistryValueOptions.DoNotExpandEnvironmentNames)).Trim().ToLower();
                    try
                    {
                        result = trueSet.Contains(value);
                    }
                    catch
                    {
                        try
                        {
                            result = falseSet.Contains(value);
                        }
                        catch
                        {
                            result = defaultValue;
                        }
                    }
                }
            }
            catch
            {
                result = defaultValue;
            }
            return result;
        }

        internal static byte[] GetByteArrayFromRegistry(RegistryKey registryKey, string entryName, byte[] defaultValue)
        {
            byte[] result = defaultValue;
            try
            {
                RegistryValueKind kind = registryKey.GetValueKind(entryName);
                if (kind == RegistryValueKind.Binary)
                {
                    result = (byte[])registryKey.GetValue(entryName, defaultValue, RegistryValueOptions.DoNotExpandEnvironmentNames);
                }
                else if (kind == RegistryValueKind.String)
                {
                    string value = (string)registryKey.GetValue(entryName, string.Empty, RegistryValueOptions.DoNotExpandEnvironmentNames);
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        result = System.Text.Encoding.Unicode.GetBytes(Regex.Unescape(value));
                    }
                }
            }
            catch
            {
                result = defaultValue;
            }
            return result;
        }

        internal static int GetIntegerFromRegistry(RegistryKey registryKey, string entryName, int defaultValue)
        {
            int result = 0;
            try
            {
                RegistryValueKind kind = registryKey.GetValueKind(entryName);
                if ((kind == RegistryValueKind.DWord) || (kind == RegistryValueKind.QWord))
                {
                    result = (int)registryKey.GetValue(entryName, defaultValue, RegistryValueOptions.DoNotExpandEnvironmentNames);
                }
                else if (kind == RegistryValueKind.String)
                {
                    string value = (string)registryKey.GetValue(entryName, string.Empty, RegistryValueOptions.DoNotExpandEnvironmentNames);
                    if (!int.TryParse(value, (System.Globalization.NumberStyles.Number | System.Globalization.NumberStyles.HexNumber), System.Globalization.CultureInfo.InvariantCulture, out result))
                    {
                        result = defaultValue;
                    }
                }
            }
            catch
            {
                result = defaultValue;
            }
            return result;
        }

        internal static string GetStringFromRegistry(RegistryKey registryKey, string entryName, string defaultValue)
        {
            string result = defaultValue;
            try
            {
                RegistryValueKind kind = registryKey.GetValueKind(entryName);
                if (kind == RegistryValueKind.DWord)
                {
                    result = Convert.ToString((int)registryKey.GetValue(entryName));
                }
                if (kind == RegistryValueKind.QWord)
                {
                    result = Convert.ToString((long)registryKey.GetValue(entryName));
                }
                else if (kind == RegistryValueKind.String)
                {
                    result = (string)registryKey.GetValue(entryName, defaultValue, RegistryValueOptions.DoNotExpandEnvironmentNames);
                }
            }
            catch
            {
                result = defaultValue;
            }
            return result;
        }

        #endregion Device Common Utility function

        #region BinaryConversion related function

        internal static string ToStringFromByteArray(byte[] value, int binaryConversion)
        {
            if (value == null) return null;
            if (value.Length <= 0) return String.Empty;

            StringBuilder Result = new StringBuilder(value.Length);

            switch (binaryConversion)
            {
                default: // Invalid value assumes to be None

                case 0: // None
                    foreach (byte b in value)
                    {
                        Result.Append((char)b);
                    }

                    break;

                case 1: // Nibble
                    Result.Capacity = value.Length * 2;

                    foreach (byte b in value)
                    {
                        Result.Append((char)(((b & 0xF0) >> 4) + '0'));
                        Result.Append((char)((b & 0x0F) + '0'));
                    }

                    break;

                case 2: // Decimal
                    Result.Capacity = value.Length * 3;

                    foreach (byte b in value)
                    {
                        Result.AppendFormat("D3", b);
                    }

                    break;
            }

            return Result.ToString();
        }

        internal static byte[] ToByteArrayFromString(string value, int binaryConversion)
        {
            if (value == null) return null;

            byte[] Result = null;
            int ResultLength = value.Length;
            int Index = 0;
            if (value.Length <= 0) return Result;

            switch (binaryConversion)
            {
                default: // Invalid value assumes to be None

                case 0: // None
                    Result = new byte[ResultLength];

                    foreach (char c in value)
                    {
                        Result[Index++] = (byte)c;
                    }

                    break;

                case 1: // Nibble
                    ResultLength /= 2;
                    Result = new byte[ResultLength];

                    for (int i = 0; i < ResultLength; i++)
                    {
                        Result[i] = (byte)(((value[Index++] & 0x0F) << 4) + (value[Index++] & 0x0F));
                    }

                    break;

                case 2: // Decimal
                    ResultLength /= 3;
                    Result = new byte[ResultLength];

                    for (int i = 0; i < ResultLength; i++)
                    {
                        Result[i] = (byte)(int.Parse(value.Substring((i * 3), 3)));
                    }

                    break;
            }

            return Result;
        }

        #endregion BinaryConversion related function

        #region SerialPort Configuration from Registry

        internal static int GetSerialConfigFromRegistry(ref SerialPort serialPort, string keyPath)
        {
            int result = (int)OPOS_Constants.OPOS_E_NOEXIST;

            using (RegistryKey portKey = Registry.LocalMachine.OpenSubKey(keyPath))
            {
                serialPort.PortName = SOCommon.GetStringFromRegistry(portKey, "PortName", serialPort.PortName);
                serialPort.BaudRate = SOCommon.GetIntegerFromRegistry(portKey, "BaudRate", serialPort.BaudRate);
                serialPort.DataBits = SOCommon.GetIntegerFromRegistry(portKey, "DataBits", serialPort.DataBits);
                serialPort.StopBits = SOCommonEnum<StopBits>.ToEnumFromString(SOCommon.GetStringFromRegistry(portKey, "StopBits", string.Empty), serialPort.StopBits);
                serialPort.Parity = SOCommonEnum<Parity>.ToEnumFromString(SOCommon.GetStringFromRegistry(portKey, "Parity", string.Empty), serialPort.Parity);
                serialPort.DiscardNull = SOCommon.GetBoolFromRegistry(portKey, "DiscardNull", serialPort.DiscardNull);
                serialPort.DtrEnable = SOCommon.GetBoolFromRegistry(portKey, "DTREnable", serialPort.DtrEnable);
                serialPort.RtsEnable = SOCommon.GetBoolFromRegistry(portKey, "RTSEnable", serialPort.RtsEnable);
                serialPort.Handshake = SOCommonEnum<Handshake>.ToEnumFromString(SOCommon.GetStringFromRegistry(portKey, "Handshake", string.Empty), serialPort.Handshake);
                try
                {
                    serialPort.NewLine = Regex.Unescape(SOCommon.GetStringFromRegistry(portKey, "NewLine", serialPort.NewLine));
                }
                catch { }
                serialPort.ParityReplace = (byte)(SOCommon.GetIntegerFromRegistry(portKey, "ParityReplace", serialPort.ParityReplace));
                serialPort.ReadBufferSize = SOCommon.GetIntegerFromRegistry(portKey, "ReadBufferSize", serialPort.ReadBufferSize);
                serialPort.ReadTimeout = SOCommon.GetIntegerFromRegistry(portKey, "ReadTimeout", serialPort.ReadTimeout);
                serialPort.ReceivedBytesThreshold = SOCommon.GetIntegerFromRegistry(portKey, "ReceivedBytesThreshold", serialPort.ReceivedBytesThreshold);
                serialPort.WriteBufferSize = SOCommon.GetIntegerFromRegistry(portKey, "WriteBufferSize", serialPort.WriteBufferSize);
                serialPort.WriteTimeout = SOCommon.GetIntegerFromRegistry(portKey, "WriteTimeout", serialPort.WriteTimeout);

                result = (int)OPOS_Constants.OPOS_SUCCESS;
            }
            return result;
        }

        #endregion SerialPort Configuration from Registry
    }

    #region Enum convert function

    internal static class SOCommonEnum<TEnum>
    {
        internal static TEnum ToEnumFromInteger(int value, TEnum defaultValue)
        {
            try
            {
                return (TEnum)Enum.ToObject(typeof(TEnum), value);
            }
            catch
            {
                return defaultValue;
            }
        }

        internal static TEnum ToEnumFromString(string value, TEnum defaultValue)
        {
            try
            {
                return (TEnum)Enum.Parse(typeof(TEnum), value);
            }
            catch
            {
                return defaultValue;
            }
        }
    }

    #endregion Enum convert function

    #region Claim / Release Management class by MemoryMappedFile

    [ComVisible(false)]
    internal class ClaimManager : IDisposable
    {
        internal bool _claimed;
        internal string _deviceIdName;
        internal byte[] _guidbytes;
        internal int _guidbytessize;
        internal Mutex _claimAccessMutex;
        internal EventWaitHandle _dummyWaitEvent;
        internal EventWaitHandle _cancelWaitEvent;
        internal MemoryMappedFile _mmf;
        internal MemoryMappedViewAccessor _mmva;

        internal const int _areasize = 1024;
        internal static readonly byte[] _clearbytes = Enumerable.Repeat<byte>(0, _areasize).ToArray();

        #region IDisposable Support Constructer / Destructer

        internal ClaimManager(string deviceClass, string deviceName)
        {
            _claimed = false;
            _deviceIdName = "OLEforRetail_OpenPOS_MMF_" + deviceClass + "_" + deviceName;
            string mmFileName = @"Global\ClaimInfo_" + _deviceIdName;
            _guidbytes = new Guid().ToByteArray();
            _guidbytessize = _guidbytes.Length;
            _claimAccessMutex = new Mutex(false, (@"Global\ClaimAccess_" + _deviceIdName));
            _dummyWaitEvent = new EventWaitHandle(false, EventResetMode.ManualReset);
            _cancelWaitEvent = new EventWaitHandle(false, EventResetMode.ManualReset);
            bool result = false;
            try
            {
                result = _claimAccessMutex.WaitOne(0);
            }
            catch (AbandonedMutexException)
            {
                result = true;
            }
            if (result)
            {
                bool created = false;
                try
                {
                    _mmf = MemoryMappedFile.OpenExisting(mmFileName);
                }
                catch
                {
                    _mmf = MemoryMappedFile.CreateNew(mmFileName, _areasize);
                    created = true;
                }
                _mmva = _mmf.CreateViewAccessor();
                if (created)
                {
                    _mmva.WriteArray<byte>(0, _clearbytes, 0, _areasize);
                }
                _claimAccessMutex.ReleaseMutex();
            }
            else
            {
                _mmf = MemoryMappedFile.OpenExisting(mmFileName);
                _mmva = _mmf.CreateViewAccessor();
            }
        }

        private bool _disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    if (_claimed)
                    {
                        byte[] claimedguid = new byte[_guidbytessize];
                        _mmva.ReadArray<byte>(0, claimedguid, 0, _guidbytessize);
                        if (SOCommon.CompareBytes(claimedguid, _guidbytes, _guidbytessize) == 1)
                        {
                            _mmva.WriteArray<byte>(0, _clearbytes, 0, _guidbytessize);
                            _claimed = false;
                        }
                        _mmva.Dispose();
                        _mmf.Dispose();
                    }
                    _cancelWaitEvent.Dispose();
                    _dummyWaitEvent.Dispose();
                    _claimAccessMutex.Dispose();
                }
                _disposedValue = true;
            }
        }

        ~ClaimManager()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Support Constructer / Destructer

        #region 2 pattern of ClaimDevice method

        internal int ClaimDevice(int timeout)
        {
            if (_claimed)
            {
                return (int)OPOS_Constants.OPOS_SUCCESS;
            }
            _dummyWaitEvent.Reset();
            return ClaimDevice(timeout, _dummyWaitEvent, (int)OPOS_Constants.OPOS_E_FAILURE);
        }

        internal int ClaimDevice(int timeout, EventWaitHandle otherReason, int otherValue)
        {
            if (_claimed)
            {
                return (int)OPOS_Constants.OPOS_SUCCESS;
            }
            TimeSpan timeSpan = TimeSpan.FromMilliseconds(timeout);
            DateTime limitTime = DateTime.Now + timeSpan;
            WaitHandle[] waitHandles = { otherReason, _cancelWaitEvent, _claimAccessMutex };
            int result = (int)OPOS_Constants.OPOS_E_TIMEOUT;
            _cancelWaitEvent.Reset();

            do
            {
                if ((timeout != 0) && (timeout != (int)OPOS_Constants.OPOS_FOREVER))
                {
                    timeSpan = limitTime - DateTime.Now;
                    if (timeSpan.Milliseconds < 0)
                    {
                        break;
                    }
                }
                try
                {
                    int index = WaitHandle.WaitAny(waitHandles, timeSpan);
                    if (index == WaitHandle.WaitTimeout)
                    {
                        break;
                    }
                    else if (index == 0)    // otherReason
                    {
                        result = otherValue;
                        break;
                    }
                    else if (index == 1)    // _cancelWait
                    {
                        result = (int)OPOS_Constants.OPOS_E_FAILURE;
                        break;
                    }
                    else if (index == 2)    // _releaseWait
                    {
                        byte[] claimedguid = new byte[_guidbytessize];
                        _mmva.ReadArray<byte>(0, claimedguid, 0, _guidbytessize);
                        int compared = SOCommon.CompareBytes(claimedguid, _guidbytes, _guidbytessize);
                        switch (compared)
                        {
                            case -1:    // different guid : already claimed by other CO&SO in same thread
                                _claimAccessMutex.ReleaseMutex();
                                result = (int)OPOS_Constants.OPOS_E_CLAIMED;
                                break;

                            case 0:     // null : normaly claim
                                _mmva.WriteArray<byte>(0, _guidbytes, 0, _guidbytessize);
                                _claimed = true;
                                result = (int)OPOS_Constants.OPOS_SUCCESS;
                                break;

                            case 1:     // same guid : abnormal? already claimed by this CO&SO
                                _claimed = true;
                                result = (int)OPOS_Constants.OPOS_SUCCESS;
                                _claimAccessMutex.ReleaseMutex();
                                break;
                        }
                        break;
                    }
                }
                catch (AbandonedMutexException)
                {
                    _mmva.WriteArray<byte>(0, _guidbytes, 0, _guidbytessize);
                    _claimed = true;
                    result = (int)OPOS_Constants.OPOS_SUCCESS;
                    break;
                }
            } while (true);
            return result;
        }

        #endregion 2 pattern of ClaimDevice method

        internal void CancelClaim()
        {
            _cancelWaitEvent.Set();
        }

        internal int ReleaseDevice()
        {
            int result = (int)OPOS_Constants.OPOS_E_ILLEGAL;
            if (!_claimed)
            {
                return result;
            }
            byte[] claimedguid = new byte[_guidbytessize];
            _mmva.ReadArray<byte>(0, claimedguid, 0, _guidbytessize);
            if (SOCommon.CompareBytes(claimedguid, _guidbytes, _guidbytessize) == 1)
            {
                _mmva.WriteArray<byte>(0, _clearbytes, 0, _guidbytessize);
                _claimAccessMutex.ReleaseMutex();
            }
            _claimed = false;
            return (int)OPOS_Constants.OPOS_SUCCESS;
        }

        internal int GetClaimStatus()
        {
            int result = -1;    // claimed by other
            if (_claimAccessMutex.WaitOne(0))
            {
                byte[] claimedguid = new byte[_guidbytessize];
                _mmva.ReadArray<byte>(0, claimedguid, 0, _guidbytessize);
                result = SOCommon.CompareBytes(claimedguid, _guidbytes, _guidbytessize);
                _claimAccessMutex.ReleaseMutex();
            }
            return result;
        }
    }

    #endregion Claim / Release Management class by MemoryMappedFile

    #region EventDataStructure

    [ComVisible(false)]
    internal class OposEvent : IDisposable
    {
        internal const int EVENT_DATA = 1;
        internal const int EVENT_DIRECTIO = 2;
        internal const int EVENT_ERROR = 4;
        internal const int EVENT_OUTPUTCOMPLETE = 8;
        internal const int EVENT_STATUSUPDATE = 16;
        internal const int EVENT_TRANSITION = 32;
        internal const int EVENT_ALL = -1;  // 0xFFFFFFFF

        internal const int EVENT_MUTEX_TIMEOUT = 10000;
        internal const int EVENT_CYCLE_TIMEOUT = 10000;

        internal int eventType;             // Indicate the following event type

        //------------------------------
        internal int status;                // DataEvent

        //------------------------------
        internal int eventNumber;           // DirectIOEvent, TransitionEvent

        internal int pData;
        internal string pString;

        //------------------------------
        internal int resultCode;            // ErrorEvent

        internal int resultCodeExtended;
        internal int errorLocus;
        internal int pErrorResponse;

        //------------------------------
        internal int outputID;              // OutputCompleteEvent

        //------------------------------
        internal int data;                  // StatusUpdateEvent

        //------------------------------

        #region IDisposable Support Constructer / Destructer

        public OposEvent()
        {
            eventType = EVENT_ALL;
            status = 0;
            eventNumber = 0;
            pData = 0;
            pString = string.Empty;
            resultCode = 0;
            resultCodeExtended = 0;
            errorLocus = 0;
            pErrorResponse = 0;
            outputID = 0;
            data = 0;
        }

        private bool _disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    pString = null;
                }
                _disposedValue = true;
            }
        }

        ~OposEvent()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Support Constructer / Destructer
    }

    #endregion EventDataStructure
}