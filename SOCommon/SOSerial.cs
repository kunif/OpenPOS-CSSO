
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
    using System;
    using System.Collections.Generic;
    using System.IO.MemoryMappedFiles;
    using System.IO.Ports;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    //using System.Security.AccessControl;
    //using System.Security.Principal;
    using POS.Devices;

    #region SerialPort handling class

    [ComVisible(false)]
    internal class SerialManager : IDisposable
    {
        internal string _portName;
        internal int _baudRate;
        internal int _dataBits;
        internal StopBits _stopBits;
        internal Parity _parityType;
        internal bool _discardNull;
        internal bool _dtrEnable;
        internal bool _rtsEnable;
        internal Handshake _handshake;
        internal string _newLine;
        internal byte _parityReplace;
        internal int _readBufferSize;
        internal int _readTimeout;
        internal int _receivedBytesThreshold;
        internal int _writeBufferSize;
        internal int _writeTimeout;
        //internal int _ctsHolding;
        //internal int _dsrHolding;
        //internal int _breakState;
        //internal bool _isOpen;

        internal SerialPort _serialPort;
        internal int _localBufferSize;
        internal int _localDataSize;
        internal byte[] _localBuffer;
        internal Mutex _bufferControlMutex;

        #region IDisposable Support Constructer / Destructer

        internal SerialManager()
        {
            _portName = string.Empty;
            _baudRate = 9600;
            _dataBits = 8;
            _stopBits = StopBits.One;
            _parityType = Parity.None;
            _discardNull = false;
            _dtrEnable = true;
            _rtsEnable = true;
            _handshake = Handshake.RequestToSend;
            _newLine = string.Empty;
            _parityReplace = 0xFF;
            _readBufferSize = 4096;
            _readTimeout = 300;
            _receivedBytesThreshold = 1;
            _writeBufferSize = 4096;
            _writeTimeout = 300;

            _serialPort = null;
            _localBufferSize = 4096;
            _localDataSize = 0;
            _localBuffer = null;

            _bufferControlMutex = new Mutex(false);
        }

        private bool _disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    _portName = null;
                    _newLine = null;
                    if (_serialPort != null)
                    {
                        _serialPort.Dispose();
                    }
                    if (_localBuffer != null)
                    {
                        _localBuffer = null;
                    }
                    _bufferControlMutex.Dispose();
                }
                _disposedValue = true;

            }
        }

        ~SerialManager()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Support Constructer / Destructer

        #region methods

        internal int GetConfigFromRegistry(string keyPath)
        {
            int result = (int)OPOS_Constants.OPOS_E_NOEXIST;

            using (RegistryKey portKey = Registry.LocalMachine.OpenSubKey(keyPath + @"\SerialPort"))
            {
                _portName = SOCommon.GetStringFromRegistry(portKey, "PortName", _portName);
                _baudRate = SOCommon.GetIntegerFromRegistry(portKey, "BaudRate", _baudRate);
                _dataBits = SOCommon.GetIntegerFromRegistry(portKey, "DataBits", _dataBits);
                _stopBits = SOCommonEnum<StopBits>.ToEnumFromString(SOCommon.GetStringFromRegistry(portKey, "StopBits", string.Empty), _stopBits);
                _parityType = SOCommonEnum<Parity>.ToEnumFromString(SOCommon.GetStringFromRegistry(portKey, "Parity", string.Empty), _parityType);
                _discardNull = SOCommon.GetBoolFromRegistry(portKey, "DiscardNull", _discardNull);
                _dtrEnable = SOCommon.GetBoolFromRegistry(portKey, "DTREnable", _dtrEnable);
                _rtsEnable = SOCommon.GetBoolFromRegistry(portKey, "RTSEnable", _rtsEnable);
                _handshake = SOCommonEnum<Handshake>.ToEnumFromString(SOCommon.GetStringFromRegistry(portKey, "Handshake", string.Empty), _handshake);
                try
                {
                    _newLine = Regex.Unescape(SOCommon.GetStringFromRegistry(portKey, "NewLine", _newLine));
                }
                catch { }
                _parityReplace = (byte)(SOCommon.GetIntegerFromRegistry(portKey, "ParityReplace", _parityReplace));
                _readBufferSize = SOCommon.GetIntegerFromRegistry(portKey, "ReadBufferSize", _readBufferSize);
                _readTimeout = SOCommon.GetIntegerFromRegistry(portKey, "ReadTimeout", _readTimeout);
                _receivedBytesThreshold = SOCommon.GetIntegerFromRegistry(portKey, "ReceivedBytesThreshold", _receivedBytesThreshold);
                _writeBufferSize = SOCommon.GetIntegerFromRegistry(portKey, "WriteBufferSize", _writeBufferSize);
                _writeTimeout = SOCommon.GetIntegerFromRegistry(portKey, "WriteTimeout", _writeTimeout);

                _localBufferSize = SOCommon.GetIntegerFromRegistry(portKey, "LocalBufferSize", _localBufferSize);

                result = (int)OPOS_Constants.OPOS_SUCCESS;
            }
            return result;
        }

        internal int CreatePortObject()
        {
            int result = (int)OPOS_Constants.OPOS_E_FAILURE;
            try
            {
                _serialPort = new SerialPort();
                _serialPort.PortName = _portName;
                _serialPort.BaudRate = _baudRate;
                _serialPort.DataBits = _dataBits;
                _serialPort.StopBits = _stopBits;
                _serialPort.Parity = _parityType;
                _serialPort.DiscardNull = _discardNull;
                _serialPort.DtrEnable = _dtrEnable;
                _serialPort.RtsEnable = _rtsEnable;
                _serialPort.Handshake = _handshake;
                _serialPort.NewLine = _newLine;
                _serialPort.ParityReplace = _parityReplace;
                _serialPort.ReadBufferSize = _readBufferSize;
                _serialPort.ReadTimeout = _readTimeout;
                _serialPort.ReceivedBytesThreshold = _receivedBytesThreshold;
                _serialPort.WriteBufferSize = _writeBufferSize;
                _serialPort.WriteTimeout = _writeTimeout;

                _localBuffer = new byte[_localBufferSize];
                _localDataSize = 0;

                result = (int)OPOS_Constants.OPOS_SUCCESS;
            }
            catch
            {

            }
            return result;
        }

        #endregion methods
    }

    #endregion SerialPort handling class
}
