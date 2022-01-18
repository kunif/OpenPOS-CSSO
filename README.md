# OpenPOS-CSSO

These are the skeleton codes for creating an OpenPOS service object in C#.
All 45 devices of OpenPOS(UnifiedPOS) ver1.16 are available.

Since it is under development, it has not been fully verified, and its functions and configuration may change significantly.


## Development/Execuion environment

To develop and run this program and the service objects based on it, you need:

- Visual Studio 2022 or Visual Studio Community 2022 version 17.0.5
- .NET Framework 4.8
- The following DLL of OPOS-CCO ver1.16 and its PIA(Primary Interop Assembly) must be registered.
   - Opos_Interfaces.dll
   - Opos_Internals.dll
   - Opos_Constants.dll


## Overview and Known issues

It provides the skeleton code of the common source code required to create, register the registry, and operate the service object of each device class of OpenPOS in C#.

By using this framework to incorporate processing for real devices, we aim to be able to quickly prototype and easily apply it to our products.

- The common processing with the method entry point is already built-in.
- Do not use AnyCPU for the target Platform, and be sure to explicitly specify x86/x64.
- The source code of the barcode scanner device incorporates the one that actually works as an example, but the device-dependent part will be separated and independent.
- At first, all devices incorporate exclusive control type processing. Processing of devices that are not exclusive control such as CashDrawer/Keylock will be incorporated in the future.
- Since the source code of each device is created by mechanically copying and pasting, there may be some parts that do not comply with the specifications.
- If you find a problem, please share the information.


## Customize

- When customizing, be sure to change the following contents.
   - Project name/file name (CSxxxxSO) (xxxx is OpenPOS device class name: same below)
   - namespace (OpenPOS.CSSO)
   - Guid
   - ProgId (OPOS.xxxx.OpenPOS.CSSO.CSxxxxSO.1)
   - AssemblyInfo (OpenPOS.CSSO.CSxxxxSO and other information)


## License

Licensed under the [zlib License](./LICENSE).
