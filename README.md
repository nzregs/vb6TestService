# vb6TestService
a test nt service written in vb6 using ntsvc.ocx


## Installation

### Register the vb runtime
* regsvr32.exe msvbvm60.dll
* (64bit) %windir%\syswow64\regsvr32.exe msvbvm60.dll

### Register the nt service active X control
* regsvr32.exentsvc.ocx
* (64bit) %windir%\syswow64\regsvr32.exe ntsvc.ocx

### Install the nt service
* TestVB6Service.exe -install
