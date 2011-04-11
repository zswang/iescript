unit IEScript20App_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : $Revision:   1.130  $
// File generated on 2011-4-12 0:45:50 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\Projects\GoogleCode\iescript\src\IEScript20App.tlb (1)
// LIBID: {BA1F7536-B591-4D06-830D-6E83A6555324}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINDOWS\system32\stdvcl40.dll)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}

interface

uses ActiveX, Classes, Graphics, StdVCL, Variants, Windows;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  IEScript20AppMajorVersion = 1;
  IEScript20AppMinorVersion = 0;

  LIBID_IEScript20App: TGUID = '{BA1F7536-B591-4D06-830D-6E83A6555324}';

  IID_IIescriptLogger: TGUID = '{19B85BAF-0E2C-4133-B764-E1C9DDEB5F37}';
  CLASS_IescriptLogger: TGUID = '{43ED587C-A537-4BBE-89F8-D9B8C6CA09C8}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IIescriptLogger = interface;
  IIescriptLoggerDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  IescriptLogger = IIescriptLogger;


// *********************************************************************//
// Interface: IIescriptLogger
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {19B85BAF-0E2C-4133-B764-E1C9DDEB5F37}
// *********************************************************************//
  IIescriptLogger = interface(IDispatch)
    ['{19B85BAF-0E2C-4133-B764-E1C9DDEB5F37}']
    procedure log(const AMsg: WideString); safecall;
  end;

// *********************************************************************//
// DispIntf:  IIescriptLoggerDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {19B85BAF-0E2C-4133-B764-E1C9DDEB5F37}
// *********************************************************************//
  IIescriptLoggerDisp = dispinterface
    ['{19B85BAF-0E2C-4133-B764-E1C9DDEB5F37}']
    procedure log(const AMsg: WideString); dispid 1;
  end;

// *********************************************************************//
// The Class CoIescriptLogger provides a Create and CreateRemote method to          
// create instances of the default interface IIescriptLogger exposed by              
// the CoClass IescriptLogger. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoIescriptLogger = class
    class function Create: IIescriptLogger;
    class function CreateRemote(const MachineName: string): IIescriptLogger;
  end;

implementation

uses ComObj;

class function CoIescriptLogger.Create: IIescriptLogger;
begin
  Result := CreateComObject(CLASS_IescriptLogger) as IIescriptLogger;
end;

class function CoIescriptLogger.CreateRemote(const MachineName: string): IIescriptLogger;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_IescriptLogger) as IIescriptLogger;
end;

end.
