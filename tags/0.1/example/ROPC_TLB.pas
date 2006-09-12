unit ROPC_TLB;

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

// PASTLWTR : 1.2
// File generated on 9/11/2006 2:25:50 PM from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\dopc\trunk\example\opc_test.tlb (1)
// LIBID: {63E0E954-A095-40C2-916A-2580685D44F0}
// LCID: 0
// Helpfile: 
// HelpString: ROPC OPC Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\STDOLE2.TLB)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  ROPCMajorVersion = 2;
  ROPCMinorVersion = 0;

  LIBID_ROPC: TGUID = '{63E0E954-A095-40C2-916A-2580685D44F0}';

  IID_IDA2: TGUID = '{C82218F9-2F0D-4B03-8746-A12D3EB4F4D7}';
  CLASS_DA2: TGUID = '{E8E5FF04-0CD1-46C7-AB54-D09503817522}';
  IID_IOPCGroup: TGUID = '{02EFEFC8-2A35-4C7D-A7E2-28CC24C7BA82}';
  CLASS_OPCGroup: TGUID = '{2E830FD8-F895-4E6C-A82E-BB477A34315C}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IDA2 = interface;
  IDA2Disp = dispinterface;
  IOPCGroup = interface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  DA2 = IDA2;
  OPCGroup = IOPCGroup;


// *********************************************************************//
// Interface: IDA2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C82218F9-2F0D-4B03-8746-A12D3EB4F4D7}
// *********************************************************************//
  IDA2 = interface(IDispatch)
    ['{C82218F9-2F0D-4B03-8746-A12D3EB4F4D7}']
  end;

// *********************************************************************//
// DispIntf:  IDA2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C82218F9-2F0D-4B03-8746-A12D3EB4F4D7}
// *********************************************************************//
  IDA2Disp = dispinterface
    ['{C82218F9-2F0D-4B03-8746-A12D3EB4F4D7}']
  end;

// *********************************************************************//
// Interface: IOPCGroup
// Flags:     (0)
// GUID:      {02EFEFC8-2A35-4C7D-A7E2-28CC24C7BA82}
// *********************************************************************//
  IOPCGroup = interface(IUnknown)
    ['{02EFEFC8-2A35-4C7D-A7E2-28CC24C7BA82}']
  end;

// *********************************************************************//
// The Class CoDA2 provides a Create and CreateRemote method to          
// create instances of the default interface IDA2 exposed by              
// the CoClass DA2. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDA2 = class
    class function Create: IDA2;
    class function CreateRemote(const MachineName: string): IDA2;
  end;

// *********************************************************************//
// The Class CoOPCGroup provides a Create and CreateRemote method to          
// create instances of the default interface IOPCGroup exposed by              
// the CoClass OPCGroup. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOPCGroup = class
    class function Create: IOPCGroup;
    class function CreateRemote(const MachineName: string): IOPCGroup;
  end;

implementation

uses ComObj;

class function CoDA2.Create: IDA2;
begin
  Result := CreateComObject(CLASS_DA2) as IDA2;
end;

class function CoDA2.CreateRemote(const MachineName: string): IDA2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DA2) as IDA2;
end;

class function CoOPCGroup.Create: IOPCGroup;
begin
  Result := CreateComObject(CLASS_OPCGroup) as IOPCGroup;
end;

class function CoOPCGroup.CreateRemote(const MachineName: string): IOPCGroup;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OPCGroup) as IOPCGroup;
end;

end.
