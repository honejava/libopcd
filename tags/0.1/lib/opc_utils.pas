unit opc_utils;

{$IFDEF VER150}
{$WARN UNSAFE_CODE OFF}
{$ENDIF}

interface

uses windows, variants, messages, sysutils, classes, graphics, stdctrls, forms,
  dialogs, controls, shellapi, activex, opc_da;

type
 itemIDStrings = record
  trunk,branch,leaf:string[255];
 end;

type
  itemProps = record
    PropID: longword;
    tagname: string[64];
    dataType:integer;
  end;

const
  IID_IUnknown: TIID = '{00000000-0000-0000-C000-000000000046}';

  io2Read = 1;
  io2Write = 2;
  io2Refresh = 3;

function DataTimeToOPCTime(time: TDateTime): TFileTime;
function ConvertVariant(const value: variant; requestedType: TVarType): variant;

type
  TTaskMemAllocKind = (tmHResult, tmItemState, tmItemResult, tmServerStatus,
    tmDWORD, tmPOleStr, tmVarType, tmOleVariant, tmWord, tmFileTime,
    tmItemAttribute, tmLCID);

function taskMemAlloc(dwCount: DWORD; tm: TTaskMemAllocKind): pointer;

implementation

function DataTimeToOPCTime(time: TDateTime): TFileTime;
var
  sysTime: TSystemTime;
begin
  DateTimeToSystemTime(time, sysTime);
  SystemTimeToFileTime(sysTime, result);
  LocalFileTimeToFileTime(result, result);
end;

function ConvertVariant(const value: variant; requestedType: TVarType): variant;
begin
  try
    result := VarAsType(value, requestedType);
  except
    on EVariantError do result := DISP_E_TYPEMISMATCH;
  end;
end;

const
  allocSizes : array [TTaskMemAllocKind] of longword = (
    sizeof(HRESULT), sizeof(OPCITEMSTATE), sizeof(OPCITEMRESULT),
    sizeof(OPCSERVERSTATUS), sizeof(DWORD), sizeof(POleStr), sizeof(TVarType),
    sizeof(OleVariant), sizeof(word), sizeof(TFileTime), sizeof(OPCITEMATTRIBUTES),
    sizeof(LCID));

function taskMemAlloc(dwCount: DWORD; tm: TTaskMemAllocKind): pointer;
var
  size: DWORD;
begin
  try
    size := dwCount * allocSizes[tm];
    result := CoTaskMemAlloc(size);
    if result <> nil then fillChar(result^, size, 0);
  except
    result := nil;
  end;
end;

end.
