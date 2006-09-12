unit opc_enum;

interface

uses windows, classes, comobj, activex, comserv, opc_da;


type
  TOPCStringsEnumerator = class(TComObject, IEnumString)
  private
    _next: integer;
    _list: TStringList;
  public
    constructor create(list: TStringList);
    destructor destroy;override;
    function Next(celt: longint; out elt; pceltFetched: plongint): HResult; stdcall;
    function Skip(celt: longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumString): HResult; stdcall;
  end;

  TS3UnknownEnumerator = class(TComObject, IEnumUnknown)
  private
    _next: integer;
    _list: TList;
  public
    constructor create(list: TList);
    destructor destroy; override;
    function Next(celt: longint; out elt; pceltFetched: plongint): HResult; stdcall;
    function Skip(celt: longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumUnknown): HResult; stdcall;
  end;

  TOPCItemAttEnumerator = class(TComObject, IEnumOPCItemAttributes)
  private
    _next: longword;
    _list: TList;
  public
    constructor create(list: TList);
    destructor destroy; override;
    procedure PopulateRecord(var rec: OPCITEMATTRIBUTES; i: integer);
    function Next(celt: cardinal; out ppItemArray: POPCITEMATTRIBUTESARRAY;
      out pceltFetched: cardinal): HResult; stdcall;
    function Skip(celt: cardinal): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out ppEnumItemAttributes: IEnumOPCItemAttributes): HResult; stdcall;
  end;

implementation

uses opc_item, opc_utils;

const
 IID_IUnknown: TIID = '{00000000-0000-0000-C000-000000000046}';  //is in ole2.pas

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

constructor TOPCStringsEnumerator.create(list: TStringList);
begin
  inherited create;
  _list := TStringList.create;
  _list.AddStrings(list);
end;

destructor TOPCStringsEnumerator.destroy;
begin
  _list.free;
  inherited destroy;
end;

function TOPCStringsEnumerator.Next(celt: longint; out elt;
  pceltFetched: plongint): HResult;
var
  i: integer;
begin
  i := 0;
  if celt < 1 then begin
    result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
    exit;
  end;
  if pceltFetched = nil then begin
    result := E_INVALIDARG;
    exit;
  end;

  result := S_FALSE;
  while (i < celt) do begin
    if (_next < _list.Count) then begin
      TPointerList(elt)[i] := StringToLPOLESTR(_list[_next]);
      i := succ(i);
      _next := succ(_next);
    end else begin
      result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
      break;
    end;
  end;

  pceltFetched^ := i;
  if i = celt then result := S_OK;
end;

function TOPCStringsEnumerator.Skip(celt: longint): HResult;
begin
  if (_next + celt) <= _list.Count then begin
    _next := _next + celt;
    result := S_OK;
  end else begin
    _next := _list.count;
    result := S_FALSE;
  end;
end;

function TOPCStringsEnumerator.Reset: HResult;
begin
  _next := 0;
  result := S_OK;
end;

function TOPCStringsEnumerator.Clone(out enm: IEnumString): HResult;
begin
  try
    enm := TOPCStringsEnumerator.Create(_list);
    result := S_OK;
  except
    result := E_UNEXPECTED;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

constructor TS3UnknownEnumerator.create(list: TList);
begin
  inherited create;
  _list := TList.Create;
  _list.Assign(list);
end;

destructor TS3UnknownEnumerator.destroy;
begin
  _list.Free;
  inherited Destroy;
end;

function TS3UnknownEnumerator.Next(celt: longint; out elt;
  pceltFetched: plongint): HResult;
var
  i: integer;
begin
  i := 0;
  if celt < 1 then begin
    result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
    exit;
  end;

  if pceltFetched = nil then begin
    result := E_INVALIDARG;
    exit;
  end;

  result := S_FALSE;
  while (i < celt) do begin
    if (_next < _list.count) then begin
      TPointerList(elt)[i] := _list[_next];
      i := succ(i);
      _next := succ(_next);
    end else begin
      result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
      break;
    end;
  end;

  pceltFetched^ := i;
  if i = celt then result := S_OK;
end;

function TS3UnknownEnumerator.Skip(celt: longint): HResult;
begin
  if (_next + celt) <= _list.Count then begin
    _next := _next + celt;
    result:=S_OK;
  end else begin
    _next := _list.count;
    result := S_FALSE;
  end;
end;

function TS3UnknownEnumerator.Reset: HResult;
begin
  _next := 0;
  result := S_OK;
end;

function TS3UnknownEnumerator.Clone(out enm: IEnumUnknown): HResult;
begin
  try
    enm := TS3UnknownEnumerator.Create(_list);
    result := S_OK;
  except
    result := E_UNEXPECTED;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

constructor TOPCItemAttEnumerator.create(list: TList);
begin
  inherited create;
  _list := list;
end;

destructor TOPCItemAttEnumerator.destroy;
var
  i: integer;
begin
  for i := 0 to _list.count - 1 do
    TOPCItemAttributes(_list[i]).free;
  _list.free;
  inherited destroy;
end;

procedure TOPCItemAttEnumerator.PopulateRecord(var rec: OPCITEMATTRIBUTES;
  i: integer);
begin
  with TOPCItemAttributes(_list[i]) do begin
    rec.szAccessPath:=StringToLPOLESTR(_accessPath);
    rec.szItemID:=StringToLPOLESTR(_itemID);
    rec.bActive:=_active;
    rec.hClient:=_clientHandle;
    rec.hServer:=_serverHandle;
    rec.dwAccessRights:=_accessRights;
    rec.dwBlobSize:=0;
    rec.pBlob:=nil;
    rec.vtRequestedDataType:=_requestedDataType;
    rec.vtCanonicalDataType:=_canonicalDataType;
    rec.dwEUType:=_euType;
    rec.vEUInfo:=_euInfo;
  end;
end;

function TOPCItemAttEnumerator.Next(celt: cardinal;
  out ppItemArray: POPCITEMATTRIBUTESARRAY; out pceltFetched: cardinal):HResult;
var
  i: cardinal;
begin
  i := 0;
  pceltFetched := i;
  if celt < 1 then begin
    result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
    exit;
  end;

  result:=E_FAIL;
  ppItemArray := taskMemAlloc(celt, tmItemAttribute);
  if ppItemArray = nil then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  while (i < celt) do begin
    if (_next < cardinal(_list.count)) then begin
      PopulateRecord(ppItemArray[i], _next);
      i := succ(i);
      _next := succ(_next);
    end else begin
      result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
      break;
    end;
  end;

  pceltFetched := i;
  if i = celt then result := S_OK;
end;

function TOPCItemAttEnumerator.Skip(celt: Cardinal): HResult;
begin
  if (_next + celt) <= cardinal(_list.count) then begin
    _next := _next + celt;
    result := S_OK;
  end else begin
    _next := _list.count;
    result := S_FALSE;
  end;
end;

function TOPCItemAttEnumerator.Reset: HResult;
begin
  _next := 0;
  result := S_OK;
end;

function TOPCItemAttEnumerator.Clone(
  out ppEnumItemAttributes: IEnumOPCItemAttributes): HResult;
begin
  try
    ppEnumItemAttributes := TOPCItemAttEnumerator.Create(_list);
    result := S_OK;
  except
    result := E_UNEXPECTED;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

end.
