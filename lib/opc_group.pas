unit opc_group;

{$IFDEF VER150}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses windows, classes, sysutils, activex, axctrls, comobj, dialogs, syncobjs,
  opc_types, opc_error, opc_da, opc_server, opc_utils;

type
  TOPCGroup = class(TTypedComObject, IOPCItemMgt, IOPCGroupStateMgt,
    IOPCPublicGroupStateMgt, IOPCSyncIO, IConnectionPointContainer, IOPCAsyncIO2)
  private
    _connectionPoints: TConnectionPoints;
  protected
    property iFIConnectionPoints: TConnectionPoints read _connectionPoints
      write _connectionPoints implements IConnectionPointContainer;

    //IOPCItemMgt
    function AddItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
      out ppAddResults: POPCITEMRESULTARRAY; out ppErrors: PResultList): HResult; stdcall;
    function ValidateItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
      bBlobUpdate: BOOL; out ppValidationResults: POPCITEMRESULTARRAY;
      out ppErrors: PResultList): HResult; stdcall;
    function RemoveItems(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      out ppErrors: PResultList): HResult; stdcall;
    function SetActiveState(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      bActive: BOOL; out ppErrors: PResultList): HResult; stdcall;
    function SetClientHandles(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      phClient: POPCHANDLEARRAY; out ppErrors: PResultList): HResult; stdcall;
    function SetDatatypes(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      pRequestedDatatypes: PVarTypeList; out ppErrors: PResultList): HResult; stdcall;
    function CreateEnumerator(const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;

    //IOPCGroupStateMgt
    function GetState(out pUpdateRate: DWORD; out pActive: BOOL;
      out ppName: POleStr; out pTimeBias: Longint; out pPercentDeadband: Single;
      out pLCID: TLCID; out phClientGroup: OPCHANDLE;
      out phServerGroup: OPCHANDLE): HResult; overload; stdcall;
    function SetState(pRequestedUpdateRate: PDWORD;
      out pRevisedUpdateRate: DWORD; pActive: PBOOL; pTimeBias: PLongint;
      pPercentDeadband: PSingle; pLCID: PLCID; phClientGroup: POPCHANDLE): HResult; stdcall;
    function SetName(szName: POleStr): HResult; stdcall;
    function CloneGroup(szName: POleStr; const riid: TIID;
      out ppUnk: IUnknown): HResult;stdcall;

    //IOPCPublicGroupStateMgt
    function GetState(out pPublic: BOOL): HResult; overload; stdcall;
    function MoveToPublic: HResult; stdcall;

    //IOPCSyncIO
    function Read(dwSource: OPCDATASOURCE; dwCount: DWORD;
      phServer: POPCHANDLEARRAY; out ppItemValues: POPCITEMSTATEARRAY;
      out ppErrors: PResultList): HResult; overload; stdcall;
    function Write(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      pItemValues: POleVariantArray; out ppErrors: PResultList): HResult; overload; stdcall;

    //IOPCAsyncIO2
    function Read(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      dwTransactionID: DWORD; out pdwCancelID: DWORD;
      out ppErrors: PResultList): HResult; overload; stdcall;
    function Write(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      pItemValues: POleVariantArray; dwTransactionID: DWORD;
      out pdwCancelID: DWORD; out ppErrors: PResultList): HResult; overload; stdcall;
    function Refresh2(dwSource: OPCDATASOURCE; dwTransactionID: DWORD;
      out pdwCancelID: DWORD): HResult; stdcall;
    function Cancel2(dwCancelID: DWORD): HResult; stdcall;
    function SetEnable(bEnable: BOOL): HResult; stdcall;
    function GetEnable(out pbEnable: BOOL): HResult; stdcall;
  private
    _lock: TCriticalSection;
    _server: TDA2; //the owner
    _name: string; //the name of this group
    _clientHandle: longword; //the client generates we pass to client
    _serverHandle: longword; //we generate the client will passes to us
    _requestedUpdateRate: longword; //update rate in mills
    _LCID: longword; //lanugage id

    _items, _asyncList: TList;
    _active, _public, _onDataChangeEnabled: longbool;
    _clientSink: IUnknown;
    _lastUpdate: longword;

    _updateList: array of OPCHANDLE;
    _updateAlloc, _updateCount: integer;

    _lastCancelId: longword;

  public
    constructor create(server: TDA2; name: string; active: BOOL;
      requestedUpdateRate: DWORD; clientHandle: OPCHANDLE; LCID: DWORD);
    constructor clone(source: TOPCGroup; name: string);
    destructor destroy; override;

    function validateRequestedUpDateRate(requestedUpdateRate: DWORD): DWORD;
    procedure callBackOnConnect(const sink: IUnknown; connecting: boolean);
    function findItemByServerHandle(handle: OPCHANDLE): pointer;
    function findItemByClientHandle(handle: OPCHANDLE): pointer;
    function getActiveItems: TList;
    function makeAsyncCancelID: longword;
    procedure scan;
    function isItemMarkedForUpdate(item: pointer): boolean;
    procedure addUpdatedItem(item: pointer);
    procedure asyncScan;

    function clientSink: IUnknown;
    function items: TList;
    function clientHandle: OPCHANDLE;
    function serverHandle: OPCHANDLE;
    function requestedUpdateRate: DWORD;
    function name: string;
    function isPublic: boolean;
  end;

implementation

uses comserv, opc_item, opc_async, opc_enum, opc_item_proxy;

function isVariantTypeOK(vType: integer): boolean;
begin
  result := boolean(vType in [varEmpty..$14]);
end;

////////////////////////////////////////////////////////////////////////////////

function TOPCGroup.AddItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
  out ppAddResults: POPCITEMRESULTARRAY; out ppErrors: PResultList): HResult;

  procedure ClearResultsArray;
  var
    i: integer;
  begin
    for i := 0 to dwCount - 1 do begin
      ppAddResults[i].hServer := 0;
      ppAddResults[i].vtCanonicalDataType := 0;
      ppAddResults[i].wReserved := 0;
      ppAddResults[i].dwAccessRights := 0;
      ppAddResults[i].dwBlobSize := 0;
      ppAddResults[i].pBlob := nil;
    end;
  end;

var
  i: integer;
  item: TOPCItem;
  inItemDef: POPCITEMDEF;
  proxy: TOPCItemProxy;
begin
  result := S_OK;
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppAddResults := taskMemAlloc(dwCount, tmItemResult);
  ppErrors := taskMemAlloc(dwCount, tmHResult);

  if (ppAddResults = nil) or (ppErrors = nil) then begin
    if ppAddResults <> nil then CoTaskMemFree(ppAddResults);
    if ppErrors <> nil then CoTaskMemFree(ppErrors);
    result := E_OUTOFMEMORY;
    exit;
  end;

  ClearResultsArray;

  _lock.acquire;
  try
    for i := 0 to dwCount - 1 do begin
      ppErrors[i] := S_OK;
      inItemDef := @pItemArray[i];
      if length(inItemDef.szItemID) = 0 then begin
        result := S_FALSE; ppErrors[i] := OPC_E_INVALIDITEMID; continue;
      end;
      proxy := _server.findProxy(inItemDef.szItemID);
      if proxy = nil then begin
        result := S_FALSE; ppErrors[i] := OPC_E_UNKNOWNITEMID; continue;
      end;
      if not IsVariantTypeOK(inItemDef.vtRequestedDataType) then begin
        result := S_FALSE; ppErrors[i] := OPC_E_BADTYPE; continue;
      end;
      item := TOPCItem.create(_server, self, proxy, inItemDef.szItemID,
        inItemDef.hClient, inItemDef.vtRequestedDataType, inItemDef.bActive);
      _items.add(item);
      ppAddResults[i].hServer := item.serverHandle;
      ppAddResults[i].vtCanonicalDataType := item.canonicalDataType;
      ppAddResults[i].dwAccessRights := item.accessRights;
      ppAddResults[i].dwBlobSize := 0;
      ppAddResults[i].pBlob := nil;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.ValidateItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
  bBlobUpdate: BOOL; out ppValidationResults: POPCITEMRESULTARRAY;
  out ppErrors: PResultList): HResult;

  procedure ClearResultsArray;
  var
    i: integer;
  begin
    for i := 0 to dwCount - 1 do begin
      ppValidationResults[i].vtCanonicalDataType := 0;
      ppValidationResults[i].dwAccessRights := 0;
      ppValidationResults[i].dwBlobSize := 0;
      ppValidationResults[i].hServer := 0;
    end;
  end;

var
  i: integer;
  inItemDef: POPCITEMDEF;
  proxy: TOPCItemProxy;
begin
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppValidationResults := taskMemAlloc(dwCount, tmItemResult);
  ppErrors := taskMemAlloc(dwCount, tmHResult);

  if (ppValidationResults = nil) or (ppErrors = nil) then begin
    if ppValidationResults <> nil then  CoTaskMemFree(ppValidationResults);
    if ppErrors <> nil then  CoTaskMemFree(ppErrors);
    result := E_OUTOFMEMORY;
    exit;
  end;

  result := S_OK;
  ClearResultsArray;
  _lock.acquire;
  try
    for i := 0 to dwCount - 1 do begin
      inItemDef := @pItemArray[i];

      proxy := _server.findProxy(inItemDef.szItemID);
      if proxy = nil then begin
        result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDITEMID;
        continue;
      end;

      ppValidationResults[i].vtCanonicalDataType := proxy.datatype;
      ppValidationResults[i].dwAccessRights := proxy.accessRights;
      ppValidationResults[i].dwBlobSize := 0;
      ppErrors[i] := S_OK;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.RemoveItems(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  out ppErrors: PResultList): HResult;
var
  i: integer;
  item: TOPCItem;
begin
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppErrors := taskMemAlloc(dwCount, tmHResult);
  if ppErrors = nil then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    result := S_OK;
    for i:= 0 to dwCount -1 do begin
      item := findItemByServerHandle(phServer[i]);
      if item <> nil then begin
        _items.remove(item);
        item.free;
        ppErrors[i] := S_OK;
      end else begin
        result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  finally _lock.release; end;
end;


function TOPCGroup.SetActiveState(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  bActive: BOOL; out ppErrors: PResultList): HResult;
var
  i: integer;
  item: TOPCItem;
begin
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppErrors := taskMemAlloc(dwCount, tmHResult);
  if ppErrors = nil then begin
    result := E_OUTOFMEMORY;
    Exit;
  end;

  _lock.acquire;
  try
    result:=S_OK;
    for i:= 0 to dwCount - 1 do begin
      item := findItemByServerHandle(phServer[i]);
      if item <> nil then begin
        item.setActive(bActive);
        ppErrors[i] := S_OK;
      end else begin
        result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.SetClientHandles(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  phClient: POPCHANDLEARRAY; out ppErrors: PResultList): HResult;
var
  i: integer;
  item: TOPCItem;
begin
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppErrors := taskMemAlloc(dwCount, tmHResult);
  if ppErrors = nil then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    result := S_OK;
    for i := 0 to dwCount - 1 do begin
      item := findItemByServerHandle(phServer[i]);
      if item <> nil then begin
        item.setClientHandle(phClient[i]);
        ppErrors[i] := S_OK;
     end else begin
       result := S_FALSE;
       ppErrors[i] := OPC_E_INVALIDHANDLE;
     end;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.SetDatatypes(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pRequestedDatatypes: PVarTypeList; out ppErrors: PResultList): HResult;
var
  i: integer;
  item: TOPCItem;
begin
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppErrors := taskMemAlloc(dwCount, tmHResult);
  if ppErrors = nil then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    result := S_OK;
    for i := 0 to dwCount - 1 do begin
      item := findItemByServerHandle(phServer[i]);
      if item <> nil then begin
        if not IsVariantTypeOK(pRequestedDatatypes[i]) then
          ppErrors[i] := OPC_E_BADTYPE
        else begin
          item.setRequestedDataType(pRequestedDatatypes[i]);
          ppErrors[i] := S_OK;
        end;
      end else begin
        result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.CreateEnumerator(const riid: TIID;
  out ppUnk: IUnknown): HResult;
var
  i: integer;
  list: TList;
begin
  if (_items = nil) or (_items.count = 0) then begin
    result:=S_FALSE;
    exit;
  end;

  list := TList.Create;
  if list = nil then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    for i := 0 to _items.count - 1 do
      list.Add(TOPCItemAttributes.create(_items[i]));
  finally _lock.release; end;

  ppUnk := TOPCItemAttEnumerator.Create(list);
  result := S_OK;
end;

function TOPCGroup.GetState(out pUpdateRate: DWORD; out pActive: BOOL;
  out ppName: POleStr; out pTimeBias: Longint; out pPercentDeadband: Single;
  out pLCID: TLCID; out phClientGroup: OPCHANDLE;
  out phServerGroup: OPCHANDLE): HResult;
begin
  _lock.acquire;
  try
    pUpdateRate := _requestedUpdateRate;
    pActive := _active;
    ppName := StringToLPOLESTR(_name);
    pTimeBias := 0;
    pPercentDeadband := 0;
    pLCID := _LCID;
    phClientGroup := _clientHandle;
    phServerGroup := _serverHandle;
    result := S_OK;
  finally _lock.release; end;
end;

function TOPCGroup.SetState(pRequestedUpdateRate: PDWORD;
  out pRevisedUpdateRate: DWORD; pActive: PBOOL; pTimeBias: PLongint;
  pPercentDeadband: PSingle; pLCID: PLCID; phClientGroup: POPCHANDLE):HResult;
begin
  result := S_OK;
  if assigned(pRequestedUpdateRate) then
    ValidateRequestedUpdateRate(pRequestedUpdateRate^);

  _lock.acquire;
  try
    if assigned(pActive) then
      _active := pActive^;
    if assigned(pLCID) then
      _LCID := pLCID^;
    if assigned(phClientGroup) then
      _clientHandle := phClientGroup^;

    if (addr(pRevisedUpdateRate) <> nil) then
      pRevisedUpdateRate := _requestedUpdateRate;
  finally _lock.release; end;
end;

function TOPCGroup.SetName(szName: POleStr): HResult;
begin
  result := S_OK;
  if length(szName) = 0 then begin
    result:=E_INVALIDARG;
    exit;
  end;
  _lock.acquire;
  try
    if _server.findGroupByName(szName) <> nil then begin
      result := OPC_E_DUPLICATENAME;
      exit;
    end;
    _name := szName;
  finally _lock.release; end;
end;

function TOPCGroup.CloneGroup(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult;
var
  s1: string;
  i: integer;
begin
  if not (IsEqualIID(riid, IID_IOPCGroupStateMgt) or IsEqualIID(riid, IID_IUnknown)) then begin
    result := E_NOINTERFACE;
    exit;
  end;

  s1 := szName;
  if (length(s1) <> 0) and (_server.findGroupByName(s1) <> nil) then begin
    result := OPC_E_DUPLICATENAME;
    Exit;
  end;

  _lock.acquire;
  try
    i := 0;
    while _server.findGroupByName(s1) <> nil do begin
      inc(i);
      s1 := _name + inttostr(i);
    end;
    ppUnk := TOPCGroup.clone(self, s1);
    if ppUnk = nil then
      result := E_OUTOFMEMORY
    else
      result := S_OK;
  finally _lock.release; end;
end;

function TOPCGroup.GetState(out pPublic: BOOL): HResult;
begin
  pPublic := _public;
  result := S_OK;
end;

function TOPCGroup.MoveToPublic: HResult;
begin
  _public := true;
  result := S_OK;
end;

function TOPCGroup.Read(dwSource: OPCDATASOURCE; dwCount: DWORD;
  phServer: POPCHANDLEARRAY; out ppItemValues: POPCITEMSTATEARRAY;
  out ppErrors: PResultList): HResult;
var
  i: integer;
  ppServer: PDWORDARRAY;
  item: TOPCItem;

  procedure ClearResultsArray;
  var
    i: integer;
  begin
    for i := 0 to dwCount - 1 do begin
      ppItemValues[i].hClient := 0;
      ppItemValues[i].wReserved := 0;
      ppItemValues[i].vDataValue := 0;
      ppItemValues[i].wQuality := 0;
    end;
  end;

begin
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppItemValues := taskMemAlloc(dwCount, tmItemState);
  ppErrors := taskMemAlloc(dwCount, tmHResult);

  if (ppItemValues = nil) or (ppErrors = nil) then begin
    if ppItemValues <> nil then  CoTaskMemFree(ppItemValues);
    if ppErrors <> nil then  CoTaskMemFree(ppErrors);
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    result := S_OK;
    ppServer := @phServer^;
    ClearResultsArray;
    for i:= 0 to dwCount -1 do begin
      item := findItemByServerHandle(ppServer[i]);
      if item <> nil then begin
        item.ReadItemValueStateTime(dwSource, ppItemValues[i]);
        if (dwSource <> OPC_DS_DEVICE) and not _active then
          ppItemValues[i].wQuality := OPC_QUALITY_OUT_OF_SERVICE;
        ppErrors[i] := S_OK;
      end else begin
        result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.Write(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pItemValues: POleVariantArray; out ppErrors: PResultList): HResult;
var
  i: integer;
  ppServer: PDWORDARRAY;
  item: TOPCItem;
begin
  if dwCount < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppErrors := taskMemAlloc(dwCount, tmHResult);
  if ppErrors = nil then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    result := S_OK;
    ppServer := @phServer^;
    for i := 0 to dwCount - 1 do begin
      item := findItemByServerHandle(ppServer[i]);
      if item <> nil then begin
        if not item.writeable then begin
          ppErrors[i] := OPC_E_BADRIGHTS;
          result := S_FALSE;
        end else begin
          item.writeItemValue(pItemValues[i]);
          ppErrors[i] := S_OK
        end;
      end else begin
        result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.Read(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  dwTransactionID: DWORD; out pdwCancelID: DWORD; out ppErrors: PResultList): HResult;
var
  i: integer;
  asyncObj: TAsyncIO2;
begin
  if _clientSink = nil then begin
    result:=CONNECT_E_NOCONNECTION;
    exit;
  end;

  if (dwCount < 1) then begin
    result := E_INVALIDARG;
    exit;
  end;

  ppErrors := taskMemAlloc(dwCount, tmHResult);
  asyncObj := TAsyncIO2.Create(self, io2read, dwTransactionID);

  if (ppErrors = nil) or (asyncObj = nil) then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    pdwCancelID := asyncObj.cancelID;
    for i := 0 to dwCount - 1 do begin
      if findItemByServerHandle(phServer[i]) <> nil then begin
        asyncObj.addItem;
        asyncObj.setServerHandle(phServer[i]);
        ppErrors[i] := S_OK;
      end else
        ppErrors[i] := OPC_E_INVALIDHANDLE;
    end;
    _asyncList.Add(asyncObj);
    result := S_OK;
  finally _lock.release; end;
end;

function TOPCGroup.Write(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pItemValues: POleVariantArray; dwTransactionID: DWORD; out pdwCancelID: DWORD;
  out ppErrors: PResultList): HResult;
var
  i: longword;
  asyncObj: TAsyncIO2;
begin
  if _clientSink = nil then begin
    result := CONNECT_E_NOCONNECTION;
    exit;
  end;

  if (dwCount < 1) then begin
    result := E_INVALIDARG;
    exit;
  end;

  asyncObj := TAsyncIO2.Create(self, io2write, dwTransactionId);
  ppErrors := taskMemAlloc(dwCount, tmHResult);

  if (asyncObj = nil) or (ppErrors = nil) then begin
    result := E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    pdwCancelID := asyncObj.cancelID;
    for i := 0 to dwCount - 1 do begin
      if findItemByServerHandle(phServer[i]) <> nil then begin
        asyncObj.addItem;
        asyncObj.setServerHandle(phServer[i]);
        asyncObj.setValues(pItemValues[i]);
        ppErrors[i] := S_OK;
      end else
        ppErrors[i] := OPC_E_INVALIDHANDLE;
    end;
    _asyncList.Add(asyncObj);
    result := S_OK;
  finally _lock.release; end;
end;

function TOPCGroup.Refresh2(dwSource: OPCDATASOURCE; dwTransactionID: DWORD;
  out pdwCancelID: DWORD): HResult;
var
  asyncObj: TAsyncIO2;
begin
  if _clientSink = nil then begin
    result := CONNECT_E_NOCONNECTION;
    Exit;
  end;
  if (not _active) then begin
    result := E_FAIL;
    Exit;
  end;

  _lock.acquire;
  try
    asyncObj := TAsyncIO2.Create(self, io2refresh, dwTransactionId);
    if (asyncObj = nil) then begin
      result := E_OUTOFMEMORY;
      exit;
    end;
    pdwCancelID := asyncObj.cancelID;
    _asyncList.Add(asyncObj);
    result := S_OK;
  finally _lock.release; end;
end;

function TOPCGroup.Cancel2(dwCancelID: DWORD): HResult;
var
  i: integer;
begin
  result := E_FAIL;
  _lock.acquire;
  try
    if (_asyncList = nil) or (_asyncList.count = 0) then exit;
    for i := 0 to _asyncList.count - 1 do
      if TAsyncIO2(_asyncList[i]).cancelID = dwCancelID then begin
        result := S_OK;
        TAsyncIO2(_asyncList[i]).cancel;
        break;
      end;
  finally _lock.release; end;
end;

function TOPCGroup.SetEnable(bEnable: BOOL): HResult;
begin
  if _clientSink = nil then begin
    result:=CONNECT_E_NOCONNECTION;
    exit;
  end;
  _lock.acquire;
  try
    _onDataChangeEnabled := bEnable;
    result := S_OK;
  finally _lock.release; end;
end;

function TOPCGroup.GetEnable(out pbEnable: BOOL): HResult;
begin
  if _clientSink = nil then begin
    result := CONNECT_E_NOCONNECTION;
    exit;
  end;
  _lock.acquire;
  try
    pbEnable := _onDataChangeEnabled;
    result:=S_OK;
  finally _lock.release; end;
end;

constructor TOPCGroup.create(server: TDA2; name: string; active: BOOL;
  requestedUpdateRate: DWORD; clientHandle: OPCHANDLE; LCID: DWORD);
begin
  inherited create;
  _lock := TCriticalSection.create;
  _connectionPoints := TConnectionPoints.create(self);
  _connectionPoints.CreateConnectionPoint(IID_IOPCDataCallback, ckMulti,
    CallBackOnConnect);
  _server := server;
  _name := name;
  _active := active;
  ValidateRequestedUpDateRate(requestedUpdateRate);
  _clientHandle := clientHandle;
  _LCID := LCID;
  _serverHandle := _server.makeGroupServerHandle;
  _public := false;
  _server.touch;
  _onDataChangeEnabled := true;

  _items := TList.Create;
  _asyncList := TList.Create;
  _updateAlloc := 64;
  setLength(_updateList, _updateAlloc);
  _onDataChangeEnabled := true;

  _server.addGroupRef(self);
end;

constructor TOPCGroup.clone(source: TOPCGroup; name: string);
var
  i: integer;
begin
  inherited create;
  _lock := TCriticalSection.create;
  _connectionPoints := TConnectionPoints.create(self);
  _server := source._server;
  _name := name;
  _active := source._active;
  ValidateRequestedUpDateRate(source._requestedUpdateRate);
  _clientHandle := source._clientHandle;
  _LCID := source._LCID;
  _serverHandle := _server.makeGroupServerHandle;
  _public := source._public;
  _server.touch;
  _onDataChangeEnabled := true;
  _items := TList.Create;
  _asyncList := TList.Create;
  _updateAlloc := 64;
  setLength(_updateList, _updateAlloc);
  _onDataChangeEnabled := true;

  for i:=0 to source._items.count-1 do
    _items.Add(TOPCItem.clone(source._items[i]));

  _server.addGroupRef(self);
end;

destructor TOPCGroup.Destroy;
var
  i: integer;
begin
  for i := _items.count - 1 downto 0 do
    TOPCItem(_items[i]).free;
  _items.free;
  for i := _asyncList.count - 1 downto 0 do
    TAsyncIO2(_asyncList[i]).free;
  _asyncList.free;
  _connectionPoints.Free;
  setlength(_updateList, 0);
  _server.removeGroupRef(self);
  _lock.free;

  inherited destroy;
end;

function TOPCGroup.ValidateRequestedUpDateRate(requestedUpdateRate: DWORD): DWORD;
begin
  if (requestedUpdateRate < 10) then
    _requestedUpdateRate := 10
  else
    _requestedUpdateRate := requestedUpdateRate;
  result := _requestedUpdateRate;
end;

procedure TOPCGroup.callBackOnConnect(const sink: IUnknown; connecting: boolean);
begin
  if connecting then
    _clientSink := sink
  else
    _clientSink := nil;
end;

function TOPCGroup.findItemByServerHandle(handle: OPCHANDLE): pointer;
var
  i: integer;
begin
  for i := 0 to _items.count - 1 do
    if TOPCItem(_items[i]).serverHandle = handle then begin
      result := _items[i];
      exit;
    end;
  result := nil;
end;

function TOPCGroup.findItemByClientHandle(handle: OPCHANDLE): pointer;
var
  i: integer;
begin
  for i := 0 to _items.count - 1 do
    if TOPCItem(_items[i]).clientHandle = handle then begin
      result := _items[i];
      exit;
    end;
  result := nil;
end;

function TOPCGroup.getActiveItems: TList;
var
  i: integer;
begin
  result := TList.create;
  for i := 0 to _items.count - 1 do
    if TOPCItem(_items[i]).active then result.add(_items[i]);
end;

function TOPCGroup.makeAsyncCancelID: longword;
begin
  _lock.acquire;
  try
    inc(_lastCancelId);
    result := _lastCancelId;
  finally
    _lock.release;
  end;
end;

procedure TOPCGroup.scan;
begin
  _lock.acquire;
  try
    if (not _active) or (_items = nil) or (_items.count = 0) then exit;
    if getTickCount - _lastUpdate > _requestedUpdateRate then begin
      _lastUpdate := getTickCount;
      asyncScan;
      if (_clientSink <> nil) and _onDataChangeEnabled and (_updateCount > 0) then begin
        asyncOnChange(self, 0, _updateList, _updateCount);
        _updateCount := 0;
      end;
    end;
  finally _lock.release; end;
end;

function TOPCGroup.isItemMarkedForUpdate(item: pointer): boolean;
var
  i: integer;
begin
  for i := 0 to _updateCount - 1 do
    if _updateList[i] = TOPCItem(item).serverHandle then begin
      result := true;
      exit;
    end;
  result := false;
end;

procedure TOPCGroup.addUpdatedItem(item: pointer);
begin
  _lock.acquire;
  try
    if isItemMarkedForUpdate(item) then exit;
    inc(_updateCount);
    if _updateCount > _updateAlloc then begin
      _updateAlloc := _updateCount + 32;
      setlength(_updateList, _updateAlloc);
    end;
    _updateList[_updateCount - 1] := TOPCItem(item).serverHandle;
  finally
    _lock.release;
  end;
end;

procedure TOPCGroup.asyncScan;
var
  i: integer;
  asyncObj: TAsyncIO2;
begin
  if (_requestedUpdateRate = 0) or (_items.count = 0) or
    (_clientSink =  nil) then begin
    for i := _asyncList.count - 1 downTo 0 do
      TAsyncIO2(_asyncList[i]).Free;
    exit;
  end;

  asyncObj := nil;
  if (_asyncList <> nil) and (_asyncList.count > 0) then begin
    for i:= 0 to _asyncList.count - 1 do try
      asyncObj := TAsyncIO2(_asyncList[i]);
      asyncObj.scan;
    finally
      FreeAndNil(asyncObj);
    end;
    _asyncList.Clear;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

function TOPCGroup.clientSink: IUnknown;
begin
  result := _clientSink;
end;

function TOPCGroup.items: TList;
begin
  result := _items;
end;

function TOPCGroup.clientHandle: OPCHANDLE;
begin
  result := _clientHandle;
end;

function TOPCGroup.serverHandle: OPCHANDLE;
begin
  result := _serverHandle;
end;

function TOPCGroup.requestedUpdateRate: DWORD;
begin
  result := _requestedUpdateRate;
end;

function TOPCGroup.name: string;
begin
  result := _name;
end;

function TOPCGroup.isPublic: boolean;
begin
  result := _public;
end;

////////////////////////////////////////////////////////////////////////////////

end.
