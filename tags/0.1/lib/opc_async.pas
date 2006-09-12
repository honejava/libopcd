unit opc_async;

{$IFDEF VER150}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses windows, classes, sysutils, dialogs, comobj, activex, axctrls,
  opc_types, opc_error, opc_da, opc_server, opc_group, opc_item, opc_utils;

type
  TAsyncIO2 = class
  private
    _group: TOPCGroup;
    _kind, _clientTransId, _count, _alloc, _cancelId: longword;
    _serverHandles: array of OPCHANDLE;
    _values: array of variant;
    _cancelled: boolean;
  public
    constructor create(group: TOPCGroup; kind, clientTransId: longword);
    destructor destroy; override;

    procedure addItem;
    procedure setServerHandle(handle: OPCHANDLE);
    procedure setValues(const value: variant);

    procedure scan;
    procedure cancel;

    function cancelId: longword;
  end;

procedure asyncRead(group: TOPCGroup; clientTransId: longword;
  serverHandles: array of OPCHANDLE; count: integer);
procedure asyncWrite(group: TOPCGroup; clientTransId: longword;
  serverHandles: array of OPCHANDLE; values: array of variant; count: integer);
procedure asyncRefresh(group: TOPCGroup; clientTransId: longword);
procedure asyncOnChange(group: TOPCGroup; clientTransId: longword;
  serverHandles: array of OPCHANDLE; count: integer);

implementation

type
  WORDARRAY = array[0..65535] of WORD;
  PWORDARRAY = ^WORDARRAY;

type
  TFileTimeARRAY = array[0..65535] of TFileTime;
  PTFileTimeARRAY = ^TFileTimeARRAY;

constructor TAsyncIO2.create(group: TOPCGroup; kind, clientTransId: longword);
begin
  inherited create;
  _group := group;
  _kind := kind;
  _clientTransId := clientTransId;
  _cancelID := group.makeAsyncCancelID;
end;

destructor TAsyncIO2.destroy;
begin
  setLength(_serverHandles, 0);
  setLength(_values, 0);
  inherited destroy;
end;

procedure TAsyncIO2.addItem;
begin
  inc(_count);
  if _count > _alloc then begin
    _alloc := _count + 32;
    setLength(_serverHandles, _alloc);
    setLength(_values, _alloc);
  end;
end;

procedure TAsyncIO2.setServerHandle(handle: OPCHANDLE);
begin
  _serverHandles[_count - 1] := handle;
end;

procedure TAsyncIO2.setValues(const value: variant);
begin
  _values[_count - 1] := value;
end;

procedure TAsyncIO2.scan;
begin
  if _cancelled then exit;
  case _kind of
    io2Read: asyncRead(_group, _clientTransId, _serverHandles, _count);
    io2Write: asyncWrite(_group, _clientTransId, _serverHandles, _values, _count);
    io2Refresh: asyncRefresh(_group, _clientTransId);
  end;
end;

procedure TAsyncIO2.cancel;
begin
  _cancelled := true;
end;

function TAsyncIO2.cancelId: longword;
begin
  result := _cancelId;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

procedure asyncRead(group: TOPCGroup; clientTransId: longword;
  serverHandles: array of OPCHANDLE; count: integer);
var
  i, idx: longword;
  obj: pointer;
  item: TOPCItem;
  fileTime: TFileTime;
  ppErrors: PResultList;
  ppQualities: PWORDARRAY;
  ppClientHandles: PDWORDARRAY;
  ppValues: POleVariantArray;
  ppTimes: PTFileTimeARRAY;
  masterResult, masterQuality: HRESULT;
begin
  if not Succeeded(group.clientSink.QueryInterface(IOPCDataCallback, obj)) then exit;
  ppClientHandles := nil;
  ppValues := nil;
  ppErrors := nil;
  ppQualities := nil;
  ppTimes := nil;
  try
    ppClientHandles := taskMemAlloc(count, tmDWORD);
    ppValues := taskMemAlloc(count, tmOleVariant);
    ppErrors := taskMemAlloc(count, tmHResult);
    ppQualities := taskMemAlloc(count, tmWord);
    ppTimes := taskMemAlloc(count, tmFileTime);

    if (ppClientHandles = nil) or (ppValues = nil) or (ppErrors = nil) or
      (ppQualities = nil) or (ppTimes = nil) then exit;

    fileTime := DataTimeToOPCTime(now);
    masterResult := S_OK;
    masterQuality := OPC_QUALITY_GOOD;
    idx := 0;
    for i := 0 to count - 1 do begin
      item := group.findItemByServerHandle(serverHandles[i]);
      if item <> nil then begin
        ppTimes[idx] := fileTime;
        item.callbackRead(ppClientHandles[idx], ppValues[idx], ppQualities[idx]);
        if ppQualities[idx] <> OPC_QUALITY_GOOD then masterQuality := OPC_QUALITY_BAD;
        ppErrors[idx] := S_OK;
        inc(idx);
      end;
    end;

    IOPCDataCallback(obj).OnReadComplete(clientTransId, group.clientHandle,
      masterQuality, masterResult, idx, @ppClientHandles^, ppValues,
      @ppQualities^, @ppTimes^, ppErrors);
  finally
    if ppClientHandles <> nil then  CoTaskMemFree(ppClientHandles);
    if ppValues <> nil then CoTaskMemFree(ppValues);
    if ppErrors <> nil then CoTaskMemFree(ppErrors);
    if ppQualities <> nil then CoTaskMemFree(ppQualities);
    if ppTimes <> nil then CoTaskMemFree(ppTimes);
  end;
end;

procedure asyncWrite(group: TOPCGroup; clientTransId: longword;
  serverHandles: array of OPCHANDLE; values: array of variant; count: integer);
var
  obj: pointer;
  item: TOPCItem;
  ppErrors: PResultList;
  i, masterResult: longword;
  ppClientHandles: PDWORDARRAY;
  idx: longword;
begin
  if not Succeeded(group.clientSink.QueryInterface(IOPCDataCallback, obj)) then exit;
  ppClientHandles := nil;
  ppErrors := nil;
  try
    ppClientHandles := taskMemAlloc(count, tmDWord);
    ppErrors := taskMemAlloc(count, tmHResult);
    if (ppClientHandles = nil) or (ppErrors = nil) then exit;

    masterResult := S_OK;
    idx := 0;
    for i := 0 to count - 1 do begin
      item := group.findItemByServerHandle(serverHandles[i]);
      if item <> nil then begin
        ppClientHandles[idx] := item.clientHandle;
        if not item.writeable then begin
          ppErrors[idx] := OPC_E_BADRIGHTS;
          masterResult := S_FALSE;
          continue;
        end;
        item.WriteItemValue(values[i]);
        ppErrors[idx] := S_OK;
        inc(idx);
      end;
    end;

    IOPCDataCallback(obj).OnWriteComplete(clientTransId, group.clientHandle,
      masterResult, idx, @ppClientHandles^, ppErrors);
  finally
    if ppClientHandles <> nil then  CoTaskMemFree(ppClientHandles);
    if ppErrors <> nil then CoTaskMemFree(ppErrors);
  end;
end;

procedure asyncRefresh(group: TOPCGroup; clientTransId: longword);
var
  count: integer;
  obj: pointer;
  fileTime: TFileTime;
  ppErrors: PResultList;
  i, masterResult: longword;
  ppQualities: PWORDARRAY;
  ppClientHandles: PDWORDARRAY;
  ppValues: POleVariantArray;
  ppTimes: PTFileTimeARRAY;
  list: TList;
  item: TOPCItem;
begin
  if not Succeeded(group.clientSink.QueryInterface(IOPCDataCallback, obj)) then exit;
  ppClientHandles := nil;
  ppValues := nil;
  ppErrors := nil;
  ppQualities := nil;
  ppTimes := nil;

  try
    list := group.getActiveItems;
    count := list.count;
    if count = 0 then exit;

    ppClientHandles := taskMemAlloc(count, tmDWord);
    ppValues := taskMemAlloc(count, tmOleVariant);
    ppErrors := taskMemAlloc(count, tmHResult);
    ppQualities := taskMemAlloc(count, tmWord);
    ppTimes := taskMemAlloc(count, tmFileTime);

    if (ppClientHandles = nil) or (ppValues = nil) or (ppErrors = nil) or
      (ppQualities = nil) or (ppTimes = nil) then exit;

    fileTime := DataTimeToOPCTime(now);
    masterResult := S_OK;

    for i := 0 to count - 1 do begin
      item := TOPCItem(list[i]);
      ppTimes[i] := fileTime;
      ppClientHandles[i] := item.clientHandle;
      ppValues[i] := item.getCurrentValue;
      ppQualities[i] := item.quality;
      ppErrors[i] := S_OK;
    end;

    IOPCDataCallback(obj).OnDataChange(clientTransId, group.clientHandle,
      OPC_QUALITY_GOOD, masterResult, count, @ppClientHandles^, ppValues,
      @ppQualities^, @ppTimes^, ppErrors);
  finally
    if ppClientHandles <> nil then CoTaskMemFree(ppClientHandles);
    if ppValues <> nil then CoTaskMemFree(ppValues);
    if ppErrors <> nil then CoTaskMemFree(ppErrors);
    if ppQualities <> nil then CoTaskMemFree(ppQualities);
    if ppTimes <> nil then CoTaskMemFree(ppTimes);
  end;
end;

procedure asyncOnChange(group: TOPCGroup; clientTransId: longword;
  serverHandles: array of OPCHANDLE; count: integer);
var
  obj: pointer;
  fileTime: TFileTime;
  ppErrors: PResultList;
  i, idx, masterResult: longword;
  ppQualities: PWORDARRAY;
  ppClientHandles: PDWORDARRAY;
  ppValues: POleVariantArray;
  ppTimes: PFileTimeARRAY;
  item: TOPCItem;
begin
  if not Succeeded(group.clientSink.QueryInterface(IOPCDataCallback, obj)) then exit;
  ppClientHandles := nil;
  ppValues := nil;
  ppErrors := nil;
  ppQualities := nil;
  ppTimes := nil;
  idx := 0;
  try
    ppClientHandles := taskMemAlloc(count, tmDWord);
    ppValues := taskMemAlloc(count, tmOleVariant);
    ppErrors := taskMemAlloc(count, tmHResult);
    ppQualities := taskMemAlloc(count, tmWord);
    ppTimes := taskMemAlloc(count, tmFileTime);

    fileTime := DataTimeToOPCTime(now);
    masterResult := S_OK;

    for i := 0 to count - 1 do begin
      item := group.findItemByServerHandle(serverHandles[i]);
      if item <> nil then begin
        ppTimes[idx] := fileTime;
        ppClientHandles[idx] := item.clientHandle;
        item.callBackRead(ppClientHandles[idx], ppValues[idx], ppQualities[idx]);
        ppErrors[idx] := S_OK;
        inc(idx);
      end;
    end;

    IOPCDataCallback(obj).OnDataChange(clientTransId, group.clientHandle,
      OPC_QUALITY_GOOD, masterResult, idx, @ppClientHandles^, ppValues,
      @ppQualities^, ppTimes, ppErrors);
  finally
    if ppClientHandles <> nil then CoTaskMemFree(ppClientHandles);
    if ppValues <> nil then begin
      for i := 0 to idx - 1 do VarClear(ppValues[i]);
      CoTaskMemFree(ppValues);
    end;
    if ppErrors <> nil then CoTaskMemFree(ppErrors);
    if ppQualities <> nil then CoTaskMemFree(ppQualities);
    if ppTimes <> nil then CoTaskMemFree(ppTimes);
 end;
end;



end.

