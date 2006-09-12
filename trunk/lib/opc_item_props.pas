unit opc_item_props;

{$IFDEF VER150}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses windows, sysutils, classes, activex, comobj, dialogs, axctrls,
  opc_da, opc_error, opc_types;

type
  TOPCItemProp = class
  private
    _server: pointer;
  public
    constructor create(server: pointer);

    function QueryAvailableProperties(szItemID: POleStr; out pdwCount: DWORD;
      out ppPropertyIDs: PDWORDARRAY; out ppDescriptions: POleStrList;
      out ppvtDataTypes: PVarTypeList): HResult; stdcall;
    function GetItemProperties(szItemID: POleStr; dwCount: DWORD;
      pdwPropertyIDs: PDWORDARRAY; out ppvData: POleVariantArray;
      out ppErrors: PResultList): HResult; stdcall;
    function LookupItemIDs(szItemID: POleStr; dwCount: DWORD;
      pdwPropertyIDs: PDWORDARRAY; out ppszNewItemIDs: POleStrList;
      out ppErrors: PResultList): HResult; stdcall;
  end;

implementation

uses opc_server, opc_item_proxy, opc_utils;

constructor TOPCItemProp.create(server: pointer);
begin
  inherited create;
  _server := server;
end;

function TOPCItemProp.QueryAvailableProperties(szItemID: POleStr;
  out pdwCount: DWORD; out ppPropertyIDs: PDWORDARRAY;
  out ppDescriptions: POleStrList; out ppvtDataTypes: PVarTypeList): HResult;
var
  proxy: TOPCItemProxy;
begin
  proxy := TDA2(_server).findProxy(szItemID);
  if proxy = nil then begin
    result := OPC_E_INVALIDITEMID;
    exit;
  end;

  pdwCount := 1;
  ppPropertyIDs := taskMemAlloc(pdwCount, tmDWord);
  ppDescriptions := taskMemAlloc(pdwCount, tmPOleStr);
  ppvtDataTypes := taskMemAlloc(pdwCount, tmVarType);

  if (ppPropertyIDs = nil) or (ppDescriptions = nil) or (ppvtDataTypes = nil) then begin
    if ppPropertyIDs <> nil then  CoTaskMemFree(ppPropertyIDs);
    if ppDescriptions <> nil then  CoTaskMemFree(ppDescriptions);
    if ppvtDataTypes <> nil then  CoTaskMemFree(ppvtDataTypes);
    result := E_OUTOFMEMORY;
    exit;
  end;

{  ppPropertyIDs[0] := proxy.id;
  ppDescriptions[0] := szItemId;
  ppvtDataTypes[0] := proxy.canonicalDataType;}
  result:=S_OK;
end;

function TOPCItemProp.GetItemProperties(szItemID: POleStr; dwCount: DWORD;
  pdwPropertyIDs: PDWORDARRAY; out ppvData: POleVariantArray;
  out ppErrors: PResultList): HResult;
var
  data: variant;
  i: integer;
  ppArray: PDWORDARRAY;
  proxy: TOPCItemProxy;
begin
  proxy := TDA2(_server).findProxy(szItemID);
  if proxy = nil then begin
    result := OPC_E_INVALIDITEMID;
    exit;
  end;

  ppvData := taskMemAlloc(dwCount, tmOleVariant);
  ppErrors := taskMemAlloc(dwCount, tmHResult);

  if (ppvData = nil) or (ppErrors = nil) then begin
    if ppvData <> nil then CoTaskMemFree(ppvData);
    if ppErrors <> nil then CoTaskMemFree(ppErrors);
    result := E_OUTOFMEMORY;
    exit;
  end;

  ppArray := @pdwPropertyIDs^;
  result := S_OK;
  for i := 0 to dwCount - 1 do begin
    case ppArray[i] of
    1:
      data := 0{proxy.canonicalDatatype};
    2:
      data := 0{proxy.currentValue};
    5:
      begin
        if true{proxy.writeable} then
          data := OPC_READABLE or OPC_WRITEABLE
        else
          data := OPC_READABLE;
      end
    else
      begin
        ppErrors[i] := OPC_E_INVALID_PID;
        result := S_FALSE;
        continue;
      end;
    end;
    ppvData[i] := data;
    ppErrors[i] := S_OK;
  end;
end;

function TOPCItemProp.LookupItemIDs(szItemID: POleStr; dwCount: DWORD;
  pdwPropertyIDs: PDWORDARRAY; out ppszNewItemIDs: POleStrList;
  out ppErrors: PResultList): HResult;
var
  i: integer;
  proxy: TOPCItemProxy;
begin
  proxy := TDA2(_server).findProxy(szItemID);
  if proxy = nil then begin
    result := OPC_E_INVALIDITEMID;
    exit;
  end;

  ppszNewItemIDs := taskMemAlloc(dwCount, tmPOleStr);
  ppErrors := taskMemAlloc(dwCount, tmHResult);

  if (ppszNewItemIDS = nil) or (ppErrors = nil) then begin
    if ppszNewItemIDs <> nil then CoTaskMemFree(ppszNewItemIDs);
    if ppErrors <> nil then CoTaskMemFree(ppErrors);
    result := E_OUTOFMEMORY;
    exit;
  end;

  for i := 0 to dwCount - 1 do begin
    ppszNewItemIDs[i] := StringToLPOLESTR(szItemID);
    ppErrors[i] := S_OK;
  end;

  result:=S_OK;
end;

end.
