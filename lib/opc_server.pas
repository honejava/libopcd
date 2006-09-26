unit opc_server;

{$IFDEF VER150}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses windows, syncobjs, comobj, activex, axctrls, sysutils, dialogs, classes, stdvcl,
  opc_da, opc_common, opc_types, opc_enum, opc_item_props, opc_error, opc_utils,
  opc_error_strings, opc_item_proxy;

type
  TDA2 = class(TAutoObject, IOPCServer, IOPCCommon, IOPCServerPublicGroups,
    IOPCBrowseServerAddressSpace, IPersist, IPersistFile,
    IConnectionPointContainer,IOPCItemProperties)
  private
    _opcItemProperties: TOPCItemProp;
    _connectionPoints: TConnectionPoints;
  protected
    property iFIConnectionPoints:TConnectionPoints read _connectionPoints
      write _connectionPoints implements IConnectionPointContainer;

//IOPCServer
    function AddGroup(szName: POleStr; bActive: BOOL;
      dwRequestedUpdateRate: DWORD; hClientGroup: OPCHANDLE;
      pTimeBias: PLongint; pPercentDeadband: PSingle; dwLCID: DWORD;
      out phServerGroup: OPCHANDLE; out pRevisedUpdateRate: DWORD;
      const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;
    function GetErrorString(dwError: HResult; dwLocale: TLCID;
      out ppString: POleStr): HResult; overload; stdcall;
    function GetGroupByName(szName: POleStr; const riid: TIID;
      out ppUnk: IUnknown): HResult; stdcall;
    function GetStatus(out ppServerStatus: POPCSERVERSTATUS): HResult; stdcall;
    function RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;
    function CreateGroupEnumerator(dwScope: OPCENUMSCOPE; const riid: TIID;
      out ppUnk:IUnknown): HResult; stdcall;

//IOPCCommon
    function SetLocaleID(dwLcid: TLCID): HResult; stdcall;
    function GetLocaleID(out pdwLcid: TLCID): HResult; stdcall;
    function QueryAvailableLocaleIDs(out pdwCount: UINT;
      out pdwLcid: PLCIDARRAY): HResult; stdcall;
    function GetErrorString(dwError: HResult; out ppString: POleStr): HResult;
      overload; stdcall;
    function SetClientName(szName: POleStr): HResult; stdcall;

//IOPCServerPublicGroups
    function GetPublicGroupByName(szName: POleStr; const riid:TIID;
      out ppUnk: IUnknown): HResult; stdcall;
    function RemovePublicGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;

//IOPCBrowseServerAddressSpace
    function QueryOrganization(out pNameSpaceType: OPCNAMESPACETYPE): HResult; stdcall;
    function ChangeBrowsePosition(dwBrowseDirection: OPCBROWSEDIRECTION;
      szString: POleStr): HResult; stdcall;
    function BrowseOPCItemIDs(dwBrowseFilterType: OPCBROWSETYPE;
      szFilterCriteria: POleStr; vtDataTypeFilter: TVarType;
      dwAccessRightsFilter: DWORD; out ppIEnumString: IEnumString): HResult; stdcall;
    function GetItemID(szItemDataID: POleStr; out szItemID: POleStr): HResult; stdcall;
    function BrowseAccessPaths(szItemID: POleStr;
      out ppIEnumString: IEnumString): HResult;stdcall;

//IPersistFile
    function GetClassID(out classID: TCLSID): HResult; stdcall;
    function IsDirty: HResult; stdcall;
    function Load(pszFileName: POleStr; dwMode: longint): HResult; stdcall;
    function Save(pszFileName: POleStr; fRemember: BOOL): HResult; stdcall;
    function SaveCompleted(pszFileName: POleStr): HResult; stdcall;
    function GetCurFile(out pszFileName: POleStr): HResult; stdcall;
//IPersistFile end
  protected
    _lock: TCriticalSection;
    _groups: TList;
    _localID: longword;
    _clientName: string;
    _startTime, _lastUpdateTime: TDateTime;
    _onSDConnect: TConnectEvent;
    _clientSink: IUnknown;

    _lastGroupServerHandle, _lastItemServerHandle: DWORD;

  public
    property iFIOPCItemProperties:TOPCItemProp read _opcItemProperties
      write _opcItemProperties implements IOPCItemProperties;

    procedure initialize; override;
    procedure shutdownOnConnect(const sink: IUnknown; connecting: boolean);

    constructor create;
    destructor destroy;override;
    function makeGroupServerHandle: longword;
    function makeItemServerHandle: longword;

    function findGroupByServerHandle(serverHandle: DWORD): pointer;
    function findGroupByName(const name: string): pointer;
    procedure addGroupRef(group: TObject);
    procedure removeGroupRef(group: TObject);
    function fillGroupNameList(list: TStringList;publicFlag: boolean): TStringList;
    function fillGroupInterfaceList(list: TList;publicFlag: boolean): TList;
    procedure scan; virtual;
    function createGroup(server: TDA2; name: string; active: boolean;
      requestedUpdateRate: DWORD; clientHandle: OPCHANDLE; LCID: DWORD): pointer; virtual; abstract;

    function CreateGroupNameEnumerator(filter, publicFlag: boolean): IUnknown;
    function CreateGroupInterfaceEnumerator(filter, publicFlag: boolean): IUnknown;

    function lastUpdateTime: TDateTime;
    procedure touch;
    procedure getServerInfo(var ppServerStatus: POPCSERVERSTATUS); virtual;
    function findProxy(const ref: string): TOPCItemProxy; virtual; abstract;
    procedure fillItemRefList(list: TStringList); virtual; abstract;
    function checkItemRef(var ref: string): boolean; virtual; abstract;
  end;

procedure scanOPCServers;
procedure KillOPCServers;

implementation

uses comserv, opc_group;

var
  _servers: TList;
  _serversLock: TCriticalSection;

procedure addServer(server: TDA2);
begin
  _serversLock.acquire;
  try
    _servers.add(server);
  finally
    _serversLock.release;
  end;
end;

procedure removeServer(server: TDA2);
begin
  _serversLock.acquire;
  try
    _servers.remove(server);
  finally
    _serversLock.release;
  end;
end;

procedure scanOPCServers;
var
  i: integer;
begin
  _serversLock.acquire;
  try
    for i := 0 to _servers.count - 1 do
      TDA2(_servers[i]).scan;
  finally
    _serversLock.release;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

function TDA2.AddGroup(szName: POleStr; bActive: BOOL;
  dwRequestedUpdateRate: DWORD; hClientGroup: OPCHANDLE; pTimeBias: PLongint;
  pPercentDeadband: PSingle; dwLCID: DWORD; out phServerGroup: OPCHANDLE;
  out pRevisedUpdateRate: DWORD; const riid: TIID; out ppUnk: IUnknown): HResult;
var
  newName: string;
  i: longint;
  group: TOPCGroup;
begin
  newName := szName;
  i:=0;

  //generate a unique name
  while findGroupByName(newName) <> nil do begin
    inc(i); newName := szName + inttostr(i);
  end;

  _lock.acquire;
  try
    group := self.createGroup(self, newName, bActive, dwRequestedUpdateRate,
      hClientGroup, dwLCID);
    if group = nil then begin
     result := E_OUTOFMEMORY;
     exit;
    end;
    pRevisedUpdateRate := group.requestedUpdateRate;
    ppUnk := group;
    result := S_OK;
  finally _lock.release; end;
end;

function TDA2.GetErrorString(dwError: HResult; dwLocale: TLCID;
  out ppString: POleStr): HResult;
begin
 ppString := StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result := S_OK;
end;

function TDA2.GetGroupByName(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult;
var
  group: TOPCGroup;
begin
  _lock.acquire;
  try
    group := findGroupByName(szName);
    if (addr(ppUnk) = nil) or (group = nil) then begin
      result := E_INVALIDARG;
      exit;
    end;
    result := IUnknown(group).QueryInterface(riid, ppUnk);
  finally _lock.release; end;
end;

function TDA2.GetStatus(out ppServerStatus: POPCSERVERSTATUS): HResult;
begin
  if (addr(ppServerStatus) = nil) then begin
    result:=E_INVALIDARG;
    exit;
  end;
  ppServerStatus := taskMemAlloc(1, tmServerStatus);
  if ppServerStatus = nil then begin
    result:=E_OUTOFMEMORY;
    exit;
  end;

  _lock.acquire;
  try
    ppServerStatus.ftStartTime := DataTimeToOPCTime(_startTime);
    ppServerStatus.ftCurrentTime := DataTimeToOPCTime(now);
    ppServerStatus.ftLastUpdateTime := DataTimeToOPCTime(_lastUpdateTime);
    ppServerStatus.dwServerState := OPC_STATUS_RUNNING;
    ppServerStatus.dwGroupCount := _groups.count;
    ppServerStatus.dwBandWidth := 100;
    getServerInfo(ppServerStatus);
    result := S_OK;
  finally _lock.release; end;
end;

function TDA2.RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult;
var
  group: TOPCGroup;
begin
  if hServerGroup < 1 then begin
    result := E_INVALIDARG;
    exit;
  end;
  _lock.acquire;
  try
    group := findGroupByServerHandle(hServerGroup);
    if group = nil then begin
      result := E_INVALIDARG;
      exit;
    end;
    if (group.refCount > 2) and not bForce then begin
     result := OPC_S_INUSE;
     exit;
    end;
    group.Free;
    result := S_OK;
  finally _lock.release; end;
end;

function TDA2.CreateGroupEnumerator(dwScope: OPCENUMSCOPE; const riid: TIID;
  out ppUnk: IUnknown): HResult;
begin
  if not (IsEqualIID(riid, IEnumUnknown) or IsEqualIID(riid, IEnumString)) then
  begin
    result := E_NOINTERFACE;
    exit;
  end;

  result := S_OK;
  _lock.acquire;
  try
    if IsEqualIID(riid, IEnumString) then
      case dwScope of
      OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
        ppUnk := createGroupNameEnumerator(true, false);
      OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
        ppUnk := createGroupNameEnumerator(true, true);
      OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
        ppUnk := createGroupNameEnumerator(false, false);
      else
        result := E_INVALIDARG;
      end
    else
      case dwScope of
      OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
        ppUnk := createGroupInterfaceEnumerator(true, false);
      OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
        ppUnk := createGroupInterfaceEnumerator(true, true);
      OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
        ppUnk := createGroupInterfaceEnumerator(false, false);
      else
        result := E_INVALIDARG;
      end;
  finally _lock.release; end;
end;

function TDA2.SetLocaleID(dwLcid: TLCID): HResult;
begin
  if (dwLcid = LOCALE_SYSTEM_DEFAULT) or (dwLcid = LOCALE_USER_DEFAULT) then begin
    _localID := dwLcid;
    result := S_OK;
  end else
    result := E_INVALIDARG;
end;

function TDA2.GetLocaleID(out pdwLcid: TLCID): HResult;
begin
  pdwLcid := _localID;
  result := S_OK;
end;

function TDA2.QueryAvailableLocaleIDs(out pdwCount: UINT;
  out pdwLcid: PLCIDARRAY): HResult;
begin
  pdwCount := 2;
  pdwLcid := taskMemAlloc(pdwCount, tmLCID);
  if (pdwLcid = nil) then begin
    result := E_OUTOFMEMORY;
    exit;
  end;
  pdwLcid[0] := LOCALE_SYSTEM_DEFAULT;
  pdwLcid[1] := LOCALE_USER_DEFAULT;
  result := S_OK;
end;

function TDA2.GetErrorString(dwError: HResult; out ppString: POleStr): HResult;
begin
  ppString := StringToLPOLESTR(OPCErrorCodeToString(dwError));
  result := S_OK;
end;

function TDA2.SetClientName(szName: POleStr): HResult;
begin
  if (addr(szName) = nil) then begin
    result := E_INVALIDARG;
    exit;
  end;
  _clientName := szName;
  result := S_OK;
end;

function TDA2.GetPublicGroupByName(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult;
begin
  result := GetGroupByName(szName, riid, ppUnk);
end;

function TDA2.RemovePublicGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult;
begin
  result := RemoveGroup(hServerGroup, bForce);
end;

function TDA2.QueryOrganization(out pNameSpaceType: OPCNAMESPACETYPE):HResult;
begin
  pNameSpaceType := OPC_NS_FLAT;
  result := S_OK;
end;

function TDA2.ChangeBrowsePosition(dwBrowseDirection: OPCBROWSEDIRECTION;
  szString: POleStr): HResult;
begin
  result := E_FAIL;
end;

function TDA2.BrowseOPCItemIDs(dwBrowseFilterType: OPCBROWSETYPE;
  szFilterCriteria: POleStr; vtDataTypeFilter: TVarType;
  dwAccessRightsFilter: DWORD; out ppIEnumString: IEnumString): HResult;
var
  list: TStringList;
begin
  list := nil;
  _lock.acquire;
  try
    list := TStringList.create;
    if list = nil then begin
      result := E_OUTOFMEMORY;
      exit;
    end;
    fillItemRefList(list);
    ppIEnumString := TOPCStringsEnumerator.Create(list);
    result := S_OK;
  finally
    if list <> nil then list.free;
    _lock.release;
  end;
end;

function TDA2.GetItemID(szItemDataID: POleStr; out szItemID: POleStr): HResult;
var
  ref: string;
begin
  _lock.acquire;
  try
    ref := szItemDataID;
    if checkItemRef(ref) then begin
      szItemId := StringToLPOLESTR(ref);
      result := S_OK;
    end else
      result := OPC_E_UNKNOWNITEMID;
  finally _lock.release; end;
end;

function TDA2.BrowseAccessPaths(szItemID: POleStr;
  out ppIEnumString: IEnumString): HResult;
begin
  result :=E_NOTIMPL;
end;

function TDA2.GetClassID(out classID: TCLSID): HResult;
begin
  result := S_FALSE;
end;

function TDA2.IsDirty: HResult;
begin
  result := S_FALSE;
end;

function TDA2.Load(pszFileName: POleStr; dwMode: Longint): HResult;
begin
  result := S_FALSE;
end;

function TDA2.Save(pszFileName: POleStr; fRemember: BOOL): HResult;
begin
  result := S_FALSE;
end;

function TDA2.SaveCompleted(pszFileName: POleStr): HResult;
begin
  result := S_FALSE;
end;

function TDA2.GetCurFile(out pszFileName: POleStr): HResult;
begin
  result := S_FALSE;
end;

procedure TDA2.initialize;
begin
  inherited Initialize;
  _lock := TCriticalSection.create;
  _startTime := now;
  _lastUpdateTime := 0;
  _localID := LOCALE_SYSTEM_DEFAULT;

  _connectionPoints := TConnectionPoints.Create(self);
  _opcItemProperties := TOPCItemProp.Create(self);

  _onSDConnect := ShutdownOnConnect;
  _connectionPoints.CreateConnectionPoint(IID_IOPCShutdown, ckSingle, _onSDConnect);

  _groups := TList.Create;

  addServer(self);
end;

procedure TDA2.shutdownOnConnect(const sink: IUnknown; connecting: boolean);
begin
  if connecting then
    _clientSink := sink
  else
    _clientSink := nil
end;

constructor TDA2.create;
begin
  inherited create;
  messagebeep($FFFF);
end;

destructor TDA2.destroy;
var
  i: integer;
begin
  removeServer(self);

  for i:= _groups.count - 1 downto 0 do
    TOPCGroup(_groups[i]).free;
  _groups.free;

  if Assigned(_connectionPoints) then _connectionPoints.Free;
  if Assigned(_opcItemProperties) then _opcItemProperties.Free;

  _lock.free;
  inherited destroy;
end;

function TDA2.makeGroupServerHandle: longword;
begin
  _lock.acquire;
  try
    inc(_lastGroupServerHandle);
    result := _lastGroupServerHandle;
  finally _lock.release; end;
end;

function TDA2.makeItemServerHandle: longword;
begin
  _lock.acquire;
  try
    inc(_lastItemServerHandle);
    result := _lastItemServerHandle;
  finally _lock.release; end;
end;

function TDA2.findGroupByServerHandle(serverHandle: DWORD): pointer;
var
  i: integer;
begin
  _lock.acquire;
  try
    for i := 0 to _groups.count - 1 do
      if TOPCGroup(_groups[i]).serverHandle = serverHandle then begin
        result := _groups[i];
        exit;
      end;
    result := nil;
  finally _lock.release; end;
end;

function TDA2.findGroupByName(const name: string): pointer;
var
  i: integer;
begin
  _lock.acquire;
  try
    for i := 0 to _groups.count - 1 do
      if TOPCGroup(_groups[i]).name = name then begin
        result := _groups[i];
        exit;
      end;
    result := nil;
  finally _lock.release; end;
end;

procedure TDA2.addGroupRef(group: TObject);
begin
  _lock.acquire;
  try
    _groups.add(group);
  finally _lock.release; end;
end;

procedure TDA2.removeGroupRef(group: TObject);
begin
  _lock.acquire;
  try
    _groups.remove(group);
  finally _lock.release; end;
end;

function TDA2.fillGroupNameList(list: TStringList; publicFlag: boolean): TStringList;
var
  i: integer;
begin
  _lock.acquire;
  try
    if list = nil then list := TStringList.create;
    result := list;
    for i := 0 to _groups.count - 1 do
      if TOPCGroup(_groups[i]).isPublic = publicFlag then
        result.Add(TOPCGroup(_groups[i]).name);
  finally _lock.release; end;
end;

function TDA2.fillGroupInterfaceList(list: TList;publicFlag: boolean): TList;
var
  i: integer;
  obj: pointer;
begin
  _lock.acquire;
  try
    if list = nil then list := TList.create;
    result := list;
    for i := 0 to _groups.count - 1 do
      if TOPCGroup(_groups[i]).isPublic = publicFlag then begin
        obj := nil;
        IUnknown(TOPCGroup(_groups[i])).QueryInterface(IUnknown, obj);
        if Assigned(obj) then list.Add(obj);
      end;
  finally _lock.release; end;
end;

procedure TDA2.scan;
var
  i: integer;
begin
  _lock.acquire;
  try
    _lastUpdateTime := now;
    for i := 0 to _groups.count - 1 do
      TOPCGroup(_groups[i]).scan;
  finally _lock.release; end;
end;

function TDA2.CreateGroupNameEnumerator(filter, publicFlag: boolean): IUnknown;
var
  list: TStringList;
begin
  list := nil;
  try
    if filter then
      list := fillGroupNameList(nil, publicFlag)
    else begin
      list := fillGroupNameList(nil, false);
      fillGroupNameList(list, true);
    end;
    result := TOPCStringsEnumerator.Create(list);
  finally if list <> nil then list.free; end;
end;

function TDA2.CreateGroupInterfaceEnumerator(filter, publicFlag: boolean): IUnknown;
var
  list: TList;
begin
  list := nil;
  try
    if filter then
      list := fillGroupInterfaceList(nil, publicFlag)
    else begin
      list := fillGroupInterfaceList(nil, false);
      fillGroupInterfaceList(list, true);
    end;
    result := TS3UnknownEnumerator.Create(list);
  finally if list <> nil then list.free; end;
end;

////////////////////////////////////////////////////////////////////////////////

function TDA2.lastUpdateTime: TDateTime;
begin
  result := _lastUpdateTime;
end;

procedure TDA2.touch;
begin
  _lastUpdateTime := now;
end;

procedure TDA2.getServerInfo(var ppServerStatus: POPCSERVERSTATUS);
begin
  ppServerStatus.wMajorVersion := 1;
  ppServerStatus.wMinorVersion := 2;
  ppServerStatus.wBuildNumber := 5;
  ppServerStatus.szVendorInfo := StringToLPOLESTR('Vendor Info');
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

procedure KillOPCServers;
var
  i: integer;
  server: TDA2;
begin
  for i := _servers.count - 1 downto 0 do
  try
    server := _servers[i];
    CoDisconnectObject(server as IUnknown,0);
    server.Free;
  except
  end;
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  _servers := TList.create;
  _serversLock := TCriticalSection.create;
  ComServer.UIInteractive:=false;
finalization
  try KillOPCServers; except end;
  CoUninitialize;
end.
