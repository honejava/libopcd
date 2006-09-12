unit opc_item;

interface

uses windows, classes, sysutils, dialogs, activex, comobj, axctrls,
  opc_da, opc_server, opc_types, opc_utils, opc_item_proxy;

type
  TOPCItem = class; //forward declaration

  TOPCItemAttributes = class
  public
    _active: longbool;
    _euType: integer;
    _euInfo: OleVariant;
    _accessPath, _itemID: string;
    _requestedDataType, _canonicalDataType: word;
    _clientHandle, _serverHandle, _accessRights: longword;

    constructor create(item: TOPCItem);
  end;

  TOPCItem = class (TOPCItemProxySubscriber)
  private
    _server: TDA2;
    _group: pointer;
    _requestedDataType: TVarType;
    _ref: string;
    _active: boolean;
    _serverHandle, _clientHandle: OPCHANDLE;
    _proxy: TOPCItemProxy;
  public
    constructor create(server: TDA2; group: pointer; proxy: TOPCItemProxy;
      const ref: string; clientHandle: OPCHANDLE; requestedDataType: TVarType;
      active: boolean);
    constructor clone(source: TOPCItem);
    destructor destroy;override;

    procedure valueChanged; override;

    procedure setActive(state: longbool);
    procedure setClientHandle(handle: OPCHANDLE);
    function getCurrentValue: variant;
    procedure readItemValueStateTime(source: word; var stateRec: OPCITEMSTATE);
    procedure writeItemValue(const value: variant);
    procedure setRequestedDataType(datatype: TVarType);
    procedure callBackRead(var handle: OPCHANDLE; var value: OleVariant;
      var quality: word);

    function active: boolean;
    function writeable: boolean;
    function accessRights: longword;
    function quality: word;
    function serverHandle: OPCHANDLE;
    function clientHandle: OPCHANDLE;
    function canonicalDataType: TVarType;
  end;

implementation

uses opc_group;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

constructor TOPCItemAttributes.create(item: TOPCItem);
begin
  inherited create;
  _active := item.active;
  _euType := OPC_NOENUM;
  _euInfo := VT_EMPTY;
  _accessPath := '';
  _itemID := item._ref;
  _requestedDataType := item._requestedDataType;
  _canonicalDataType := item.canonicalDataType;
  _clientHandle := item._clientHandle;
  _serverHandle := item._serverHandle;
  _accessRights := item.accessRights;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

constructor TOPCItem.create(server: TDA2; group: pointer; proxy: TOPCItemProxy;
  const ref: string; clientHandle: OPCHANDLE; requestedDataType: TVarType;
  active: boolean);
begin
  inherited create;
  _server := server;
  _group := group;
  _proxy := proxy;
  _requestedDataType := requestedDataType;
  _ref := ref;
  _active := active;
  _serverHandle := _server.makeItemServerHandle;
  _clientHandle := clientHandle;

  _proxy.addSubscriber(self);
end;

constructor TOPCItem.clone(source: TOPCItem);
begin
  inherited create;
  _server := source._server;
  _proxy := source._proxy;
  _requestedDataType := source._requestedDataType;
  _ref := source._ref;
  _active := source._active;
  _serverHandle := _server.makeItemServerHandle;
  _clientHandle := source._clientHandle;

  _proxy.addSubscriber(self);
end;

destructor TOPCItem.destroy;
begin
  _proxy.delSubscriber(self);
  inherited destroy;
end;

procedure TOPCItem.valueChanged;
begin
  TOPCGroup(_group).addUpdatedItem(self);
end;

procedure TOPCItem.setActive(state: longbool);
begin
  _active := state;
end;

procedure TOPCItem.setClientHandle(handle: OPCHANDLE);
begin
  _clientHandle := handle;
end;

function TOPCItem.getCurrentValue: variant;
var
  value: variant;
begin
  value := _proxy.value;
  if (_requestedDataType <> _proxy.dataType) and (_requestedDataType <> 0) then
    result := convertVariant(value, _requestedDataType)
 else
   result := value;
end;

procedure TOPCItem.readItemValueStateTime(source: word;
  var stateRec:OPCITEMSTATE);
begin
  stateRec.wQuality := _proxy.quality;
  stateRec.hClient := _clientHandle;
  stateRec.vDataValue := _proxy.value;
  stateRec.ftTimeStamp := DataTimeToOPCTime(_proxy.lastUpdate);
end;

procedure TOPCItem.writeItemValue(const value: variant);
begin
  _proxy.write(value);
end;

procedure TOPCItem.setRequestedDataType(datatype: TVarType);
begin
  if datatype = VT_EMPTY then
    _requestedDataType := _proxy.dataType
  else
    _requestedDataType := datatype;
end;

procedure TOPCItem.callBackRead(var handle: OPCHANDLE; var value:OleVariant;
  var quality: word);
begin
  handle := _clientHandle;
  quality := _proxy.quality;
  value := getCurrentValue;
end;

////////////////////////////////////////////////////////////////////////////////

function TOPCItem.active: boolean;
begin
  result := _active;
end;

function TOPCItem.writeable: boolean;
begin
  result := _proxy.writeable;
end;

function TOPCItem.accessRights: longword;
begin
  result := _proxy.accessRights;
end;

function TOPCItem.quality: word;
begin
  if not _active then
    result := OPC_QUALITY_OUT_OF_SERVICE
  else
    result := _proxy.quality;
end;

function TOPCItem.serverHandle: OPCHANDLE;
begin
  result := _serverHandle;
end;

function TOPCItem.clientHandle: OPCHANDLE;
begin
  result := _clientHandle;
end;

function TOPCItem.canonicalDataType: TVarType;
begin
  result := _proxy.dataType;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

end.
