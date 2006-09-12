unit opc_item_proxy;

interface

uses windows, classes, sysutils, dialogs, activex, comobj, axctrls, syncobjs,
  opc_da, opc_types, opc_utils;

type
  TOPCItemProxySubscriber = class
    procedure valueChanged; virtual; abstract;
  end;

  TOPCItemProxy = class
  private
    _ref: string;
    _lock: TCriticalSection;
    _subscribers: TList;
  public
    constructor create(const ref: string);
    destructor destroy; override;

    procedure addSubscriber(obj: TOPCItemProxySubscriber);
    procedure delSubscriber(obj: TOPCItemProxySubscriber);
    procedure notifySubscribers;

    function quality: longword; virtual;
    function lastUpdate: TDateTime; virtual;
    function value: variant; virtual;
    function datatype: TVarType; virtual;
    function writeable: boolean; virtual;
    function accessRights: longword;
    procedure write(const value: variant); virtual;

    property ref: string read _ref;
  end;

implementation

uses variants;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

constructor TOPCItemProxy.create(const ref: string);
begin
  inherited create;
  _ref := ref;
  _lock := TCriticalSection.create;
  _subscribers := TList.create;
end;

destructor TOPCItemProxy.destroy;
begin
  _subscribers.free;
  _lock.free;
  inherited destroy;
end;

procedure TOPCItemProxy.addSubscriber(obj: TOPCItemProxySubscriber);
begin
  _lock.acquire;
  try
    _subscribers.add(obj);
  finally
    _lock.release;
  end;
end;

procedure TOPCItemProxy.delSubscriber(obj: TOPCItemProxySubscriber);
begin
  _lock.acquire;
  try
    _subscribers.remove(obj);
  finally
    _lock.release;
  end;
end;

procedure TOPCItemProxy.notifySubscribers;
var
  i: integer;
begin
  _lock.acquire;
  try
    for i := 0 to _subscribers.count - 1 do
      TOPCItemProxySubscriber(_subscribers[i]).valueChanged;
  finally
    _lock.release;
  end;
end;

function TOPCItemProxy.quality: longword;
begin
  result := 0;
end;

function TOPCItemProxy.lastUpdate: TDateTime;
begin
  result := 0;
end;

function TOPCItemProxy.value: variant;
begin
  result := null;
end;

function TOPCItemProxy.datatype: TVarType;
begin
  result := varNull;
end;

function TOPCItemProxy.writeable: boolean;
begin
  result := false;
end;

function TOPCItemProxy.accessRights: longword;
begin
  result := OPC_READABLE;
  if writeable then result := result or OPC_WRITEABLE;
end;

procedure TOPCItemProxy.write(const value: variant);
begin
end;

end.
