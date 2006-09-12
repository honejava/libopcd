program opc_test;

{%ToDo 'opc_test.todo'}

uses
  windows, forms, sysutils, classes, variants,
  comobj, opc_da,
  activex,
  comcat in '../lib/comcat.pas',
  form_main in 'form_main.pas' {Form1},
  opc_async in '../lib/opc_async.pas',
  opc_enum in '../lib/opc_enum.pas',
  opc_error_strings in '../lib/opc_error_strings.pas',
  opc_group in '../lib/opc_group.pas',
  opc_item in '../lib/opc_item.pas',
  opc_item_props in '../lib/opc_item_props.pas',
  opc_item_proxy in '../lib/opc_item_proxy.pas',
  opc_register in '../lib/opc_register.pas',
  opc_server in '../lib/opc_server.pas',
  opc_types in '../lib/opc_types.pas',
  opc_utils in '../lib/opc_utils.pas',
  ROPC_TLB in 'ROPC_TLB.pas',
  opc_test_TLB in 'opc_test_TLB.pas';

{$R *.TLB}

{$R *.res}

type
  TTimeProxy = class (TOPCItemProxy)
  private
    _lastUpdate: longword;
  public
    procedure scan;

    function quality: longword; override;
    function lastUpdate: TDateTime; override;
    function value: variant; override;
    function datatype: TVarType; override;
  end;

  TTickProxy = class (TOPCItemProxy)
  private
    _lastUpdate: longword;
  public
    procedure scan;

    function quality: longword; override;
    function lastUpdate: TDateTime; override;
    function value: variant; override;
    function datatype: TVarType; override;
  end;

  TValueProxy = class (TOPCItemProxy)
  private
    _lastUpdated: TDateTime;
    _value: variant;
  public
    procedure write(const value: variant); override;

    function quality: longword; override;
    function lastUpdate: TDateTime; override;
    function value: variant; override;
    function datatype: TVarType; override;
    function writeable: boolean; override;
  end;

  TRealityOPCServer = class (TDA2, IDA2)
  public
    function createGroup(server: TDA2; name: string; active: boolean;
      requestedUpdateRate: longword; clientHandle: OPCHANDLE; LCID: longword): pointer; override;
    function findProxy(const ref: string): TOPCItemProxy; override;
    procedure fillItemRefList(list: TStringList); override;
    function checkItemRef(const ref: string): boolean; override;
    procedure scan(time: TDateTime); override;
  end;

  TRealityOPCGroup = class (TOPCGroup, IOPCGroup)
  end;

////////////////////////////////////////////////////////////////////////////////

procedure TTimeProxy.scan;
begin
  if getTickCount - _lastUpdate >= 10 then begin
    _lastUpdate := getTickCount;
    notifySubscribers;
  end;
end;

function TTimeProxy.quality: longword;
begin
  result := OPC_QUALITY_GOOD;
end;

function TTimeProxy.lastUpdate: TDateTime;
begin
  result := now;
end;

function TTimeProxy.value: variant;
begin
  result := now;
end;

function TTimeProxy.datatype: TVarType;
begin
  result := varDouble;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TTickProxy.scan;
begin
  if getTickCount - _lastUpdate >= 10 then begin
    _lastUpdate := getTickCount;
    notifySubscribers;
  end;
end;

function TTickProxy.quality: longword;
begin
  result := OPC_QUALITY_GOOD;
end;

function TTickProxy.lastUpdate: TDateTime;
begin
  result := now;
end;

function TTickProxy.value: variant;
begin
  result := getTickCount;
end;

function TTickProxy.datatype: TVarType;
begin
  result := varLongWord;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TValueProxy.write(const value: variant);
begin
  _value := value;
  _lastUpdated := now;
  notifySubscribers;
end;

function TValueProxy.quality: longword;
begin
  if _lastUpdated <> 0 then
    result := OPC_QUALITY_GOOD
  else
    result := OPC_QUALITY_UNCERTAIN;
end;

function TValueProxy.lastUpdate: TDateTime;
begin
  result := _lastUpdated;
end;

function TValueProxy.value: variant;
begin
  result := _value;
end;

function TValueProxy.datatype: TVarType;
begin
  result := VarType(_value);
end;

function TValueProxy.writeable: boolean;
begin
  result := true;
end;

////////////////////////////////////////////////////////////////////////////////

function TRealityOPCServer.createGroup(server: TDA2; name: string;
  active: boolean; requestedUpdateRate: longword; clientHandle: OPCHANDLE;
  LCID: longword): pointer;
begin
  result := TRealityOPCGroup.create(server, name, active, requestedUpdateRate,
    clientHandle, LCID);
end;

var
  timeProxy: TTimeProxy = nil;
  tickProxy: TTickProxy = nil;
  valueProxy: TValueProxy = nil;

function TRealityOPCServer.findProxy(const ref: string): TOPCItemProxy;
begin
  if ref = 'time' then begin
    if timeProxy = nil then timeProxy := TTimeProxy.create('time');
    result := timeProxy;
  end else if ref = 'tick' then begin
    if tickProxy = nil then tickProxy := TTickProxy.create('tick');
    result := tickProxy;
  end else if ref = 'V1' then begin
    if valueProxy = nil then valueProxy := TValueProxy.create('V1');
    result := valueProxy;
  end else
    result := nil;
end;

procedure TRealityOPCServer.fillItemRefList(list: TStringList);
begin
  list.add('time');
  list.add('tick');
  list.add('V1');
end;

function TRealityOPCServer.checkItemRef(const ref: string): boolean;
begin
  result := (ref = 'time') or (ref = 'tick') or (ref = 'V1');
end;

procedure TRealityOPCServer.scan(time: TDateTime);
begin
  inherited scan(time);
  if timeProxy <> nil then
    timeProxy.scan;
  if tickProxy <> nil then
    tickProxy.scan;
end;

begin
  CoInitialize(nil);
  registerOPCServer('ROPC.DA2', 'ROPC', TRealityOPCServer, CLASS_DA2,
    TRealityOPCGroup, CLASS_OPCGroup);

  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
