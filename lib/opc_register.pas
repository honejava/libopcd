unit opc_register;

interface

uses windows, registry, comobj, comcat, comserv, activex;

procedure registerOPCServer(const name, desc: string; serverClass: TAutoClass;
  serverGuid: TGUID; groupClass: TTypedComClass; groupGuid: TGUID);

implementation

uses opc_da, opc_enum;

procedure registerEnumerators;
begin
  TComObjectFactory.Create(ComServer, TS3UnknownEnumerator, IEnumUnknown,
    'TS3UnknownEnumerator', 'Unknown Enumerator',ciMultiInstance,
    tmFree);

  TComObjectFactory.Create(ComServer, TOPCItemAttEnumerator,
    IID_IEnumOPCItemAttributes, 'TOPCItemAttEnumerator',
    'Item Attribute Enumerator', ciMultiInstance, tmFree);

  TComObjectFactory.Create(ComServer, TOPCStringsEnumerator, IEnumString,
    'TOPCStringsEnumerator', 'String Enumerator', ciMultiInstance,
    tmFree);
end;

procedure registerOPCServer(const name, desc: string; serverClass: TAutoClass;
  serverGuid: TGUID; groupClass: TTypedComClass; groupGuid: TGUID);
var
  reg: TRegistry;
  classIDString: string;
  buffer: array [0..255] of wideChar;
begin
  TAutoObjectFactory.Create(ComServer, serverClass, serverGuid, ciMultiInstance, tmFree);
  TTypedComObjectFactory.Create(ComServer, groupClass, groupGuid, ciMultiInstance, tmFree);
  registerEnumerators;

  classIDString := GUIDToString(serverGuid);
  reg := TRegistry.Create;
  reg.RootKey := HKEY_CLASSES_ROOT;

  reg.OpenKey(name, true);
  reg.WriteString('', desc);
  reg.CloseKey;

  reg.OpenKey(name + '\Clsid', true);
  reg.WriteString('', classIDString);
  reg.CloseKey;

  reg.OpenKey('CLSID\' + classIDString, true);
  reg.WriteString('', desc);

  reg.OpenKey('CLSID\' + classIDString + '\ProgID', true);
  reg.WriteString('', name);

  reg.CloseKey;
  reg.Free;

  StringToWideChar(desc, buffer, sizeof(buffer));
  CreateComponentCategory(CATID_OPCDAServer20, buffer);
  RegisterCLSIDInCategory(serverGuid, CATID_OPCDAServer20);
end;

end.
