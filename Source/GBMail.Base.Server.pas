unit GBMail.Base.Server;

interface

uses
  GBMail.Interfaces,
  System.Classes,
  System.SysUtils;

type
  TGBMailServerBase = class (TInterfacedObject, IGBMailServer)
  private
    [Weak]
    FParent: IGBMail;
    FHost: string;
    FPort: Integer;
    FUsername: string;
    FPassword: string;
    FUseSSL: Boolean;
    FUseTLS: Boolean;
    FRequireAuthentication: Boolean;
    FConnectTimeOut: Integer;
    FReadTimeOut: Integer;
  public
    constructor Create(AParent: IGBMail);
    class function New(AParent: IGBMail): IGBMailServer;

    function &Begin: IGBMailServer;
    function &End: IGBMail;

    function Host(AValue: string): IGBMailServer; overload;
    function Port(AValue: Integer): IGBMailServer; overload;
    function Username(AValue: string): IGBMailServer; overload;
    function Password(AValue: string): IGBMailServer; overload;
    function UseSSL(AValue: Boolean): IGBMailServer; overload;
    function UseTLS(AValue: Boolean): IGBMailServer; overload;
    function ConnectTimeOut(AValue: Integer): IGBMailServer; overload;
    function ReadTimeOut(AValue: Integer): IGBMailServer; overload;
    function RequireAuthentication(AValue: Boolean): IGBMailServer; overload;

    function Host: string; overload;
    function Port: Integer; overload;
    function Username: string; overload;
    function Password: string; overload;
    function UseSSL: Boolean; overload;
    function UseTLS: Boolean; overload;
    function ConnectTimeOut: Integer; overload;
    function ReadTimeOut: Integer; overload;
    function RequireAuthentication: Boolean; overload;
  end;

implementation

const
  BASE_CONNECT_TIME_OUT = 10000;
  BASE_READ_TIME_OUT = 10000;

{ TGBMailServerBase }

function TGBMailServerBase.&Begin: IGBMailServer;
begin
  Result := Self;
end;

function TGBMailServerBase.&End: IGBMail;
begin
  Result := FParent;
end;

function TGBMailServerBase.ConnectTimeOut: Integer;
begin
  Result := FConnectTimeOut;
end;

function TGBMailServerBase.ConnectTimeOut(AValue: Integer): IGBMailServer;
begin
  Result := Self;
  FConnectTimeOut := AValue;
end;

constructor TGBMailServerBase.Create(AParent: IGBMail);
begin
  Self.FParent := AParent;
  Self.FUseSSL := True;
  Self.FUseTLS := True;
  Self.FRequireAuthentication := True;
  Self.FConnectTimeOut := BASE_CONNECT_TIME_OUT;
  Self.FReadTimeOut := BASE_READ_TIME_OUT;
end;

function TGBMailServerBase.Host(AValue: string): IGBMailServer;
begin
  Result := Self;
  FHost  := AValue;
end;

function TGBMailServerBase.Host: string;
begin
  Result := FHost;
end;

class function TGBMailServerBase.New(AParent: IGBMail): IGBMailServer;
begin
  Result := Self.Create(AParent);
end;

function TGBMailServerBase.Password(AValue: string): IGBMailServer;
begin
  Result := Self;
  FPassword := AValue;
end;

function TGBMailServerBase.Password: string;
begin
  Result := FPassword;
end;

function TGBMailServerBase.Port: Integer;
begin
  Result := FPort;
end;

function TGBMailServerBase.Port(AValue: Integer): IGBMailServer;
begin
  Result := Self;
  FPort := AValue;
end;

function TGBMailServerBase.ReadTimeOut: Integer;
begin
  Result := FReadTimeOut;
end;

function TGBMailServerBase.ReadTimeOut(AValue: Integer): IGBMailServer;
begin
  Result := Self;
  FReadTimeOut := AValue;
end;

function TGBMailServerBase.RequireAuthentication: Boolean;
begin
  Result := FRequireAuthentication;
end;

function TGBMailServerBase.RequireAuthentication(AValue: Boolean): IGBMailServer;
begin
  Result := Self;
  FRequireAuthentication := AValue;
end;

function TGBMailServerBase.Username(AValue: string): IGBMailServer;
begin
  Result := Self;
  FUsername := AValue;
end;

function TGBMailServerBase.Username: string;
begin
  Result := FUsername;
end;

function TGBMailServerBase.UseSSL(AValue: Boolean): IGBMailServer;
begin
  Result := Self;
  FUseSSL := AValue;
end;

function TGBMailServerBase.UseSSL: Boolean;
begin
  Result := FUseSSL;
end;

function TGBMailServerBase.UseTLS(AValue: Boolean): IGBMailServer;
begin
  Result := Self;
  FUseTLS := AValue;
end;

function TGBMailServerBase.UseTLS: Boolean;
begin
  Result := FUseTLS;
end;

end.
