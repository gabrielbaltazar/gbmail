unit GBMail.Interfaces;

interface

uses
  System.SysUtils,
  System.Classes;

type
  EGBMailException = class(Exception);
  IGBMailServer = interface;

  IGBMail = interface
    ['{2C44519A-0B71-449D-BC2D-6D429EA37BE3}']
    function Server: IGBMailServer;

    function From(AAddress: string; AName: string = ''): IGBMail;
    function AddRecipient(AAddress: string; AName: string = ''): IGBMail;
    function AddCcRecipient(AAddress: string; AName: string = ''): IGBMail;
    function AddBccRecipient(AAddress: string; AName: string = ''): IGBMail;
    function AddAttachment(AFileName: string): IGBMail; overload;
    function AddAttachment(AStream: TMemoryStream; AFileName: string): IGBMail; overload;
    function AddHtmlImage(AFileName: string; out AID: string): IGBMail; overload;
    function AddHtmlImage(AImage: TMemoryStream; out AID: string): IGBMail; overload;

    function UseHtml(AValue: Boolean): IGBMail;
    function Subject(AValue: string): IGBMail;
    function Message(AValue: string): IGBMail; overload;
    function Message(AValue: TStrings): IGBMail; overload;
    function Send: IGBMail;
  end;

  IGBMailServer = interface
    ['{05CBBF52-507A-4DD0-9163-F389DB4EB7D5}']
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

    function Host: string; overload;
    function Port: Integer; overload;
    function Username: string; overload;
    function Password: string; overload;
    function UseSSL: Boolean; overload;
    function UseTLS: Boolean; overload;
    function ConnectTimeOut: Integer; overload;
    function ReadTimeOut: Integer; overload;

    function RequireAuthentication(AValue: Boolean): IGBMailServer; overload;
    function RequireAuthentication: Boolean; overload;
  end;

  IGBMailFactory = interface
    ['{1E6828E6-3F29-44D5-A469-760C24474990}']
    function MailDefault: IGBMail;
    function MailIndy: IGBMail;
  end;

function GetMailFactory: IGBMailFactory;
function GetMailDefault: IGBMail;
function GBMailDefault: IGBMail;

implementation

uses
  GBMail.Factory;

function GetMailDefault: IGBMail;
begin
  Result := GetMailFactory.MailDefault;
end;

function GetMailFactory: IGBMailFactory;
begin
  Result := TGBMailFactory.New;
end;

function GBMailDefault: IGBMail;
begin
  Result := GetMailDefault;
end;

end.
