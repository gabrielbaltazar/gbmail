unit GBMail.Base;

interface

uses
  GBMail.Interfaces,
  GBMail.Base.Server,
  System.Classes,
  System.SysUtils,
  System.StrUtils,
  System.Generics.Collections;

type
  TGBMailBase = class abstract(TInterfacedObject, IGBMail)
  protected
    FServer: IGBMailServer;
    FFromAddress: string;
    FFromName: string;
    FSubject: string;
    FUseHtml: Boolean;
    FRecipients: TStrings;
    FCcRecipients: TStrings;
    FBCcRecipients: TStrings;
    FMessage: TStrings;
    FAttachments: TStrings;
    FHtmlImages: TStringList;
    FAttachmentsStream: TObjectDictionary<string, TMemoryStream>;
    FHtmlImagesStream: TObjectDictionary<string, TMemoryStream>;
    procedure ClearStream;
  public
    constructor Create; virtual;
    class function New: IGBMail;
    destructor Destroy; override;

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
    function Send: IGBMail; virtual; abstract;
  end;

implementation

{ TGBMailBase }

function TGBMailBase.AddAttachment(AFileName: string): IGBMail;
begin
  Result := Self;
  if FileExists(AFileName) then
    FAttachments.Add(AFileName);
end;

function TGBMailBase.AddAttachment(AStream: TMemoryStream; AFileName: string): IGBMail;
var
  LName: string;
  LStream: TMemoryStream;
begin
  Result := Self;
  LStream := TMemoryStream.Create;
  try
    LStream.LoadFromStream(AStream);
    LStream.Position := 0;
    LName := ExtractFileName(AFileName);
    FAttachmentsStream.AddOrSetValue(LName, LStream);
  except
    LStream.Free;
    raise;
  end;
end;

function TGBMailBase.AddBccRecipient(AAddress, AName: string): IGBMail;
begin
  Result := Self;
  FBCcRecipients.Values[AAddress] := IfThen(AName.IsEmpty, AAddress, AName);
end;

function TGBMailBase.AddCcRecipient(AAddress, AName: string): IGBMail;
begin
  Result := Self;
  FCcRecipients.Values[AAddress] := IfThen(AName.IsEmpty, AAddress, AName);
end;

function TGBMailBase.AddHtmlImage(AImage: TMemoryStream; out AID: string): IGBMail;
var
  LStream: TMemoryStream;
begin
  Result := Self;
  LStream := TMemoryStream.Create;
  try
    LStream.LoadFromStream(AImage);
    LStream.Position := 0;
    AID := TGUID.NewGuid.ToString;
    FHtmlImagesStream.AddOrSetValue(AID, LStream);
  except
    LStream.Free;
    raise;
  end;
end;

function TGBMailBase.AddHtmlImage(AFileName: string; out AID: string): IGBMail;
begin
  Result := Self;
  AID := TGuid.NewGuid.ToString;
  FHtmlImages.AddPair(AID, AFileName);
end;

function TGBMailBase.AddRecipient(AAddress, AName: string): IGBMail;
begin
  Result := Self;
  FRecipients.Values[AAddress] := IfThen(AName.IsEmpty, AAddress, AName);
end;

procedure TGBMailBase.ClearStream;
var
  LKey: string;
begin
  for LKey in FAttachmentsStream.Keys do
    FAttachmentsStream.Items[LKey].Free;
  for LKey in FHtmlImagesStream.Keys do
    FHtmlImagesStream.Items[LKey].Free;
end;

constructor TGBMailBase.Create;
begin
  FServer := TGBMailServerBase.New(Self);
  FBCcRecipients := TStringList.Create;
  FCcRecipients := TStringList.Create;
  FRecipients := TStringList.Create;
  FMessage := TStringList.Create;
  FAttachments := TStringList.Create;
  FHtmlImages := TStringList.Create;
  FAttachmentsStream := TObjectDictionary<string, TMemoryStream>.create;
  FHtmlImagesStream := TObjectDictionary<string, TMemoryStream>.create;
end;

destructor TGBMailBase.Destroy;
begin
  FBCcRecipients.Free;
  FCcRecipients.Free;
  FRecipients.Free;
  FMessage.Free;
  FAttachments.Free;
  FHtmlImages.Free;
  ClearStream;
  FAttachmentsStream.Free;
  FHtmlImagesStream.Free;
  inherited;
end;

function TGBMailBase.From(AAddress, AName: string): IGBMail;
begin
  Result := Self;
  FFromAddress := AAddress;
  FFromName := IfThen(AName.IsEmpty, AAddress, AName);
end;

function TGBMailBase.Message(AValue: string): IGBMail;
begin
  Result := Self;
  FMessage.Text := AValue;
end;

function TGBMailBase.Message(AValue: TStrings): IGBMail;
begin
  Result := Self;
  FMessage.Text := AValue.Text;
end;

class function TGBMailBase.New: IGBMail;
begin
  Result := Self.Create;
end;

function TGBMailBase.Server: IGBMailServer;
begin
  Result := FServer;
end;

function TGBMailBase.Subject(AValue: string): IGBMail;
begin
  Result := Self;
  FSubject := AValue;
end;

function TGBMailBase.UseHtml(AValue: Boolean) : IGBMail;
begin
  Result := Self;
  FUseHtml := AValue;
end;

end.
