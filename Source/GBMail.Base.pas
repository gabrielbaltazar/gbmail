unit GBMail.Base;

interface

uses GBMail.Interfaces, GBMail.Base.Server, System.Classes, System.SysUtils,
     System.StrUtils, System.Generics.Collections;

type TGBMailBase = class abstract(TInterfacedObject, IGBMail)

  private
    FServer        : IGBMailServer;
    FFromAddress   : string;
    FFromName      : string;
    FSubject       : string;
    FUseHtml       : Boolean;
    FRecipients    : TStrings;
    FCcRecipients  : TStrings;
    FBCcRecipients : TStrings;
    FMessage       : TStrings;
    FAttachments   : TStrings;
    FHtmlImages    : TStringList;

    FAttachmentsStream: TObjectDictionary<String, TMemoryStream>;
    FHtmlImagesStream: TObjectDictionary<String, TMemoryStream>;

  protected
    function getFromAddress   : string;
    function getFromName      : string;
    function getSubject       : string;
    function getUseHtml       : Boolean;
    function getRecipients    : TStrings;
    function getCcRecipients  : TStrings;
    function getBCcRecipients : TStrings;
    function getMessage       : TStrings;
    function getAttachments   : TStrings;
    function getHtmlImages    : TStrings;
    function getHtmlImagesMemory: TObjectDictionary<string, TMemoryStream>;
    function getAttachmentsMemory: TObjectDictionary<string, TMemoryStream>;

    procedure ClearStream;
  public
    function Server: IGBMailServer;

    function From           (Address: string; Name: String = ''): IGBMail;
    function AddRecipient   (Address: string; Name: string = ''): IGBMail;
    function AddCcRecipient (Address: string; Name: string = ''): IGBMail;
    function AddBccRecipient(Address: string; Name: string = ''): IGBMail;
    function AddAttachment  (FileName: String): IGBMail; overload;
    function AddAttachment  (Stream: TMemoryStream; FileName: String): IGBMail; overload;

    function AddHtmlImage   (FileName: String; out ID: String): IGBMail; overload;
    function AddHtmlImage   (Image: TMemoryStream; out ID: String): IGBMail; overload;

    function UseHtml(Value: Boolean) : IGBMail;
    function Subject(Value: String)  : IGBMail;
    function Message(Value: String)  : IGBMail; overload;
    function Message(Value: TStrings): IGBMail; overload;

    function Send: IGBMail; virtual; abstract;

    constructor create; virtual;
    class function New: IGBMail;
    destructor Destroy; override;

end;

implementation

{ TGBMailBase }

function TGBMailBase.AddAttachment(FileName: String): IGBMail;
begin
  Result := Self;
  if FileExists(FileName) then
    FAttachments.Add(FileName);
end;

function TGBMailBase.AddAttachment(Stream: TMemoryStream; FileName: String): IGBMail;
var
  name: string;
  memoryStream: TMemoryStream;
begin
  result := Self;
  memoryStream := TMemoryStream.Create;
  try
    memoryStream.LoadFromStream(Stream);
    memoryStream.Position := 0;

    name := ExtractFileName(FileName);
    FAttachmentsStream.AddOrSetValue(name, memoryStream);
  except
    memoryStream.Free;
    raise;
  end;
end;

function TGBMailBase.AddBccRecipient(Address, Name: string): IGBMail;
begin
  result := Self;
  FBCcRecipients.Values[Address] := IfThen(Name.IsEmpty, Address, Name);
end;

function TGBMailBase.AddCcRecipient(Address, Name: string): IGBMail;
begin
  result := Self;
  FCcRecipients.Values[Address] := IfThen(Name.IsEmpty, Address, Name);
end;

function TGBMailBase.AddHtmlImage(Image: TMemoryStream; out ID: String): IGBMail;
var
  memoryStream: TMemoryStream;
begin
  result := Self;
  memoryStream := TMemoryStream.Create;
  try
    memoryStream.LoadFromStream(Image);
    memoryStream.Position := 0;
    ID := TGUID.NewGuid.ToString;
    FHtmlImagesStream.AddOrSetValue(ID, memoryStream);
  except
    memoryStream.Free;
    raise;
  end;
end;

function TGBMailBase.AddHtmlImage(FileName: String; out ID: String): IGBMail;
begin
  result := Self;

  ID := TGuid.NewGuid.ToString;
  FHtmlImages.AddPair(ID, FileName);
end;

function TGBMailBase.AddRecipient(Address, Name: string): IGBMail;
begin
  result := Self;
  FRecipients.Values[Address] := IfThen(Name.IsEmpty, Address, Name);
end;

procedure TGBMailBase.ClearStream;
var
  key: String;
begin
  for key in FAttachmentsStream.Keys do
    FAttachmentsStream.Items[key].Free;

  for key in FHtmlImagesStream.Keys do
    FHtmlImagesStream.Items[key].Free;
end;

constructor TGBMailBase.create;
begin
  FServer        := TGBMailServerBase.New(Self);
  FBCcRecipients := TStringList.Create;
  FCcRecipients  := TStringList.Create;
  FRecipients    := TStringList.Create;
  FMessage       := TStringList.Create;
  FAttachments   := TStringList.Create;
  FHtmlImages    := TStringList.Create;
  FAttachmentsStream:= TObjectDictionary<String, TMemoryStream>.create;
  FHtmlImagesStream:= TObjectDictionary<String, TMemoryStream>.create;
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

function TGBMailBase.From(Address, Name: String): IGBMail;
begin
  result       := Self;
  FFromAddress := Address;
  FFromName    := IfThen(Name.IsEmpty, Address, Name);
end;

function TGBMailBase.getAttachments: TStrings;
begin
  result := FAttachments;
end;

function TGBMailBase.getAttachmentsMemory: TObjectDictionary<string, TMemoryStream>;
begin
  result := FAttachmentsStream;
end;

function TGBMailBase.getBCcRecipients: TStrings;
begin
  result := FBCcRecipients;
end;

function TGBMailBase.getCcRecipients: TStrings;
begin
  result := FCcRecipients;
end;

function TGBMailBase.getFromAddress: string;
begin
  Result := FFromAddress;
end;

function TGBMailBase.getFromName: string;
begin
  result := FFromName;
end;

function TGBMailBase.getHtmlImages: TStrings;
begin
  result := FHtmlImages;
end;

function TGBMailBase.getHtmlImagesMemory: TObjectDictionary<string, TMemoryStream>;
begin
  result := FHtmlImagesStream;
end;

function TGBMailBase.getMessage: TStrings;
begin
  result := FMessage;
end;

function TGBMailBase.getRecipients: TStrings;
begin
  result := FRecipients;
end;

function TGBMailBase.getSubject: string;
begin
  result := FSubject;
end;

function TGBMailBase.getUseHtml: Boolean;
begin
  result := FUseHtml;
end;

function TGBMailBase.Message(Value: String): IGBMail;
begin
  result        := Self;
  FMessage.Text := Value;
end;

function TGBMailBase.Message(Value: TStrings): IGBMail;
begin
  result        := Self;
  FMessage.Text := Value.Text;
end;

class function TGBMailBase.New: IGBMail;
begin
  result := Self.create;
end;

function TGBMailBase.Server: IGBMailServer;
begin
  result := FServer;
end;

function TGBMailBase.Subject(Value: String): IGBMail;
begin
  result   := Self;
  FSubject := Value;
end;

function TGBMailBase.UseHtml(Value: Boolean) : IGBMail;
begin
  Result   := Self;
  FUseHtml := Value;
end;

end.
