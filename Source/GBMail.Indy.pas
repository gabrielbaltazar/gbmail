unit GBMail.Indy;
interface
uses
  GBMail.Interfaces,
  GBMail.Base,
  IdSMTP,
  IdMessage,
  IdSSLOpenSSL,
  IdExplicitTLSClientServerBase,
  IdText,
  IdAttachment,
  IdAttachmentFile,
  IdAttachmentMemory,
  System.SysUtils,
  System.Classes;
type
  EGBMailIndyException = class(EGBMailException);
  TGBMailIndy = class(TGBMailBase, IGBMail)
  private
    procedure ConfigureMessage(const AMessage: TIdMessage);
    procedure ConfigureSmtp(const ASMTP: TIdSMTP);
    procedure AddToRecipients(const AMessage: TIdMessage);
    procedure AddCcRecipients(const AMessage: TIdMessage);
    procedure AddBccRecipients(const AMessage: TIdMessage);
    procedure AddFrom(const AMessage: TIdMessage);
    procedure AddAttachments(const AMessage: TIdMessage);
    procedure AddHtmlFile(const AMessage: TIdMessage);
    procedure AddBody(const AMessage: TIdMessage);
  protected
    procedure InitializeISO88591(var AHeaderEncoding: Char; var ACharSet: string);
  public
    class function New: IGBMail;
    function Send: IGBMail; override;
  end;
implementation
{ TGBMailIndy }
procedure TGBMailIndy.AddAttachments(const AMessage: TIdMessage);
var
  I: Integer;
  LAttachment: TIdAttachmentFile;
  LAttachmentMemory: TIdAttachmentMemory;
  LName: string;
begin
  for I := 0 to Pred(FAttachments.Count) do
  begin
    if FileExists(FAttachments[I]) then
    begin
      LAttachment := TIdAttachmentFile.Create(AMessage.MessageParts, FAttachments[I]);
      LAttachment.Headers.Add(Format('Content-ID: <%s>',
        [ExtractFileName(FAttachments[I])]));
    end;
  end;
  for LName in FAttachmentsStream.Keys do
  begin
    LAttachmentMemory := TIdAttachmentMemory.Create(AMessage.MessageParts, FAttachmentsStream.Items[LName]);
    LAttachmentMemory.FileName := LName;
  end;
end;
procedure TGBMailIndy.AddBccRecipients(const AMessage: TIdMessage);
var
  I: Integer;
begin
  for I := 0 to Pred(FBCcRecipients.Count) do
  begin
    with AMessage.BccList.Add do
    begin
      Address := FBCcRecipients.Names[I];
      Name := FBCcRecipients.ValueFromIndex[I];
    end;
  end;
end;
procedure TGBMailIndy.AddBody(const AMessage: TIdMessage);
var
  LBody: TIdText;
begin
  LBody := TIdText.Create(AMessage.MessageParts);
  LBody.Body.Text := FMessage.Text;
  if FUseHtml then
  begin
    LBody.ContentType := 'text/html';
    LBody.CharSet := 'ISO-8859-1';
    LBody.ContentTransfer := '16bit';
  end
  else
    LBody.ContentType := 'text/plain';
end;
procedure TGBMailIndy.AddCcRecipients(const AMessage: TIdMessage);
var
  I: Integer;
begin
  for I := 0 to Pred(FCcRecipients.Count) do
    with AMessage.CCList.Add do
    begin
      Address := FCcRecipients.Names[I];
      Name := FCcRecipients.ValueFromIndex[I];
    end;
end;
procedure TGBMailIndy.AddFrom(const AMessage: TIdMessage);
begin
  AMessage.From.Address := FFromAddress;
  AMessage.From.Name := FFromName;
end;
procedure TGBMailIndy.AddHtmlFile(const AMessage: TIdMessage);
var
  I: Integer;
  LImage: TIdAttachment;
  LName: string;
begin
  for I := 0 to Pred(FHtmlImages.Count) do
  begin
    LImage := TIdAttachmentFile.Create(AMessage.MessageParts, FHtmlImages.ValueFromIndex[I]);
    LImage.ContentType := 'image/jpeg';
    LImage.ContentDisposition := 'inline';
    LImage.ExtraHeaders.Values['Content-ID'] := '<' + FHtmlImages.Names[I] + '>';
  end;
  for LName in FHtmlImagesStream.Keys do
  begin
    LImage := TIdAttachmentMemory
      .Create(AMessage.MessageParts, FHtmlImagesStream.Items[LName]);
    LImage.ContentType := 'image/jpeg';
    LImage.ContentDisposition := 'inline';
    LImage.ExtraHeaders.Values['Content-ID'] := '<' + LName.Replace('{', '').Replace('}', '') + '>';
  end;
end;
procedure TGBMailIndy.AddToRecipients(const AMessage: TIdMessage);
var
  I: Integer;
begin
  for I := 0 to Pred(FRecipients.Count) do
    with AMessage.Recipients.Add do
    begin
      Address := FRecipients.Names[I];
      Name := FRecipients.ValueFromIndex[I];
    end;
end;
procedure TGBMailIndy.ConfigureMessage(const AMessage: TIdMessage);
begin
  AMessage.Date := Now;
  AMessage.Subject := FSubject;
  AMessage.ContentType := 'multipart/mixed';
//  msg.IsEncoded   := True;
//  msg.ContentTransferEncoding := '16bit';
  AMessage.CharSet := 'utf-8';
  AMessage.Encoding := meMIME;
//  msg.OnInitializeISO := InitializeISO88591;
end;
procedure TGBMailIndy.ConfigureSmtp(const ASMTP: TIdSMTP);
begin
  ASMTP.ConnectTimeout := Server.ConnectTimeOut;
  ASMTP.ReadTimeout := Server.ReadTimeOut;
  ASMTP.Host := Server.Host;
  ASMTP.Username := Server.Username;
  ASMTP.Password := Server.Password;
  ASMTP.Port := Server.Port;

  if Server.RequireAuthentication then
    ASMTP.AuthType := satDefault
  else
    ASMTP.AuthType := satNone;
  ASMTP.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(ASMTP);
  if Server.UseSSL then
  begin
    TIdSSLIOHandlerSocketOpenSSL(ASMTP.IOHandler).SSLOptions.Method := sslvSSLv23;
    TIdSSLIOHandlerSocketOpenSSL(ASMTP.IOHandler).SSLOptions.Mode := sslmClient;
    ASMTP.UseTLS := utUseExplicitTLS;
  end;
  if Server.UseTLS then
  begin
    ASMTP.UseTLS := utUseRequireTLS;
//    TIdSSLIOHandlerSocketOpenSSL(smtp.IOHandler).SSLOptions.SSLVersions := [sslvTLSv1, sslvTLSv1_1, sslvTLSv1_2];
    TIdSSLIOHandlerSocketOpenSSL(ASMTP.IOHandler).SSLOptions.Method := sslvTLSv1_2;
    TIdSSLIOHandlerSocketOpenSSL(ASMTP.IOHandler).SSLOptions.Mode := sslmClient;
  end;
end;
procedure TGBMailIndy.InitializeISO88591(var AHeaderEncoding: Char; var ACharSet: string);
begin
  ACharSet := 'ISO-8859-1';
end;
class function TGBMailIndy.New: IGBMail;
begin
  Result := Self.Create;
end;
function TGBMailIndy.Send: IGBMail;
var
  LSMTP: TIdSMTP;
  LMessage: TIdMessage;
begin
  Result := Self;
  LSMTP := TIdSMTP.Create(nil);
  try
    try
      ConfigureSmtp(LSMTP);
      LMessage := TIdMessage.Create(nil);
      try
        ConfigureMessage(LMessage);
        AddToRecipients(LMessage);
        AddCcRecipients(LMessage);
        AddBccRecipients(LMessage);
        AddFrom(LMessage);
        AddAttachments(LMessage);
        AddBody(LMessage);
        AddHtmlFile(LMessage);
        LSMTP.Connect;
        try
          if Server.RequireAuthentication then
            LSMTP.Authenticate;
          LSMTP.Send(LMessage);
        finally
          LSMTP.Disconnect;
        end;
      finally
        LMessage.Free;
      end;
    except
      on E: Exception do
        raise EGBMailIndyException.Create(E.Message);
    end;
  finally
    LSMTP.Free;
  end;
end;

end.
