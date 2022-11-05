unit GBMail.Indy;

interface

uses GBMail.Interfaces, GBMail.Base, IdSMTP, IdMessage, IdSSLOpenSSL,
     IdExplicitTLSClientServerBase,  IdText, IdAttachmentFile,
     System.SysUtils, System.Classes, IdAttachment, IdAttachmentMemory;

type
  EGBMailIndyException = class(EGBMailException);

  TGBMailIndy = class(TGBMailBase, IGBMail)
  private
    procedure configureMessage (const msg  : TIdMessage);
    procedure ConfigureSmtp    (const smtp : TIdSMTP);
    procedure AddToRecipients  (const msg  : TIdMessage);
    procedure AddCcRecipients  (const msg  : TIdMessage);
    procedure AddBccRecipients (const msg  : TIdMessage);
    procedure AddFrom          (const msg  : TIdMessage);
    procedure AddAttachments   (const msg  : TIdMessage);
    procedure AddHtmlFile      (const msg  : TIdMessage);
    procedure AddBody          (const msg  : TIdMessage);

  protected
    procedure InitializeISO88591(var VHeaderEncoding: Char; var VCharSet: string);

  public

    function Send: IGBMail; override;
    class function New: IGBMail;

end;

implementation

{ TGBMailIndy }

procedure TGBMailIndy.AddAttachments(const msg: TIdMessage);
var
  i         : Integer;
  attachment: TIdAttachmentFile;
  attachmentMemory: TIdAttachmentMemory;
  name: String;
begin
  for i := 0 to Pred(GetAttachments.Count) do
  begin
    if FileExists(getAttachments[i]) then
    begin
      attachment := TIdAttachmentFile.Create(msg.MessageParts, GetAttachments[i]);
      attachment.Headers.Add(Format('Content-ID: <%s>',
        [ExtractFileName(GetAttachments[i])]));
    end;
  end;

  for name in getAttachmentsMemory.Keys do
  begin
    attachmentMemory := TIdAttachmentMemory.Create(msg.MessageParts, getAttachmentsMemory.Items[name]);
    attachmentMemory.FileName := name;
  end;
end;

procedure TGBMailIndy.AddBccRecipients(const msg: TIdMessage);
var
  i: Integer;
begin
  for i := 0 to Pred(getBCcRecipients.Count) do
  begin
    with msg.BccList.Add do
    begin
      Address := getBCcRecipients.Names[i];
      Name    := getBCcRecipients.ValueFromIndex[i];
    end;
  end;
end;

procedure TGBMailIndy.AddBody(const msg: TIdMessage);
var
  body: TIdText;
begin
  body := TIdText.Create(msg.MessageParts);
  body.Body.Text :=  GetMessage.Text;
  if getUseHtml then
  begin
    body.ContentType := 'text/html';
    body.CharSet     := 'ISO-8859-1';
    body.ContentTransfer := '16bit';
  end
  else
    body.ContentType := 'text/plain';
end;

procedure TGBMailIndy.AddCcRecipients(const msg: TIdMessage);
var
  i: Integer;
begin
  for i := 0 to Pred(getCcRecipients.Count) do
    with msg.CCList.Add do
    begin
      Address := getCcRecipients.Names[i];
      Name    := getCcRecipients.ValueFromIndex[i];
    end;
end;

procedure TGBMailIndy.AddFrom(const msg: TIdMessage);
begin
  msg.From.Address := GetFromAddress;
  msg.From.Name := GetFromName;
end;

procedure TGBMailIndy.AddHtmlFile(const msg: TIdMessage);
var
  i    : Integer;
  image: TIdAttachment;
  name : String;
//  attachmentMemory: TIdAttachmentMemory;
begin
  for i := 0 to Pred(getHtmlImages.Count) do
  begin
    image := TIdAttachmentFile.Create(msg.MessageParts, getHtmlImages.ValueFromIndex[i]);
    image.ContentType := 'image/jpeg';
    image.ContentDisposition := 'inline';
    image.ExtraHeaders.Values['Content-ID'] := '<' + getHtmlImages.Names[i] + '>';
  end;

  for name in getHtmlImagesMemory.Keys do
  begin
    image := TIdAttachmentMemory.Create(msg.MessageParts,
                            getHtmlImagesMemory.Items[name]);
    image.ContentType := 'image/jpeg';
    image.ContentDisposition := 'inline';
    image.ExtraHeaders.Values['Content-ID'] := '<' + name.Replace('{', '').Replace('}', '') + '>';
  end;
end;

procedure TGBMailIndy.AddToRecipients(const msg: TIdMessage);
var
  i: Integer;
begin
  for i := 0 to Pred(getRecipients.Count) do
    with msg.Recipients.Add do
    begin
      Address := getRecipients.Names[i];
      Name    := getRecipients.ValueFromIndex[i];
    end;
end;

procedure TGBMailIndy.configureMessage(const msg: TIdMessage);
begin
  msg.Date        := Now;
  msg.Subject     := GetSubject;
  msg.ContentType := 'multipart/mixed';
//  msg.IsEncoded   := True;
//  msg.ContentTransferEncoding := '16bit';
  msg.CharSet     := 'utf-8';
  msg.Encoding    := meMIME;
//  msg.OnInitializeISO := InitializeISO88591;
end;

procedure TGBMailIndy.ConfigureSmtp(const smtp: TIdSMTP);
begin
  smtp.ConnectTimeout := Server.ConnectTimeOut;
  smtp.ReadTimeout    := Server.ReadTimeOut;
  smtp.Host           := Server.Host;
  smtp.Username       := Server.Username;
  smtp.Password       := Server.Password;
  smtp.Port           := Server.Port;

  if Server.RequireAuthentication then
    smtp.AuthType := satDefault
  else
    smtp.AuthType := satNone;

  smtp.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(smtp);
  if Server.UseSSL then
  begin
    TIdSSLIOHandlerSocketOpenSSL(smtp.IOHandler).SSLOptions.Method := sslvSSLv23;
    TIdSSLIOHandlerSocketOpenSSL(smtp.IOHandler).SSLOptions.Mode := sslmClient;
    smtp.UseTLS := utUseExplicitTLS;
  end;

  if Server.UseTLS then
  begin
    smtp.UseTLS := utUseRequireTLS;
//    TIdSSLIOHandlerSocketOpenSSL(smtp.IOHandler).SSLOptions.SSLVersions := [sslvTLSv1, sslvTLSv1_1, sslvTLSv1_2];
    TIdSSLIOHandlerSocketOpenSSL(smtp.IOHandler).SSLOptions.Method := sslvTLSv1_2;
    TIdSSLIOHandlerSocketOpenSSL(smtp.IOHandler).SSLOptions.Mode := sslmClient;
  end;
end;

procedure TGBMailIndy.InitializeISO88591(var VHeaderEncoding: Char; var VCharSet: string);
begin
  VCharSet := 'ISO-8859-1';
end;

class function TGBMailIndy.New: IGBMail;
begin
  result := Self.create;
end;

function TGBMailIndy.Send: IGBMail;
var
  smtp: TIdSMTP;
  msg : TIdMessage;
begin
  result := Self;
  smtp   := TIdSMTP.Create(nil);
  try
    try
      ConfigureSmtp(smtp);
      msg := TIdMessage.Create(nil);
      try
        configureMessage(msg);

        AddToRecipients(msg);
        AddCcRecipients(msg);
        AddBccRecipients(msg);
        AddFrom(msg);
        AddAttachments(msg);
        AddBody(msg);
        AddHtmlFile(msg);

        smtp.Connect;
        try
          if Server.RequireAuthentication then
            smtp.Authenticate;
          smtp.Send(msg);
        finally
          smtp.Disconnect;
        end;
      finally
        msg.Free;
      end;
    except
      on E: Exception do
        raise EGBMailIndyException.Create(E.Message);
    end;
  finally
    smtp.Free;
  end;
end;

end.
