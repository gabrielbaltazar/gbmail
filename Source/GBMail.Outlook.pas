unit GBMail.Outlook;

interface

uses
  GBMail.Interfaces,
  GBMail.Base,
  Outlook2010,
  System.SysUtils,
  System.Classes,
  System.Variants;

type
  EGBMailOutlookException = class(EGBMailException);

  TGBMailOutlook = class(TGBMailBase, IGBMail)
  private
    FOpenOutlook: Boolean;

    procedure AddBody(const AEmail: MailItem);
    procedure AddToRecipients(const AEmail: MailItem);
    procedure AddCcRecipients(const AEmail: MailItem);
    procedure AddBccRecipients(const AEmail: MailItem);
    procedure AddAttachments(const AEmail: MailItem);
  public
    constructor Create(AOpenOutlook: Boolean = False); reintroduce;
    class function New(AOpenOutlook: Boolean = False): IGBMail;
    destructor Destroy; override;

    function Send: IGBMail; override;
  end;

implementation

{ TGBMailOutlook }

procedure TGBMailOutlook.AddAttachments(const AEmail: MailItem);
var
  I: Integer;
begin
  for I := 0 to Pred(FAttachments.Count) do
  begin
    if FileExists(FAttachments[I]) then
      AEmail.Attachments.Add(FAttachments[I], EmptyParam, EmptyParam, EmptyParam);
  end;
end;

procedure TGBMailOutlook.AddBccRecipients(const AEmail: MailItem);
var
  I: Integer;
begin
  for I := 0 to Pred(FBccRecipients.Count) do
    AEmail.Recipients.Add(FBccRecipients.Names[I]).type_ := olBCC;
end;

procedure TGBMailOutlook.AddBody(const AEmail: MailItem);
begin
  if FUseHtml then
  begin
    AEmail.BodyFormat := olFormatHTML;
    AEmail.HTMLBody := FMessage.Text;
  end
  else
  begin
    AEmail.BodyFormat := olFormatPlain;
    AEmail.Body := FMessage.Text;
  end;
end;

procedure TGBMailOutlook.AddCcRecipients(const AEmail: MailItem);
var
  I: Integer;
begin
  for I := 0 to Pred(FCcRecipients.Count) do
    AEmail.Recipients.Add(FCcRecipients.Names[I]).type_ := olCC;
end;

procedure TGBMailOutlook.AddToRecipients(const AEmail: MailItem);
var
  I: Integer;
begin
  for I := 0 to Pred(FRecipients.Count) do
    AEmail.Recipients.Add(FRecipients.Names[I]).type_ := olTo;
end;

constructor TGBMailOutlook.Create(AOpenOutlook: Boolean);
begin
  inherited Create;
  FOpenOutlook := AOpenOutlook;
end;

destructor TGBMailOutlook.Destroy;
begin
  inherited;
end;

class function TGBMailOutlook.New(AOpenOutlook: Boolean): IGBMail;
begin
  Result := Self.Create(AOpenOutlook);
end;

function TGBMailOutlook.Send: IGBMail;
var
  LOutlook: TOutlookApplication;
  LEmail: MailItem;
begin
  Result := Self;
  LOutlook := TOutlookApplication.Create(nil);
  try
    try
      LEmail := LOutlook.CreateItem(olMailItem) As MailItem;
      LEmail.Subject := FSubject;
      LEmail.Importance := olImportanceNormal;

      AddBody(LEmail);
      AddToRecipients(LEmail);
      AddCcRecipients(LEmail);
      AddBccRecipients(LEmail);
      AddAttachments(LEmail);

      if LEmail.Recipients.ResolveAll then
      begin
        if FOpenOutlook then
          LEmail.Display(False)
        else
          LEmail.Send;
      end
      else
        raise EGBMailOutlookException.Create('Could not send email from Outlook. Check the parameters informed.');
    finally
      LOutlook.Disconnect;
    end;
  finally
    LOutlook.Free;
  end;
end;

end.
