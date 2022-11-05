unit GBMail.Outlook;

interface

uses GBMail.Interfaces, GBMail.Base, System.SysUtils, System.Classes, Outlook2010,
     System.Variants;

type
  EGBMailOutlookException = class(EGBMailException);

  TGBMailOutlook = class(TGBMailBase, IGBMail)

  private
    FOpenOutlook: Boolean;

    procedure AddBody          (const email: MailItem);
    procedure AddToRecipients  (const email: MailItem);
    procedure AddCcRecipients  (const email: MailItem);
    procedure AddBccRecipients (const email: MailItem);
    procedure AddAttachments   (const email: MailItem);
  public
    function Send: IGBMail; override;

    constructor create(bOpenOutlook: Boolean = False);
    class function New(bOpenOutlook: Boolean = False): IGBMail;
    destructor Destroy; override;

end;

implementation

{ TGBMailOutlook }

procedure TGBMailOutlook.AddAttachments(const email: MailItem);
var
  i: Integer;
begin
  for i := 0 to Pred(GetAttachments.Count) do
  begin
    if FileExists(getAttachments[i]) then
      email.Attachments.Add(GetAttachments[i], EmptyParam, EmptyParam, EmptyParam);
  end;
end;

procedure TGBMailOutlook.AddBccRecipients(const email: MailItem);
var
  i: Integer;
begin
  for i := 0 to Pred(GetBccRecipients.Count) do
    with email.Recipients.Add(GetBccRecipients.Names[i]) do
      type_ := olBCC;
end;

procedure TGBMailOutlook.AddBody(const email: MailItem);
begin
  if getUseHtml then
  begin
    email.BodyFormat := olFormatHTML;
    email.HTMLBody   := GetMessage.Text;
  end
  else
  begin
    email.BodyFormat := olFormatPlain;
    email.Body       := GetMessage.Text;
  end;
end;

procedure TGBMailOutlook.AddCcRecipients(const email: MailItem);
var
  i: Integer;
begin
  for i := 0 to Pred(GetCcRecipients.Count) do
    with email.Recipients.Add(GetCcRecipients.Names[i]) do
      type_ := olCC;
end;

procedure TGBMailOutlook.AddToRecipients(const email: MailItem);
var
  i: Integer;
begin
  for i := 0 to Pred(GetRecipients.Count) do
    with email.Recipients.Add(GetRecipients.Names[i]) do
      type_ := olTo;
end;

constructor TGBMailOutlook.create(bOpenOutlook: Boolean);
begin
  FOpenOutlook := bOpenOutlook;
end;

destructor TGBMailOutlook.Destroy;
begin

  inherited;
end;

class function TGBMailOutlook.New(bOpenOutlook: Boolean): IGBMail;
begin
  result := Self.create(bOpenOutlook);
end;

function TGBMailOutlook.Send: IGBMail;
var
  outlook : TOutlookApplication;
  email   : MailItem;
begin
  result  := Self;
  outlook := TOutlookApplication.Create(nil);
  try
    try
      email := outlook.CreateItem(olMailItem) As MailItem;
      email.Subject := GetSubject;
      email.Importance := olImportanceNormal;

      AddBody(email);
      AddToRecipients(email);
      AddCcRecipients(email);
      AddBccRecipients(email);
      AddAttachments(email);

      if email.Recipients.ResolveAll then
      begin
        if FOpenOutlook then
          email.Display(False)
        else
          email.Send;
      end
      else
        raise EGBMailOutlookException.Create('Could not send email from Outlook. Check the parameters informed.');
    finally
      outlook.Disconnect;
    end;
  finally
    outlook.Free;
  end;
end;

end.
