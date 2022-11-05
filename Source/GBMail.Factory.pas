unit GBMail.Factory;

interface

uses
  GBMail.Interfaces,
  GBMail.Indy;

type
  TGBMailFactory = class(TInterfacedObject, IGBMailFactory)
  protected
    function MailDefault: IGBMail;
    function MailIndy: IGBMail;
  public
    class function New: IGBMailFactory;
  end;

implementation

{ TGBMailFactory }

function TGBMailFactory.MailDefault: IGBMail;
begin
  Result := MailIndy;
end;

function TGBMailFactory.MailIndy: IGBMail;
begin
  Result := TGBMailIndy.New;
end;

class function TGBMailFactory.New: IGBMailFactory;
begin
  Result := Self.Create;
end;

end.
