unit Logger;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, IEScript20App_TLB, StdVcl, Classes;

type
  TIescriptLogger = class(TAutoObject, IIescriptLogger)
  private
    FStrings: TStrings;
  protected
    { Protected declarations }
    procedure log(const AMsg: WideString); safecall;
  public
    property Strings: TStrings read FStrings write FStrings;
  end;

implementation

uses ComServ;

procedure TIescriptLogger.log(const AMsg: WideString);
begin
  if Assigned(FStrings) then FStrings.Add(AMsg);
end;

initialization
  TAutoObjectFactory.Create(ComServer, TIescriptLogger, Class_IescriptLogger,
    ciMultiInstance, tmApartment);
end.
