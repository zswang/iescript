program IEScript20App;

uses
  Forms,
  IEScript20Unit in 'IEScript20Unit.pas' {FormIEScript},
  FavoriteDialog in 'FavoriteDialog.pas',
  Search in 'Search.pas',
  MSHTMCID in 'MSHTMCID.pas',
  WindowDialog in 'WindowDialog.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormIEScript, FormIEScript);
  Application.Run;
end.
