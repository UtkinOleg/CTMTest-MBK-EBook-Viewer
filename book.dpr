program book;

uses
  Forms,
  Classes,
  exebook in 'exebook.pas' {MainForm},
  BookSav in '..\GlobalUnit\BookSav.pas',
  NewBookMarkUnit in 'FORMS\NewBookMarkUnit.pas' {NewBookMarkForm},
  pass in '..\GlobalUnit\pass.pas',
  OpenF in 'FORMS\openf.pas',
  XMLUnit in '..\..\QGEDITLE\UNITS\XMLUnit.pas';

{$R *.RES}
var
 rs : TResourceStream;

begin
  Application.Initialize;
  Application.Title := 'СТМ-Тест электронный учебник';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TNewBookMarkForm, NewBookMarkForm);
  Application.Run;
end.
