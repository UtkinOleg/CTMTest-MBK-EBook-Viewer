unit NewBookMarkUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TNewBookMarkForm = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  NewBookMarkForm: TNewBookMarkForm;

implementation

{$R *.DFM}

procedure TNewBookMarkForm.FormShow(Sender: TObject);
begin
 Memo1.Clear;
 Memo1.SetFocus;
end;

end.
