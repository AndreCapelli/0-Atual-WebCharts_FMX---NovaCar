unit unit_esmaecer_fundo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.WinXCtrls, Vcl.ExtCtrls;

type
  Tform_esmaecer_fundo = class(TForm)
    ind: TActivityIndicator;
    procedure FormShow(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  form_esmaecer_fundo: Tform_esmaecer_fundo;

implementation

{$R *.dfm}

procedure Tform_esmaecer_fundo.FormHide(Sender: TObject);
begin
  Application.ProcessMessages;
  ind.Enabled := false;
end;

procedure Tform_esmaecer_fundo.FormResize(Sender: TObject);
begin
  Application.ProcessMessages;
  ind.left := trunc((form_esmaecer_fundo.width - ind.width) / 2);
  Application.ProcessMessages;
  ind.top := trunc((form_esmaecer_fundo.height - ind.height) / 2);
end;

procedure Tform_esmaecer_fundo.FormShow(Sender: TObject);
begin
  Application.ProcessMessages;
  ind.Enabled := true;
end;

end.
