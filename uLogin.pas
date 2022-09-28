unit uLogin;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes,
  System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.Objects,
  FMX.Layouts, FMX.Edit, FMX.Controls.Presentation, FMX.StdCtrls,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error,
  FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf, FireDAC.Stan.Async,
  FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client;

type
  TfrmLogin = class(TForm)
    Layout1: TLayout;
    GradienteFundo: TRectangle;
    laCentro: TLayout;
    laUser: TLayout;
    laEntrar: TLayout;
    laSenha: TLayout;
    laTitulo: TLayout;
    RoundRect2: TRoundRect;
    txtUser: TEdit;
    Label2: TLabel;
    StyleBook1: TStyleBook;
    RoundRect4: TRoundRect;
    txtSenha: TEdit;
    Button1: TButton;
    roundEntrar: TRoundRect;
    Label3: TLabel;
    qry: TFDQuery;
    procedure GradienteFundoMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Single);
    procedure Button1Click(Sender: TObject);
    procedure roundEntrarMouseLeave(Sender: TObject);
    procedure roundEntrarMouseEnter(Sender: TObject);
    procedure roundEntrarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLogin: TfrmLogin;

implementation

{$R *.fmx}

uses unitDmDados, unitFuncoes, unitPrincipal;

procedure TfrmLogin.Button1Click(Sender: TObject);
begin
  frmLogin.Close;
end;

procedure TfrmLogin.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if dmdados.Acessou = false then
    Application.Terminate
  else
  begin
    unitHome.Show;
    unitHome.Timer1.Enabled := true;
  end;
end;

procedure TfrmLogin.FormCreate(Sender: TObject);
begin
  Application.CreateForm(TdmDados, dmdados);
end;

procedure TfrmLogin.FormShow(Sender: TObject);
begin

  unitHome.Hide;

end;

procedure TfrmLogin.GradienteFundoMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Single);
const
  sc_DragMove = $F012;
begin
  ReleaseCapture;
  Self.StartWindowDrag;
end;

procedure TfrmLogin.roundEntrarClick(Sender: TObject);
begin

  dmdados.Acessou := false;
  if txtUser.Text = '' then
  begin
    MsgErro('Informe o usuário!');
    exit;
  end;

  if txtSenha.Text = '' then
  begin
    MsgErro('Informe a senha!');
    exit;
  end;

  LimpaQuery(qry);

  qry.Open('Select * From Usuarios With(NOLOCK) Where UsLogin=' +
    QuotedStr(txtUser.Text) + ' AND UsSenha=' +
    QuotedStr(dmdados.Cript(txtSenha.Text)));

  if qry.Eof then
  begin
    MsgErro('Usuário ou senha inválidos!');
    qry.Close;
    exit;
  end;

  dmdados.GB_UsuarioID := qry.FieldByName('usuarios_ID').AsString;

  LimpaQuery(qry);
  qry.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) '
    + ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '
    + dmdados.GB_UsuarioID + ' WHERE CaEncerrada = 0 ');

  if qry.RecordCount = 0 then
  Begin
    MsgErro('Usuário não está associado em nenhuma Campanha');
    qry.Close;
    exit;
  End;

  qry.Close;

  dmdados.Acessou := true;

  frmLogin.Close;

end;

procedure TfrmLogin.roundEntrarMouseEnter(Sender: TObject);
begin
  roundEntrar.Opacity := 0.6;
end;

procedure TfrmLogin.roundEntrarMouseLeave(Sender: TObject);
begin
  roundEntrar.Opacity := 1;
end;

end.
