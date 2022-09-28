unit unitDmDados;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Phys.MSSQLDef, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.MSSQL, FireDAC.ConsoleUI.Wait, Data.DB,
  FireDAC.Comp.Client, NovaConexao,
  FireDAC.Phys.ODBCBase, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet, System.Json,
  Vcl.Dialogs, FireDAC.Comp.UI, Vcl.StdCtrls, FireDAC.VCLUI.Error;

type
  TdmDados = class(TDataModule)
    Conn: TFDConnection;
    FDPhysMSSQLDriverLink1: TFDPhysMSSQLDriverLink;
    qryGeral: TFDQuery;
    qryGeral2: TFDQuery;
    qryGeral3: TFDQuery;
    qryGeral4: TFDQuery;
    qryIdentCurrent: TFDQuery;
    FDErrorDialog: TFDGUIxErrorDialog;
    procedure DataModuleCreate(Sender: TObject);
  private
    procedure CriaPastaArquivosSistema;

    { Private declarations }
  public
    { Public declarations }
    function Cript(Valor: string): string;
    function Descript(Valor: string): string;
    function getCamposJsonString(Json, value: String): String;

    procedure GravaMsgArquivoAppLocal(Arquivo, Mensagem: string;
      { Apaga Conteudo antes
        de gravar a nova mensagem } ApagaConteudo: Boolean;
      CaminhoPassado: String = ''; ColocaData: Boolean = true;
      DataPassada: String = '');

  var
    RetornoID, RetornoBusca, RetornoBusca2: string;
    Acessou: Boolean;

  var
    GB_Servidor, GB_BancoDeDados, BD_Usuario, BD_senha, BD_servidor, BD_Base,
      GB_CaminhoAppLocalSistema, GB_UsuarioID: string;

    campanha_id, campanha_codigo: string;
  end;

var
  dmDados: TdmDados;

implementation

{%CLASSGROUP 'System.Classes.TPersistent'}

uses unitFuncoes;
{$R *.dfm}

Procedure TdmDados.CriaPastaArquivosSistema;
// Cria pasta para salvar os arquivos gerais do sistema APPLoca
var
  Caminho: string;
  exists: Boolean;
  F: textfile;
begin

  Caminho := StringReplace(GetEnvironmentVariable('TEMP'), 'Temp',
    'Calltech_Smart', [rfReplaceAll, rfIgnoreCase]);

  if not DirectoryExists(Caminho) then
  begin
    CreateDir(Caminho);
  end;

  dmDados.GB_CaminhoAppLocalSistema := Caminho;

end;

Procedure TdmDados.GravaMsgArquivoAppLocal(Arquivo, Mensagem: string;
  { Apaga Conteudo antes
    de gravar a nova mensagem } ApagaConteudo: Boolean;
  CaminhoPassado: String = ''; ColocaData: Boolean = true;
  DataPassada: String = '');
var
  F: textfile;
  vArquivo: TStringList;
  Caminho: string;
begin

  Arquivo := StringReplace(Arquivo, '/', '', [rfReplaceAll, rfIgnoreCase]);

  if CaminhoPassado = '' then
    Caminho := dmDados.GB_CaminhoAppLocalSistema + '\' + Arquivo
  else
    Caminho := CaminhoPassado + '\' + Arquivo;

  If not(fileexists(Caminho)) then
  begin
    AssignFile(F, Caminho);
    Rewrite(F);
    // abre o arquivo para escrita
    Closefile(F); // fecha o handle de arquivo
  end;

  if ApagaConteudo = true then
  begin
    AssignFile(F, Caminho);
    reset(F);
    Rewrite(F);
    Writeln(F, Mensagem);
    Closefile(F);
  end
  else
  begin
    try
      vArquivo := TStringList.Create;
      vArquivo.LoadFromFile(Caminho);
      vArquivo.Add(Mensagem + IifStr(ColocaData, IifStr(DataPassada = '',
        ' ' + datetimetostr(Now()), ' ' + DataPassada), ''));
      vArquivo.SaveToFile(Caminho);
      vArquivo.Free;
    except
      vArquivo.Free;
    end;
  end;
end;

procedure TdmDados.DataModuleCreate(Sender: TObject);
begin
  Conn.Params.Pooled := true;
  BuildConnection('maxOne'); // Chama o strCon.exe se necessário
  StartConnection(Conn, FDPhysMSSQLDriverLink1);

  // dmDados.Conn.Connected := false;

  CriaPastaArquivosSistema;

end;

Function TdmDados.Cript(Valor: string): string;
var
  Sai, texto: string;
  Semente, x: Single;
  c, i, k: integer;
begin
  Sai := '';
  Semente := 0.2;
  texto := Valor;

  for i := 1 to Length(texto) do
  begin
    k := Ord(texto[i]);
    x := 9821 * Semente + 0.211327;
    Semente := x - int(x);
    x := int(95 * Semente) + 1;
    c := k + Trunc(x);
    if c >= 124 then
    begin
      while c >= 124 do
      begin
        c := c - 92;
      end;
    end;

    if c < 32 then
    begin
      c := c + 92
    end;

    Sai := Sai + Chr(c);
  end;
  Result := Sai;
end;

Function TdmDados.Descript(Valor: string): string;
var
  Sai, texto: string;
  Semente, x: Single;
  c, i, k: integer;
begin
  Sai := '';
  Semente := 0.2;
  texto := Valor;

  for i := 1 to Length(texto) do
  begin
    k := Ord(texto[i]);
    x := 9821 * Semente + 0.211327;
    Semente := x - int(x);
    x := int(95 * Semente) + 1;
    c := k + Trunc(-x);
    if c >= 124 then
    begin
      while c >= 124 do
      begin
        c := c - 92;
      end;
    end;

    if c < 32 then
    begin
      c := c + 92
    end;

    Sai := Sai + Chr(c);
  end;
  Result := Sai;
end;

function TdmDados.getCamposJsonString(Json, value: String): String;
var
  LJSONObject: TJSonObject;
  function TrataObjeto(jObj: TJSonObject): string;
  var
    i: integer;
    jPar: TJSONPair;
  begin
    Result := '';
    for i := 0 to jObj.size - 1 do
    begin
      jPar := jObj.Get(i);
      if jPar.JSonValue Is TJSonObject then
        Result := TrataObjeto((jPar.JSonValue As TJSonObject))
      else if sametext(trim(jPar.JsonString.value), value) then
      begin
        Result := jPar.JSonValue.value;
        break;
      end;
      if Result <> '' then
        break;
    end;
  end;

begin
  try
    LJSONObject := nil;
    LJSONObject := TJSonObject.ParseJSONValue(TEncoding.ASCII.GetBytes(Json), 0)
      as TJSonObject;
    Result := TrataObjeto(LJSONObject);
    if Result = 'null' then
      Result := '';
  finally
    LJSONObject.Free;
  end;
end;

end.
