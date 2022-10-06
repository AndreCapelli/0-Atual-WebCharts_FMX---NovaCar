unit unitPrincipal;

interface

uses
  {$IFDEF MSWINDOWS}
    Winapi.Windows,
    Winapi.Messages,
  {$ENDIF}
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  uCEFChromiumCore, uCEFFMXChromium, FMX.StdCtrls, FMX.Controls.Presentation,
  uCEFFMXWindowParent, View.WebCharts, Data.DB, Datasnap.DBClient,
  uCEFInterfaces, uCEFTypes, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.UI.Intf,
  FireDAC.FMXUI.Wait, FireDAC.Comp.UI, FMX.Layouts, FMX.Objects,
  FMX.DateTimeCtrls, System.DateUtils, Utilities.Encoder, FMX.ListBox;
const
  CEF_SHOWBROWSER = WM_APP + $101;

type
  TunitHome = class(TForm)
    FMXChromium1: TFMXChromium;
    WebCharts1: TWebCharts;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    StyleBook1: TStyleBook;
    SaveDialog1: TSaveDialog;
    qryGeral: TFDQuery;
    qryGeral1: TFDQuery;
    qryGeral2: TFDQuery;
    qryGeral3: TFDQuery;
    qryGeral4: TFDQuery;
    qryGeral5: TFDQuery;
    qryGeral6: TFDQuery;
    Timer1: TTimer;
    qryCamp: TFDQuery;
    qryGrRO: TFDQuery;
    qryRO: TFDQuery;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FMXChromium1Close(Sender: TObject; const browser: ICefBrowser;
      var aAction: TCefCloseBrowserAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FMXChromium1BeforeClose(Sender: TObject;
      const browser: ICefBrowser);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    FPivotConfig : string;
    FPivotConfig2 : string;

  public
    { Public declarations }

    var
      fontFamily: String;
      dateInicial, dateFinal: String;

      campanhaCodigo, campanhaNome, campanhasID: String;

    procedure GraphOne(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure GraphTwo(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure GraphThree(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure GraphFour(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure GraphFive(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure GraphFiveI(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure GraphSix(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure GraphSeven(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');

    procedure AttGraphOne(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure AttGraphTwo(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
    procedure AttGraphTwoII(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
    procedure AttGraphTwoIII(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
    procedure AttGraphTwoIV(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
    procedure AttGraphThree(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure AttGraphFour(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure AttGraphFive(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
    procedure AttGraphFiveI(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');

    procedure AtualizaCadastros(mes: String = '');
    procedure AtualizaContatos(mes: String = ''; caCodigo: String = ''; campID: String = '');
    procedure AtualizaGrupos(mes: String = ''; caCodigo: String = ''; campID: String = ''; gruposID: String = '');


    procedure NotifyMoveOrResizeStarted;
    procedure SetBounds(ALeft: Integer; ATop: Integer; AWidth: Integer; AHeight: Integer); override;


  protected
    FCanClose : boolean;
    FClosing  : boolean;

    FMXWindowParent         : TFMXWindowParent;

    {$IFDEF MSWINDOWS}
      FCustomWindowState      : TWindowState;
      FOldWndPrc              : TFNWndProc;
      FFormStub               : Pointer;
    {$ENDIF}

    procedure ResizeChild;
    procedure CreateFMXWindowParent;
    function  GetFMXWindowParentRect : System.Types.TRect;
    function  PostCustomMessage(aMsg : cardinal; aWParam : WPARAM = 0; aLParam : LPARAM = 0) : boolean;

    {$IFDEF MSWINDOWS}
      function  GetCurrentWindowState : TWindowState;
      procedure UpdateCustomWindowState;
      procedure CreateHandle; override;
      procedure DestroyHandle; override;
      procedure CustomWndProc(var aMessage: TMessage);
    {$ENDIF}

  end;

var
  unitHome: TunitHome;

implementation

uses
  Interfaces,
  Charts.Types,
  uCEFApplication,
  uCEFConstants,
  uCEFMiscFunctions,
  FMX.Platform.Win, NovaConexao, unitDmDados, unitFuncoes,
  unit_esmaecer_fundo, uLogin;

{$R *.fmx}
{$R *.Windows.fmx MSWINDOWS}

{ TForm9 }


procedure TunitHome.CreateFMXWindowParent;
begin
  if (FMXWindowParent = nil) then
    begin
      FMXWindowParent := TFMXWindowParent.CreateNew(nil);
      FMXWindowParent.Reparent(Handle);
      ResizeChild;
      FMXWindowParent.Show;
    end;

end;

{$IFDEF MSWINDOWS}
function TunitHome.GetCurrentWindowState: TWindowState;
var
  TempPlacement : TWindowPlacement;
  TempHWND      : HWND;
begin
  Result   := TWindowState.wsNormal;
  TempHWND := FmxHandleToHWND(Handle);

  ZeroMemory(@TempPlacement, SizeOf(TWindowPlacement));
  TempPlacement.Length := SizeOf(TWindowPlacement);

  if GetWindowPlacement(TempHWND, @TempPlacement) then
    case TempPlacement.showCmd of
      SW_SHOWMAXIMIZED : Result := TWindowState.wsMaximized;
      SW_SHOWMINIMIZED : Result := TWindowState.wsMinimized;
    end;

  if IsIconic(TempHWND) then Result := TWindowState.wsMinimized;
end;

procedure TunitHome.UpdateCustomWindowState;
var
  TempNewState : TWindowState;
begin
  TempNewState := GetCurrentWindowState;

  if (FCustomWindowState <> TempNewState) then
    begin
      if (FCustomWindowState = TWindowState.wsMinimized) then
        PostCustomMessage(CEF_SHOWBROWSER);

      FCustomWindowState := TempNewState;
    end;
end;

procedure TunitHome.CreateHandle;
begin
  inherited CreateHandle;

  FFormStub  := MakeObjectInstance(CustomWndProc);
  FOldWndPrc := TFNWndProc(SetWindowLongPtr(FmxHandleToHWND(Handle), GWLP_WNDPROC, NativeInt(FFormStub)));
end;

procedure TunitHome.DestroyHandle;
begin
  SetWindowLongPtr(FmxHandleToHWND(Handle), GWLP_WNDPROC, NativeInt(FOldWndPrc));
  FreeObjectInstance(FFormStub);

  inherited DestroyHandle;
end;

procedure TunitHome.CustomWndProc(var aMessage: TMessage);
const
  SWP_STATECHANGED = $8000;
var
  TempWindowPos : PWindowPos;
begin
  try
    case aMessage.Msg of
      WM_ENTERMENULOOP :
        if (aMessage.wParam = 0) and
           (GlobalCEFApp <> nil) then
          GlobalCEFApp.OsmodalLoop := True;

      WM_EXITMENULOOP :
        if (aMessage.wParam = 0) and
           (GlobalCEFApp <> nil) then
          GlobalCEFApp.OsmodalLoop := False;

      WM_MOVE,
      WM_MOVING : NotifyMoveOrResizeStarted;

      WM_SIZE :
        if (aMessage.wParam = SIZE_RESTORED) then
          UpdateCustomWindowState;

      WM_WINDOWPOSCHANGING :
        begin
          TempWindowPos := TWMWindowPosChanging(aMessage).WindowPos;
          if ((TempWindowPos.Flags and SWP_STATECHANGED) <> 0) then
            UpdateCustomWindowState;
        end;

      WM_SHOWWINDOW :
        if (aMessage.wParam <> 0) and (aMessage.lParam = SW_PARENTOPENING) then
          PostCustomMessage(CEF_SHOWBROWSER);

      CEF_DESTROY :
        if (FMXWindowParent <> nil) then
          FreeAndNil(FMXWindowParent);

      CEF_SHOWBROWSER :
        if (FMXWindowParent <> nil) then
          begin
            FMXWindowParent.WindowState := TWindowState.wsNormal;
            FMXWindowParent.Show;
            FMXWindowParent.SetBounds(GetFMXWindowParentRect);
          end;
    end;

    aMessage.Result := CallWindowProc(FOldWndPrc, FmxHandleToHWND(Handle), aMessage.Msg, aMessage.wParam, aMessage.lParam);
  except
    on e : exception do
      if CustomExceptionHandler('TForm9.CustomWndProc', e) then raise;
  end;
end;
{$ENDIF}

procedure TunitHome.FMXChromium1BeforeClose(Sender: TObject;
  const browser: ICefBrowser);
begin
  FCanClose := True;
  PostCustomMessage(WM_CLOSE);
end;

procedure TunitHome.FMXChromium1Close(Sender: TObject; const browser: ICefBrowser;
  var aAction: TCefCloseBrowserAction);
begin
  PostCustomMessage(CEF_DESTROY);
  aAction := cbaDelay;
end;

procedure TunitHome.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := FCanClose;

  if not(FClosing) then
    begin
      FClosing := True;
      Visible  := False;
      FMXChromium1.CloseBrowser(True);
    end;
end;

procedure TunitHome.FormCreate(Sender: TObject);
begin
  ReportMemoryLeaksOnShutdown := false;
     if frmLogin = nil then
    Application.CreateForm(TfrmLogin, frmLogin);

  try
    frmLogin.ShowModal;
  finally

  end;
end;

procedure TunitHome.FormResize(Sender: TObject);
begin
  ResizeChild;
end;

procedure TunitHome.FormShow(Sender: TObject);
var
  I: Integer;
begin
  CreateFMXWindowParent;

  fontFamily := 'Roboto';
end;

function TunitHome.GetFMXWindowParentRect: System.Types.TRect;
var
  TempScale : single;
begin
  TempScale     := FMXChromium1.ScreenScale;
  Result.Left   := 0;
  Result.Top    := 0;
  Result.Right  := round(ClientWidth  * TempScale) - 1;
  Result.Bottom := round((ClientHeight) * TempScale) - 1;
end;

procedure TunitHome.GraphOne(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: string;

  I, J, tamanhoQry, columnQry: Integer;
  arrayQry: array of TFDQuery;

  grupoROID, grupoRONome: String;

  rgbColor: string;

  usuariosID: String;

  userID: Integer;
  userNome: String;

  PesQntd, Fn1Qntd, Fn2Qntd, Fn3Qntd: integer;
  FnPesV1, FnPesV2, FnPesV3, FnCR1, FnCR2, FnCR3: integer;
  CaPes: integer;
  tempoTotalFicha, tempoTotalLig, tempoMedioFicha, tempoMedioLig, tempoTotalGeral: String;
  qntdLigacoes, qntdFichas: String;
  ativos, inativos, marcPeriodo: string;
  totalMensagens: String;

  qntdV, qntdT, qntdVT: String;
  clientesTrab: String;

  mudaCor: Boolean;
  montaTableOPRO, gridCards: String;

  headerCharts, bodyCharts, footerCharts: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts: iModelHTMLChartsConfig;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;


  if caCodigo = '' then
  Begin

    if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  navBarFunction := MontaNavBarHTML('GraphOne', dateStart, dateFinish, false, '', '', true, campanhasID);

  //MONTA A TABELA DE OPERADORES X ROs
  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;
  qryGeral1.SQL.Clear;
  qryGeral1.SQL.Add(
    ' SELECT UPPER(dbo.RetornaNomeUsuario(Usuarios_ID)) Nome, Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
    ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
    ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
    ' WHERE Usuarios_ID <> 13 AND ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' ' +
    ' AND ISNULL(CaEncerrada,0) = 0 AND ISNULL(UsAtivo,0) = 1 ' +
    ' GROUP BY Usuarios_ID ORDER BY Nome ');
  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;

  dmDados.GravaMsgArquivoAppLocal('StringNovas.txt', qryGeral1.SQL.Text, true);

  qryGeral2.Active := false;
  qryGeral2.SQL.Clear;
  qryGeral2.SQL.Add(' if exists(select * ' +
      ' from INFORMATION_SCHEMA.TABLES ' +
      ' where TABLE_NAME = ''Temp_Table_ContatosFichas'') ' +
      ' drop table Temp_Table_ContatosFichas ' +

      ' create table Temp_Table_ContatosFichas ' +
      ' (nomeRO varchar(100)');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', op' + userID.ToString + ' varchar(10) ');


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , rot varchar(10)) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , rot varchar(10)) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.ExecSQL;
  LimpaQuery(qryGeral2);

  qryGeral2.SQL.Add(
        ' declare @idResumo int ' +
        ' declare @nomeRO varchar(200) ' +
        ' declare cursor1 cursor for ' +

        ' SELECT distinct(CoResumoOperacaoID), UPPER(dbo.RetornaRONome(CoResumoOperacaoID)) ' +
        ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        ' INNER JOIN ResumoDeOperacoes WITH(NOLOCK) ON CoResumoOperacaoID = ResumoDeOperacoes_ID ' +
        ' WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' AND ' +
        ' ISNULL(CoResumoOperacaoID,0) <> 0 AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ISNULL(ReOcultarRelDash,0) = 0 ' +
        ' ORDER BY UPPER(dbo.RetornaRONome(CoResumoOperacaoID)) ' +

        ' open cursor1 ' +
        ' fetch next from cursor1 into @idResumo, @nomeRO ' +
        ' while @@FETCH_STATUS = 0 ' +
        ' Begin ');

  //INSERT DENTRO DO CURSOR
  qryGeral2.SQL.Add('INSERT INTO Temp_Table_ContatosFichas (nomeRO ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', op' + userID.ToString);


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , rot) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , rot) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.SQL.Add('VALUES ( dbo.RetornaRONome(@idResumo) ');

  usuariosID := '';

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoDeOperacoes WITH(NOLOCK) ON CoResumoOperacaoID = ResumoDeOperacoes_ID ' +
        '  WHERE ISNULL(CoUsuariosID,0) = ' + userID.ToString + ' AND ' +
        '  CoResumoOperacaoID = @idResumo AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''') ');

      usuariosID := usuariosID + userID.ToString;

      if I < tamanhoQry - 1 then
        usuariosID := usuariosID + ', ';

      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoDeOperacoes WITH(NOLOCK) ON CoResumoOperacaoID = ResumoDeOperacoes_ID ' +
        '  WHERE CoResumoOperacaoID = @idResumo AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' ' +
        '  AND ISNULL(CoUsuariosID,0) IN (' + usuariosID + ') ) ) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoDeOperacoes WITH(NOLOCK) ON CoResumoOperacaoID = ResumoDeOperacoes_ID ' +
        '  WHERE CoResumoOperacaoID = @idResumo AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''') ) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.SQL.Add(' fetch next from cursor1 into @idResumo, @nomeRO' +
        ' End ' +
        ' close cursor1 ' +
        ' deallocate cursor1 ');


  qryGeral2.ExecSQL;
  LimpaQuery(qryGeral2);


  //COLOCAR TOTAL NA ÚLTIMA LINHA
  qryGeral2.SQL.Add('INSERT INTO Temp_Table_ContatosFichas (nomeRO ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', op' + userID.ToString);


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , rot) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , rot) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.SQL.Add('SELECT ''TOTAL'' ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', SUM(CONVERT(INT, op' + userID.ToString + ')) ');

      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(', SUM(CONVERT(INT, rot)) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(', SUM(CONVERT(INT, rot)) ');

  End;

  qryGeral2.SQL.Add(' FROM Temp_Table_ContatosFichas ');

  userID := 0;
  qryGeral1.First;

  qryGeral2.ExecSQL;
  LimpaQuery(qryGeral2);

  //SELECT FINAL DA TABELA
  qryGeral2.SQL.Add('SELECT UPPER(nomeRO) [RESUMO DE OPERAÇÃO] ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
      userNome := qryGeral1.FieldByName('Nome').AsString;

      qryGeral2.SQL.Add(', CASE ISNULL(op' + userID.ToString + ','''') WHEN 0 ' +
        ' THEN '''' ELSE ISNULL(op' + userID.ToString + ','''') END [' + userNome + '] ');


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , CASE ISNULL(ROT,'''') WHEN 0 THEN '''' ELSE ISNULL(ROT,'''') END [TOTAL] ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , CASE ISNULL(ROT,'''') WHEN 0 THEN '''' ELSE ISNULL(ROT,'''') END [TOTAL] ');

  End;

  qryGeral2.SQL.Add(' FROM Temp_Table_ContatosFichas');

//  qryGeral2.ParamByName('dtInicial').Value := dateInicial;
//  qryGeral2.ParamByName('dtFinal').Value := dateFinal;

  try
    qryGeral2.Active := true;
  except
    on E: EFDDBEngineException do
    begin
         //dmdados.fdErrorDialog.Execute(E);
    end;

  end;

  dmDados.GravaMsgArquivoAppLocal('StringNovas.txt', ' QRY II - ' + qryGeral2.SQL.Text, false);

  tamanhoQry := qryGeral2.RecordCount;
  columnQry := qryGeral2.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral2.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral2.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral2.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral2.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';

  dmDados.GravaMsgArquivoAppLocal('StringNovas.txt', ' TABELA HTML - ' + montaTableOPRO, false);


  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  if tamanhoQry < 4 then
    gridCards := ' <div class="row row-cols-1 row-cols-md-' + tamanhoQry.ToString + ' g-4" style=" width: 80%; margin-left: 10%"> '
  else
    gridCards := ' <div class="row row-cols-1 row-cols-md-3 g-4" style=" width: 80%; margin-left: 10%"> ';


  qryGeral1.First;

  mudaCor := true;

  for I := 1 to tamanhoQry do
  Begin

    userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
    userNome := qryGeral1.FieldByName('Nome').AsString;

    //fichas virgens por op e data de cad
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(FN_StatusFicha) QNTD ' +
    ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE ISNULL(FN_StatusFicha,'''') = ''V'' AND CONVERT(DATE, FN_DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    qntdV := qryGeral2.FieldByName('QNTD').AsString;

    //fichas trabalhadas por op e data de cad
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(FN_StatusFicha) QNTD ' +
    ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE ISNULL(FN_StatusFicha,'''') = ''T'' AND CONVERT(DATE, FN_DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    qntdT := qryGeral2.FieldByName('QNTD').AsString;

    //fichas virgens por op desde sempre
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(FN_StatusFicha) QNTD ' +
    ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE ISNULL(FN_StatusFicha,'''') = ''V'' AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    qntdVT := qryGeral2.FieldByName('QNTD').AsString;


    //CLIENTES TRABALHADOS SEM REPETIR
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(DISTINCT CoFoneListsID) QNTD ' +
    ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' INNER JOIN ResumoDeOperacoes WITH(NOLOCK) ON ResumoDeOperacoes_ID = CoResumoOperacaoID ' +
    ' WHERE ISNULL(CoUsuariosID,0) = ' + userID.ToString + ' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ' +
    ' AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ISNULL(ReOcultarRelDash,0) = 0 ');

    qryGeral2.Active := true;

    clientesTrab := qryGeral2.FieldByName('QNTD').AsString;

    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(' SELECT COUNT(*) qntdLigacoes, ' +
       ' ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(CO_Duracao,''''))), ''00:00:00.000'')), ''00:00:00'') as ligTotal, ' +
       ' ISNULL(CONVERT(TIME(0), DATEADD(SECOND, AVG(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(CO_Duracao,''''))), ''00:00:00.000'')), ''00:00:00'') as ligAvg ' +
       ' FROM Contatos_' + caCodigo + ' WITH(NOLOCK) ' +
       ' INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CO_ContatosFichaID = ContatosFichas_' + caCodigo + '_ID ' +
       ' WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(CoUsuariosID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    tempoTotalLig := qryGeral2.FieldByName('ligTotal').AsString;
    tempoMedioLig := qryGeral2.FieldByName('ligAvg').AsString;
    qntdLigacoes := qryGeral2.FieldByName('qntdLigacoes').AsString;

    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(' SELECT COUNT(*) qntdFichas, ' +
    ' ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(CoDuracaoFicha,''''))), ''00:00:00'')), ''00:00:00'') fichaTotal, ' +
    ' CONVERT(TIME(0), DATEADD(SECOND, (CONVERT(FLOAT, SUM(DATEDIFF(SECOND,CoDataInicioFicha, CoDataTerminoFicha))) / COUNT(CoUsuariosID)), ''00:00:00.00'')) fichaAVG ' +
    ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' INNER JOIN ResumoDeOperacoes WITH(NOLOCK) ON ResumoDeOperacoes_ID = CoResumoOperacaoID ' +
    ' WHERE ISNULL(ReOcultarRelDash,0) = 0 AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
    ' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(CoUsuariosID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    tempoTotalFicha := qryGeral2.FieldByName('fichaTotal').AsString;
    tempoMedioFicha := qryGeral2.FieldByName('fichaAVG').AsString;
    qntdFichas := qryGeral2.FieldByName('qntdFichas').AsString;

    //PEGAR ATIVOS
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(FN_PessoasID) QNTD ' +
      ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
      ' WHERE ISNULL(FN_AtivoNaLista,0) = 1 AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString +
      ' AND CONVERT(DATE, FN_DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');
    qryGeral2.Active := true;

    ativos := qryGeral2.FieldByName('QNTD').AsString;

    //PEGAR INATIVOS
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(FN_PessoasID) QNTD ' +
      ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
      ' WHERE ISNULL(FN_AtivoNaLista,0) = 0 AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString +
      ' AND CONVERT(DATE, FN_DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');
    qryGeral2.Active := true;

    inativos := qryGeral2.FieldByName('QNTD').AsString;

    //PEGAR AGENDAMENTOS DO PERÍODO
//    LimpaQuery(qryGeral2);
//    qryGeral2.Active := false;
//    qryGeral2.SQL.Clear;
//    qryGeral2.SQL.Add(
//      ' SELECT COUNT(FN_PessoasID) QNTD ' +
//      ' FROM Agendas_' + caCodigo + ' WITH(NOLOCK) ' +
//      ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON AG_FN_FoneList_' + caCodigo + 'ID = FN_FoneList_' + caCodigo + '_ID ' +
//      ' WHERE ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString + ' AND ISNULL(AG_Status,'''') = ''ABERTO'' AND ' +
//      ' CONVERT(DATE, AG_DataProximaLigacao) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');
//    qryGeral2.Active := true;
//
//    marcPeriodo := qryGeral2.FieldByName('QNTD').AsString;
    /// FOI MUDADO PARA PEGAR AS AGENDAS ATRASADAS, AQUI EM BAIXO V
    ///
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(FN_PessoasID) QNTD ' +
      ' FROM Agendas_' + caCodigo + ' WITH(NOLOCK) ' +
      ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON AG_FN_FoneList_' + caCodigo + 'ID = FN_FoneList_' + caCodigo + '_ID ' +
      ' WHERE ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString + ' AND ISNULL(AG_Status,'''') = ''ABERTO'' AND ' +
      ' CONVERT(DATE, AG_DataProximaLigacao) < CONVERT(DATE, GETDATE()) ');
    qryGeral2.Active := true;

    marcPeriodo := qryGeral2.FieldByName('QNTD').AsString;


    //Mensagens Whats
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(WhatsMensagens_ID) QNTD ' +
      ' FROM WhatsContatosFichas WITH(NOLOCK) ' +
      ' INNER JOIN WhatsMensagens WITH(NOLOCK) ON WhContatosFichasID = WhatsContatosFichas_ID ' +
      ' WHERE ISNULL(WhCodigoCampanha,'''') = ' + caCodigo + ' AND ISNULL(WhUsuariosID,0) = ' + userID.ToString +
      ' AND ISNULL(WhOrigemMensagem,0) = 0 AND WhEnviadaPeloSistema = 1 AND CONVERT(DATE, WhDataHora) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');
    qryGeral2.Active := true;

    totalMensagens := qryGeral2.FieldByName('QNTD').AsString;

    if mudaCor then
    Begin

      gridCards := gridCards +
      ' <div class="col"> ' +
        ' <div class="card shadow p-3 mb-5 bg-body rounded" style=" background-color: #FFFFFF; margin-bottom: 10px; "> ' +
          ' <div class="card-header">' + userNome + '</div>' +
          ' <div class="card-body"> ' +
            ' <h5 class="card-title"> Total de Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdVT + '</span></h5> ' +
            ' <p class="card-text">Ativos: <span style=" display:block; float:right; margin-left:10px;">' + ativos + '</span></p> ' +
            ' <p class="card-text">Inativos: <span style=" display:block; float:right; margin-left:10px;">' + inativos + '</span></p> ' +
            ' <p class="card-text">Agendas em Atraso: <span style=" display:block; float:right; margin-left:10px;">' + marcPeriodo + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Chegada na Campanha</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-user"></i> Status das Fichas</h5> ' +
            ' <p class="card-text">Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdV + '</span></p> ' +
            ' <p class="card-text">Trabalhadas: <span style=" display:block; float:right; margin-left:10px;">' + qntdT + '</span></p> ' +
            ' <p class="card-text">Clientes Recebidos: <span style=" display:block; float:right; margin-left:10px;">' + IntToStr(StrToInt(qntdV) + StrToInt(qntdT)) + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Contatados no Período</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-phone-volume"></i> Ligações</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdLigacoes + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioLig + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalLig + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-comment"></i> Mensagens: <span style=" display:block; float:right; margin-left:10px;">' + totalMensagens + '</span></h5> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-file-alt"></i> Fichas</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdFichas + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioFicha + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalFicha + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"> Clientes Trabalhados: <span style=" display:block; float:right; margin-left:10px;">' + clientesTrab + '</span></h5> ' +
          ' </div> ' +
        ' </div> ' +
      ' </div> ';

      mudaCor := false;

    End
    else
    Begin

     gridCards := gridCards +
      ' <div class="col"> ' +
        ' <div class="card shadow p-3 mb-5 bg-body rounded" style=" background-color: #F2F2F2; margin-bottom: 10px; "> ' +
          ' <div class="card-header">' + userNome + '</div>' +
          ' <div class="card-body"> ' +
            ' <h5 class="card-title"> Total de Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdVT + '</span></h5> ' +
            ' <p class="card-text">Ativos: <span style=" display:block; float:right; margin-left:10px;">' + ativos + '</span></p> ' +
            ' <p class="card-text">Inativos: <span style=" display:block; float:right; margin-left:10px;">' + inativos + '</span></p> ' +
            ' <p class="card-text">Agendas em Atraso: <span style=" display:block; float:right; margin-left:10px;">' + marcPeriodo + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Chegada na Campanha</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-user"></i> Status das Fichas</h5> ' +
            ' <p class="card-text">Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdV + '</span></p> ' +
            ' <p class="card-text">Trabalhadas: <span style=" display:block; float:right; margin-left:10px;">' + qntdT + '</span></p> ' +
            ' <p class="card-text">Clientes Recebidos: <span style=" display:block; float:right; margin-left:10px;">' + IntToStr(StrToInt(qntdV) + StrToInt(qntdT)) + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Contatados no Período</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-phone-volume"></i> Ligações</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdLigacoes + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioLig + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalLig + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-comment"></i> Mensagens: <span style=" display:block; float:right; margin-left:10px;">' + totalMensagens + '</span></h5> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-file-alt"></i> Fichas</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdFichas + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioFicha + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalFicha + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"> Clientes Trabalhados: <span style=" display:block; float:right; margin-left:10px;">' + clientesTrab + '</span></h5> ' +
          ' </div> ' +
        ' </div> ' +
      ' </div> ';

      mudaCor := true;

    End;

    if (I mod 4 = 0) then
    Begin

      if mudaCor then
        mudaCor := false
      else
        mudaCor := true;

    End;

    qryGeral1.Next;

  End;

  gridCards := gridCards + ' </div> ';

  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;
  qryGeral1.SQL.Add(
    ' WITH pessoas_1 AS (  ' +
    '  ' +
    ' SELECT DBO.RetornaNomeUsuario(CoUsuariosID) OP,  ' +
    ' 	COUNT(DISTINCT CoFoneListsID) COFN,  ' +
    ' 	CAST(CAST(COUNT(DISTINCT CoFoneListsID) AS decimal(18,2)) AS decimal(18,2)) / CAST((SELECT COUNT(DISTINCT CoFoneListsID)  ' +
    ' 		FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' 		WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ' AND  ' +
    '       CoResumoOperacaoID IN (SELECT CoResumoOperacaoID FROM ResumoDeOperacoes WHERE ISNULL(ReOcultarRelDash,0) = 0) AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ' +
    ' 			CoUsuariosID IN (SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S''))  AS DECIMAL(18,2)) AS TOTAL,  ' +
    ' 	''RGB'' AS RGB, ROW_NUMBER() OVER (ORDER BY COUNT(*) desc) AS classificacao  ' +
    ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = CoUsuariosID ' +
	  ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
    ' WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ' AND  ' +
    ' 	CoUsuariosID IN (SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S'') ' +
    '   AND CoResumoOperacaoID IN (SELECT CoResumoOperacaoID FROM ResumoDeOperacoes WHERE ISNULL(ReOcultarRelDash,0) = 0) AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
    ' AND CaCodigo = ' + QuotedStr(caCodigo) +
    ' GROUP BY CoUsuariosID),  ' +
    '  ' +
    ' pessoas_2 AS (  ' +
    '  ' +
    ' SELECT OP AS Label,	 ' +
    ' TOTAL as Value, ' +
    ' 	RGB, classificacao   ' +
    ' FROM pessoas_1 )  ' +
    '  ' +
    '          SELECT Label, Value, GrRGB AS RGB  ' +
    '          FROM pessoas_2 AS p  ' +
    '          JOIN GraficosRGB AS g ON p.classificacao = g.GraficosRGB_ID  ' +
    '          ORDER BY p.Value DESC ');

  qryGeral1.Active := true;

  dmDados.GravaMsgArquivoAppLocal('StringNovas.txt', ' PIE - ' + qryGeral1.SQL.Text, false);

  qryGrRO.Active := false;
  qryGrRO.Active := true;

  headerCharts := WebCharts1
          .ContinuosProject;

  config_row_1_Charts := headerCharts
            .Charts
              ._ChartType(horizontalBar)
                .Attributes
                  .Name('BAR_GrupoRO_OP')
                  .ColSpan(12)
                  .Heigth(180)
                  .Labelling
                    .Format('0')
                    .RGBColor('30,30,30')
                    .FontSize(20)
                    .FontStyle('normal')
                    .FontFamily(fontFamily)
                    .Padding(-10)
                    .PaddingX(10)
                  .&End
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                    .Tooltip
                      .Format('0')
                      .DisplayTitle(true)
                    .&End
                  .&End;


  tamanhoQry := qryGrRO.RecordCount;

  SetLength(arrayQry, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

    grupoROID := qryGrRO.FieldByName('ID').AsString;
    grupoRONome := qryGrRO.FieldByName('NOME').AsString;

     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(
        ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_GrupoRO_OP' + grupoROID + ''') ' +
        ' 	DROP TABLE TempDash_GrupoRO_OP' + grupoROID + '  ' +
        '  ' +
        '  CREATE TABLE TempDash_GrupoRO_OP' + grupoROID + '  ' +
        '     (id INT IDENTITY(1,1) NOT NULL, UsuariosID VARCHAR(10), UsNome VARCHAR(100), countRO VARCHAR(10), RGB VARCHAR(15)) ' +
        '  ' +
        ' INSERT INTO TempDash_GrupoRO_OP' + grupoROID + '(UsuariosID, UsNome) ' +
        ' SELECT Usuarios_ID, UsNome ' +
        ' FROM Usuarios WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
        ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
        ' WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(UsAtivo,0) = 1 AND CaCodigo = ' + caCodigo + ' ' +
        ' ORDER BY UsNome ' +
        '  ' +
        ' SELECT UsuariosID ' +
        ' FROM TempDash_GrupoRO_OP' + grupoROID);

       arrayQry[I].Active := True;

       for J := 0 to arrayQry[I].RecordCount - 1 do
       Begin

        usuariosID := arrayQry[I].FieldByName('UsuariosID').AsString;

        dmDados.qryGeral.ExecSQL(
          ' UPDATE  TempDash_GrupoRO_OP' + grupoROID +
          ' SET countRO = (' +
          ' SELECT COUNT(CoResumoOperacaoID) Value ' +
          ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
          ' INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = CoResumoOperacaoID  ' +
          ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = CoUsuariosID  ' +
          ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID  ' +
          ' WHERE ReGrupoResumoOperacoesID IN (' + QuotedStr(grupoROID) + ') AND  ' +
          ' CONVERT(DATE, CoDataInicioFicha) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) +
          ' AND CoUsuariosID IN (' + usuariosID + ') AND CaCodigo = ' + QuotedStr(caCodigo) + ' AND  ' +
          ' ISNULL(CoCampanhaOrigemID,0) = 0 AND ReResumoOperacoesID IN (SELECT ResumoDeOperacoes_ID FROM ResumoDeOperacoes WHERE ISNULL(ReOcultarRelDash,0) = 0) ) ' +
          ' WHERE UsuariosID = ' + usuariosID);

        arrayQry[I].Next;
       End;

       LimpaQuery(arrayQry[I]);
       arrayQry[I].Open('SELECT DBO.RetornaNomeUsuario(UsuariosID) Label, countRO Value, '''' RGB ' +
        ' FROM TempDash_GrupoRO_OP' + grupoROID);

       rgbColor := IntToStr(Random(256)) + ',' + IntToStr(Random(256)) + ',' + IntToStr(Random(256));

       if arrayQry[I].RecordCount > 0 then
       Begin

        config_row_1_Charts
              .DataSet
                .textLabel(grupoRONome)
                .DataSet(arrayQry[I])
                .BackgroundColor(rgbColor)
                .BackgroundOpacity(8)
              .&End;

       End;

     finally

     end;

     qryGrRO.Next;

  End;

  config_row_1_Charts.&End.&End;


  // MONTAGEM DOS GRÁFICOS
  WebCharts1
   .CDN(true)
   .BackgroundColor('#FFFFFF')
   .FontColor('#000000')
   .Container(fluid)
   .AddResource('<link href="css/green.css" rel="stylesheet">')
   .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
   .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
   .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
   .AddResource('<script src="/scripts/snippet-javascript-console.min.js?v=1"></script>')

     .NewProject

     .Rows
      .HTML(navBarFunction)
     .&End

       .Jumpline

       .Rows
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-2" style="font-family: Product Sans">Resumo de Operações</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
       .&End

       .Jumpline

       .Rows
        .ID('infoDate')
        .HTML(
          ' <figure class="text-center"> ' +
            ' <blockquote class="blockquote"> ' +
              ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
            ' </blockquote> ' +
            ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
              ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
            ' </figcaption> ' +
          ' </figure> '
        )
       .&End

       .Jumpline

      .Rows
        ._Div
          .ColSpan(12)
          .Add(
            WebCharts1
            .ContinuosProject
              .Charts
                ._ChartType(pie)
                  .Attributes
                    .Name('PIEROsTotal')
                    .Labelling
                      .Format('0.00%')// Consultar em http://numeraljs.com/#format
                      .RGBColor('30,30,30') // Cor RGB separado por Virgula
                      .FontSize(20) // Tamanho da Fonte
                      .FontStyle('normal') // normal, bold, italic
                      .FontFamily(fontFamily) // Open Sans, Arial, Helvetica e etc..
                      .Padding(0) // Numeros negativos e positivos
                      .PaddingX(0)
                    .&End
                    .ColSpan(12)
                    .Heigth(130)
                    .Options
                      .Legend
                        .display(true)
                        .position('bottom')
                        .Labels
                          .padding(30)
                          .fontSize(16)
                          .fontColorHEX('#000000')
                          .fontFamily(fontFamily)
                        .&End
                      .&End
                      .Title
                        .FontSize(18)
                        .fontColorHEX('#000000')
                        .text('TOTAL DE CLIENTES TRABALHADOS')
                        .display(true)
                        .position('top')
                        .fontFamily(fontFamily)
                        .fontStyle('normal')
                        .padding(20)
                      .&End
                      .Tooltip
                        .DisplayTitle(true)
                        .Format('0.00%')
                        .ToolTipNoScales
                        .Intersect(true)
                      .&End
                    .&End
                    .DataSet
                      .textLabel('R.O.s Total por Período %')
                      .DataSet(qryGeral1)
                      .BackgroundOpacity(8)
                    .&End
                  .&End
                .&End
              .&End
            .HTML
            )
        .&End
      .&End

      .Jumpline

      .Rows
        ._Div
          .ColSpan(1)
        .&End
        ._Div
          .ColSpan(10)
          .Add(
              headerCharts.HTML)
        .&End
        ._Div
          .ColSpan(1)
        .&End
      .&End

      .Jumpline
      .Jumpline

      .Rows
        .ID('ROTable')
        .HTML(montaTableOPRO)
      .&End

      .Jumpline

      .Rows
        .ID('CardsRO')
        .HTML(gridCards)
      .&End


    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .CallbackJS
      .ClassProvider(Self)
    .&End
    .Generated;


  for i := 0 to tamanhoQry - 1 do
  begin
    FreeAndNil(arrayQry[i]);
  end;

  form_esmaecer_fundo.Hide;

end;

procedure TunitHome.AttGraphOne(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  I, J, tamanhoQry, columnQry: Integer;
  arrayQry: array of TFDQuery;

  grupoROID, grupoRONome: String;

  userID: Integer;
  userNome: String;
  rgbColor: string;

  PesQntd, Fn1Qntd, Fn2Qntd, Fn3Qntd: integer;
  FnPesV1, FnPesV2, FnPesV3, FnCR1, FnCR2, FnCR3: integer;
  CaPes: integer;
  tempoTotalFicha, tempoTotalLig, tempoMedioFicha, tempoMedioLig, tempoTotalGeral: String;
  qntdLigacoes, qntdFichas: String;
  ativos, inativos, marcPeriodo: string;
  totalMensagens: String;

  qntdV, qntdT, qntdVT: String;
  clientesTrab: String;

  mudaCor: Boolean;
  montaTableOPRO, gridCards: String;
  headerCharts, bodyCharts, footerCharts: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts: iModelHTMLChartsConfig;
begin

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

     LimpaQuery(qryCamp);
     qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
     ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');

     caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
     campanhaCodigo := caCodigo;
     campanhaNome := qryCamp.FieldByName('CaNome').AsString;
     campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;


  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
      .ID('infoDate')
      .HTML(
          ' <figure class="text-center"> ' +
            ' <blockquote class="blockquote"> ' +
              ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
            ' </blockquote> ' +
            ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
              ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
            ' </figcaption> ' +
          ' </figure> '
        )
      .Update;

  //FAZ A NOVA PIE - QUANTIDADE DE COFNLISTID POR OP NA CONTATOFICHAS
  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;
  qryGeral1.SQL.Add(
    ' WITH pessoas_1 AS (  ' +
    '  ' +
    ' SELECT DBO.RetornaNomeUsuario(CoUsuariosID) OP,  ' +
    ' 	COUNT(DISTINCT CoFoneListsID) COFN,  ' +
    ' 	CAST(CAST(COUNT(DISTINCT CoFoneListsID) AS decimal(18,2)) AS decimal(18,2)) / CAST((SELECT COUNT(DISTINCT CoFoneListsID)  ' +
    ' 		FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' 		WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ' AND  ' +
    ' 			CoUsuariosID IN (SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S''))  AS DECIMAL(18,2)) AS TOTAL,  ' +
    ' 	''RGB'' AS RGB, ROW_NUMBER() OVER (ORDER BY COUNT(*) desc) AS classificacao  ' +
    ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ' AND  ' +
    ' 	CoUsuariosID IN (SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S'') ' +
    ' GROUP BY CoUsuariosID),  ' +
    '  ' +
    ' pessoas_2 AS (  ' +
    '  ' +
    ' SELECT OP AS Label,	 ' +
    ' TOTAL as Value, ' +
    ' 	RGB, classificacao   ' +
    ' FROM pessoas_1 )  ' +
    '  ' +
    '          SELECT Label, Value, GrRGB AS RGB  ' +
    '          FROM pessoas_2 AS p  ' +
    '          JOIN GraficosRGB AS g ON p.classificacao = g.GraficosRGB_ID  ' +
    '          ORDER BY p.Value DESC ');

  qryGeral1.Active := true;

    //ATUALIZA A PIE
    WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .Charts
      ._ChartType(pie)
        .Attributes
          .Name('PIEROsTotal')
          .Labelling
            .Format('0.00%')// Consultar em http://numeraljs.com/#format
            .RGBColor('30,30,30') // Cor RGB separado por Virgula
            .FontSize(20) // Tamanho da Fonte
            .FontStyle('normal') // normal, bold, italic
            .FontFamily(fontFamily) // Open Sans, Arial, Helvetica e etc..
            .Padding(0) // Numeros negativos e positivos
            .PaddingX(0)
          .&End
          .ColSpan(12)
          .Heigth(150)
          .Options
            .Legend
              .display(true)
              .position('bottom')
              .Labels
                .padding(30)
                .fontSize(16)
                .fontColorHEX('#000000')
                .fontFamily(fontFamily)
              .&End
            .&End
            .Title
              .FontSize(18)
              .fontColorHEX('#000000')
              .text('TOTAL DE CLIENTES TRABALHADOS')
              .display(true)
              .position('top')
              .fontFamily(fontFamily)
              .fontStyle('normal')
              .padding(20)
            .&End
            .Tooltip
              .DisplayTitle(true)
              .Format('0.00%')
              .ToolTipNoScales
              .Intersect(true)
            .&End
          .&End
          .DataSet
            .textLabel('TOTAL DE CLIENTES TRABALHADOS')
            .DataSet(qryGeral1)
            .BackgroundOpacity(8)
          .&End
        .&End
      .UpdateChart;


  //COMEÇA A MONTAR PARA ATUALIZAR O DE BARRAS HORIZONTAL
  headerCharts := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts
            .Charts
              ._ChartType(horizontalBar)
                .Attributes
                  .Name('BAR_GrupoRO_OP')
                  .ColSpan(12)
                  .Heigth(110)
                  .Labelling
                    .Format('0')
                    .RGBColor('30,30,30')
                    .FontSize(20)
                    .FontStyle('normal')
                    .FontFamily(fontFamily)
                  .&End
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                    .Tooltip
                      .Format('0')
                      .DisplayTitle(true)
                    .&End
                  .&End;

  qryGrRO.Active := false;
  qryGrRO.Active := true;

  tamanhoQry := qryGrRO.RecordCount;

  SetLength(arrayQry, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

    grupoROID := qryGrRO.FieldByName('ID').AsString;
    grupoRONome := qryGrRO.FieldByName('NOME').AsString;


     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(
        ' SELECT DBO.RetornaNomeUsuario(CoUsuariosID) Label, COUNT(CoResumoOperacaoID) Value, '''' RGB ' +
        ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        ' INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = CoResumoOperacaoID ' +
        ' WHERE ReGrupoResumoOperacoesID = ' + QuotedStr(grupoROID) + ' AND ' +
        ' CONVERT(DATE, CoDataInicioFicha) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) +
        ' GROUP BY CoUsuariosID ' +
        ' ORDER BY Label ');


       arrayQry[I].Active := True;

       rgbColor := IntToStr(Random(256)) + ',' + IntToStr(Random(256)) + ',' + IntToStr(Random(256));

      config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(grupoRONome)
                .BackgroundColor(rgbColor)
                .BackgroundOpacity(8)
              .&End;


     finally


     end;


     qryGrRO.Next;

  End;

  config_row_1_Charts.&End
              .UpdateChart;

  for i := 0 to tamanhoQry - 1 do
  begin
    FreeAndNil(arrayQry[i]);
  end;

  //MONTA A TABELA DE OPERADORES X ROs
  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;
  qryGeral1.SQL.Clear;
  qryGeral1.SQL.Add(' SELECT UPPER(dbo.RetornaNomeUsuario(Usuarios_ID)) Nome, Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
            ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
            ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
            ' WHERE Usuarios_ID <> 13 AND ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
            ' AND ISNULL(UsAtivo,0) = 1 ' +
            ' GROUP BY Usuarios_ID ORDER BY Nome ');
  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;


  qryGeral2.Active := false;
  qryGeral2.SQL.Clear;
  qryGeral2.SQL.Add(' if exists(select * ' +
      ' from INFORMATION_SCHEMA.TABLES ' +
      ' where TABLE_NAME = ''Temp_Table_ContatosFichas'') ' +
      ' drop table Temp_Table_ContatosFichas ' +

      ' create table Temp_Table_ContatosFichas ' +
      ' (nomeRO varchar(100)');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', op' + userID.ToString + ' varchar(10) ');


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , rot varchar(10)) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , rot varchar(10)) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.ExecSQL;

  qryGeral2.SQL.Add(
        ' declare @idResumo int ' +
        ' declare @nomeRO varchar(200) ' +
        ' declare cursor1 cursor for ' +

        ' SELECT distinct(CoResumoOperacaoID), UPPER(dbo.RetornaRONome(CoResumoOperacaoID)) ' +
        ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        ' WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN :dtInicial AND :dtFinal AND ' +
        ' ISNULL(CoResumoOperacaoID,0) <> 0 AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
        ' ORDER BY UPPER(dbo.RetornaRONome(CoResumoOperacaoID)) ' +

        ' open cursor1 ' +
        ' fetch next from cursor1 into @idResumo, @nomeRO ' +
        ' while @@FETCH_STATUS = 0 ' +
        ' Begin ');

  //INSERT DENTRO DO CURSOR
  qryGeral2.SQL.Add('INSERT INTO Temp_Table_ContatosFichas (nomeRO ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', op' + userID.ToString);


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , rot) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , rot) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.SQL.Add('VALUES ( dbo.RetornaRONome(@idResumo) ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  WHERE ISNULL(CoUsuariosID,0) = ' + userID.ToString + ' AND ' +
        '  CoResumoOperacaoID = @idResumo AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN :dtInicial AND :dtFinal) ');

      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  WHERE CoResumoOperacaoID = @idResumo AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN :dtInicial AND :dtFinal) ) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  WHERE CoResumoOperacaoID = @idResumo AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN :dtInicial AND :dtFinal) ) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.SQL.Add(' fetch next from cursor1 into @idResumo, @nomeRO' +
        ' End ' +
        ' close cursor1 ' +
        ' deallocate cursor1 ');


  //COLOCAR TOTAL NA ÚLTIMA LINHA
  qryGeral2.SQL.Add('INSERT INTO Temp_Table_ContatosFichas (nomeRO ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', op' + userID.ToString);


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , rot) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , rot) ');

  End;

  userID := 0;
  qryGeral1.First;

  qryGeral2.SQL.Add('VALUES ( ''TOTAL'' ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;

      qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  WHERE ISNULL(CoResumoOperacaoID,0) <> 0 AND ' +
        '  ISNULL(CoUsuariosID,0) = ' + userID.ToString + ' AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN :dtInicial AND :dtFinal) ');

      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  WHERE ISNULL(CoResumoOperacaoID,0) <> 0 AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN :dtInicial AND :dtFinal) ) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '  WHERE ISNULL(CoResumoOperacaoID,0) <> 0 AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ' +
        '  CONVERT(DATE, CoDataInicioFicha) BETWEEN :dtInicial AND :dtFinal) ) ');

  End;

  userID := 0;
  qryGeral1.First;

  //SELECT FINAL DA TABELA
  qryGeral2.SQL.Add('SELECT UPPER(nomeRO) [RESUMO DE OPERAÇÃO] ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
      userNome := qryGeral1.FieldByName('Nome').AsString;

      qryGeral2.SQL.Add(', CASE ISNULL(op' + userID.ToString + ','''') WHEN 0 ' +
        ' THEN '''' ELSE ISNULL(op' + userID.ToString + ','''') END [' + userNome + '] ');


      qryGeral1.Next;

    End;

    qryGeral2.SQL.Add(' , CASE ISNULL(ROT,'''') WHEN 0 THEN '''' ELSE ISNULL(ROT,'''') END [TOTAL] ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , CASE ISNULL(ROT,'''') WHEN 0 THEN '''' ELSE ISNULL(ROT,'''') END [TOTAL] ');

  End;

  qryGeral2.SQL.Add(' FROM Temp_Table_ContatosFichas');

  qryGeral2.ParamByName('dtInicial').Value := dateInicial;
  qryGeral2.ParamByName('dtFinal').Value := dateFinal;

  try
    qryGeral2.Active := true;
  except
    on E: EFDDBEngineException do
    begin
         //dmdados.fdErrorDialog.Execute(E);
    end;

  end;

  tamanhoQry := qryGeral2.RecordCount;
  columnQry := qryGeral2.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral2.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral2.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral2.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral2.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';

  //UPDATE TABLE NO DASH
  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
      .ID('ROTable')
      .HTML(montaTableOPRO)
      .Update;


  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  if tamanhoQry < 4 then
    gridCards := ' <div class="row row-cols-1 row-cols-md-' + tamanhoQry.ToString + ' g-4" style=" width: 80%; margin-left: 10%"> '
  else
    gridCards := ' <div class="row row-cols-1 row-cols-md-3 g-4" style=" width: 80%; margin-left: 10%"> ';


  qryGeral1.First;

  mudaCor := true;

  for I := 1 to tamanhoQry do
  Begin

    userID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
    userNome := qryGeral1.FieldByName('Nome').AsString;

    //fichas virgens por op e data de cad
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(FN_StatusFicha) QNTD ' +
    ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE ISNULL(FN_StatusFicha,'''') = ''V'' AND CONVERT(DATE, FN_DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    qntdV := qryGeral2.FieldByName('QNTD').AsString;

    //fichas trabalhadas por op e data de cad
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(FN_StatusFicha) QNTD ' +
    ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE ISNULL(FN_StatusFicha,'''') = ''T'' AND CONVERT(DATE, FN_DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    qntdT := qryGeral2.FieldByName('QNTD').AsString;

    //fichas virgens por op desde sempre
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(FN_StatusFicha) QNTD ' +
    ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE ISNULL(FN_StatusFicha,'''') = ''V'' AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    qntdVT := qryGeral2.FieldByName('QNTD').AsString;


    //CLIENTES TRABALHADOS SEM REPETIR
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
    ' SELECT COUNT(DISTINCT CoFoneListsID) QNTD ' +
    ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE ISNULL(CoUsuariosID,0) = ' + userID.ToString + ' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');

    qryGeral2.Active := true;

    clientesTrab := qryGeral2.FieldByName('QNTD').AsString;

    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(' SELECT COUNT(*) qntdLigacoes, ' +
       ' ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(CO_Duracao,''''))), ''00:00:00.000'')), ''00:00:00'') as ligTotal, ' +
       ' ISNULL(CONVERT(TIME(0), DATEADD(SECOND, AVG(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(CO_Duracao,''''))), ''00:00:00.000'')), ''00:00:00'') as ligAvg ' +
       ' FROM Contatos_' + caCodigo + ' WITH(NOLOCK) ' +
       ' INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CO_ContatosFichaID = ContatosFichas_' + caCodigo + '_ID ' +
       ' WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(CoUsuariosID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    tempoTotalLig := qryGeral2.FieldByName('ligTotal').AsString;
    tempoMedioLig := qryGeral2.FieldByName('ligAvg').AsString;
    qntdLigacoes := qryGeral2.FieldByName('qntdLigacoes').AsString;

    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(' SELECT COUNT(*) qntdFichas, ' +
    ' ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(CoDuracaoFicha,''''))), ''00:00:00'')), ''00:00:00'') fichaTotal, ' +
    ' ISNULL(CONVERT(TIME(0), DATEADD(SECOND, AVG(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(CoDuracaoFicha,''''))), ''00:00:00'')), ''00:00:00'') fichaAVG ' +
    ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
    ' WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ISNULL(CoUsuariosID,0) = ' + userID.ToString);

    qryGeral2.Active := true;

    tempoTotalFicha := qryGeral2.FieldByName('fichaTotal').AsString;
    tempoMedioFicha := qryGeral2.FieldByName('fichaAVG').AsString;
    qntdFichas := qryGeral2.FieldByName('qntdFichas').AsString;

    //PEGAR ATIVOS
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(FN_PessoasID) QNTD ' +
      ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
      ' WHERE ISNULL(FN_AtivoNaLista,0) = 1 AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);
    qryGeral2.Active := true;

    ativos := qryGeral2.FieldByName('QNTD').AsString;

    //PEGAR INATIVOS
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(FN_PessoasID) QNTD ' +
      ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
      ' WHERE ISNULL(FN_AtivoNaLista,0) = 0 AND ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString);
    qryGeral2.Active := true;

    inativos := qryGeral2.FieldByName('QNTD').AsString;

    //PEGAR AGENDAMENTOS DO PERÍODO
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(FN_PessoasID) QNTD ' +
      ' FROM Agendas_' + caCodigo + ' WITH(NOLOCK) ' +
      ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON AG_FN_FoneList_' + caCodigo + 'ID = FN_FoneList_' + caCodigo + '_ID ' +
      ' WHERE ISNULL(FN_OperadorProprietarioID,0) = ' + userID.ToString + ' AND ISNULL(AG_Status,'''') = ''ABERTO'' AND ' +
      ' CONVERT(DATE, AG_DataProximaLigacao) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');
    qryGeral2.Active := true;

    marcPeriodo := qryGeral2.FieldByName('QNTD').AsString;

    //PEGAR AGENDAMENTOS DO PERÍODO
    LimpaQuery(qryGeral2);
    qryGeral2.Active := false;
    qryGeral2.SQL.Clear;
    qryGeral2.SQL.Add(
      ' SELECT COUNT(WhatsMensagens_ID) QNTD ' +
      ' FROM WhatsContatosFichas WITH(NOLOCK) ' +
      ' INNER JOIN WhatsMensagens WITH(NOLOCK) ON WhContatosFichasID = WhatsContatosFichas_ID ' +
      ' WHERE ISNULL(WhCodigoCampanha,'''') = ' + caCodigo + ' AND ISNULL(WhUsuariosID,0) = ' + userID.ToString +
      ' AND ISNULL(WhOrigemMensagem,0) = 0 AND CONVERT(DATE, WhDataHora) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');
    qryGeral2.Active := true;

    totalMensagens := qryGeral2.FieldByName('QNTD').AsString;

    if mudaCor then
    Begin

      gridCards := gridCards +
      ' <div class="col"> ' +
        ' <div class="card shadow p-3 mb-5 bg-body rounded" style=" background-color: #FFFFFF; margin-bottom: 10px; "> ' +
          ' <div class="card-header">' + userNome + '</div>' +
          ' <div class="card-body"> ' +
            ' <h5 class="card-title"> Total de Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdVT + '</span></h5> ' +
            ' <p class="card-text">Ativos: <span style=" display:block; float:right; margin-left:10px;">' + ativos + '</span></p> ' +
            ' <p class="card-text">Inativos: <span style=" display:block; float:right; margin-left:10px;">' + inativos + '</span></p> ' +
            ' <p class="card-text">Marcados no Período: <span style=" display:block; float:right; margin-left:10px;">' + marcPeriodo + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Chegada na Campanha</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-user"></i> Status das Fichas</h5> ' +
            ' <p class="card-text">Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdV + '</span></p> ' +
            ' <p class="card-text">Trabalhadas: <span style=" display:block; float:right; margin-left:10px;">' + qntdT + '</span></p> ' +
            ' <p class="card-text">Clientes Recebidos: <span style=" display:block; float:right; margin-left:10px;">' + IntToStr(StrToInt(qntdV) + StrToInt(qntdT)) + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Contatados no Período</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-phone-volume"></i> Ligações</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdLigacoes + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioLig + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalLig + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-comment"></i> Mensagens: <span style=" display:block; float:right; margin-left:10px;">' + totalMensagens + '</span></h5> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-file-alt"></i> Fichas</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdFichas + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioFicha + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalFicha + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"> Clientes Trabalhados: <span style=" display:block; float:right; margin-left:10px;">' + clientesTrab + '</span></h5> ' +
          ' </div> ' +
        ' </div> ' +
      ' </div> ';

      mudaCor := false;

    End
    else
    Begin

     gridCards := gridCards +
      ' <div class="col"> ' +
        ' <div class="card shadow p-3 mb-5 bg-body rounded" style=" background-color: #F2F2F2; margin-bottom: 10px; "> ' +
          ' <div class="card-header">' + userNome + '</div>' +
          ' <div class="card-body"> ' +
            ' <h5 class="card-title"> Total de Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdVT + '</span></h5> ' +
            ' <p class="card-text">Ativos: <span style=" display:block; float:right; margin-left:10px;">' + ativos + '</span></p> ' +
            ' <p class="card-text">Inativos: <span style=" display:block; float:right; margin-left:10px;">' + inativos + '</span></p> ' +
            ' <p class="card-text">Marcados no Período: <span style=" display:block; float:right; margin-left:10px;">' + marcPeriodo + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Chegada na Campanha</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-user"></i> Status das Fichas</h5> ' +
            ' <p class="card-text">Virgens: <span style=" display:block; float:right; margin-left:10px;">' + qntdV + '</span></p> ' +
            ' <p class="card-text">Trabalhadas: <span style=" display:block; float:right; margin-left:10px;">' + qntdT + '</span></p> ' +
            ' <p class="card-text">Clientes Recebidos: <span style=" display:block; float:right; margin-left:10px;">' + IntToStr(StrToInt(qntdV) + StrToInt(qntdT)) + '</span></p> ' +
            ' <hr> ' +
            ' <h4 class="card-title">Contatados no Período</h4> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-phone-volume"></i> Ligações</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdLigacoes + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioLig + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalLig + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-comment"></i> Mensagens: <span style=" display:block; float:right; margin-left:10px;">' + totalMensagens + '</span></h5> ' +
            ' <hr> ' +
            ' <h5 class="card-title"><i class="fas fa-file-alt"></i> Fichas</h5> ' +
            ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">' + qntdFichas + '</span></p> ' +
            ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' + tempoMedioFicha + '</span></p> ' +
            ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' + tempoTotalFicha + '</span></p> ' +
            ' <hr> ' +
            ' <h5 class="card-title"> Clientes Trabalhados: <span style=" display:block; float:right; margin-left:10px;">' + clientesTrab + '</span></h5> ' +
          ' </div> ' +
        ' </div> ' +
      ' </div> ';

      mudaCor := true;

    End;

    if (I mod 4 = 0) then
    Begin

      if mudaCor then
        mudaCor := false
      else
        mudaCor := true;

    End;

    qryGeral1.Next;

  End;

  gridCards := gridCards + ' </div> ';

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
      .ID('CardsRO')
      .HTML(gridCards)
      .Update;

end;

procedure TunitHome.GraphTwo(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  I, J, tamanhoQry, columnQry: Integer;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  dtMYI, dtMYII: String;
  montaInfo: String;

  monthIHTML, monthIIHTML: String;

  navBarFunction: String;

  montaTableOPRO: String;

  roID, resumoID, resumoNome: String;

  usuarioID: Integer;
  usuarioNome, usuarioRGB: String;

  stringT: String;

  arrayQry, arrayQry2: array of TFDQuery;

  bodyCharts, footerCharts: iModelHTML;
  headerCharts_1, headerCharts_2, headerCharts_3, headerCharts_4: iModelHTML;
  headerCharts_5, headerCharts_6, headerCharts_7, headerCharts_8: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts, config_row_4_Charts: iModelHTMLChartsConfig;
  config_row_5_Charts, config_row_6_Charts, config_row_7_Charts, config_row_8_Charts: iModelHTMLChartsConfig;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;



  if caCodigo = '' then
  Begin

     if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

     caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
     campanhaCodigo := caCodigo;
     campanhaNome := qryCamp.FieldByName('CaNome').AsString;
     campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  if (monthStart <> '') and (monthFinish <> '') then
  Begin

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);

    anoInicial := monthStart.Substring(0,4);
    anoFinal := monthFinish.Substring(0,4);

    monthIHTML := anoInicial + '-' + mesInicial;
    monthIIHTML := anoFinal + '-' + mesFinal;

    dtMYI := '01/' + mesInicial + '/' + anoInicial;
    dtMYII := '01/' + mesFinal + '/' + anoFinal;

    //RETORNA NOME DO MÊS INICIAL
    Case StrToInt(mesInicial) Of
      01 : mesInicial := 'Janeiro';
      02 : mesInicial := 'Fevereiro';
      03 : mesInicial := 'Março';
      04 : mesInicial := 'Abril';
      05 : mesInicial := 'Maio';
      06 : mesInicial := 'Junho';
      07 : mesInicial := 'Julho';
      08 : mesInicial := 'Agosto';
      09 : mesInicial := 'Setembro';
      10 : mesInicial := 'Outubro';
      11 : mesInicial := 'Novembro';
      12 : mesInicial := 'Dezembro';
    end;

      //RETORNA NOME DO MÊS FINAL
    Case StrToInt(mesFinal) Of
      01 : mesFinal := 'Janeiro';
      02 : mesFinal := 'Fevereiro';
      03 : mesFinal := 'Março';
      04 : mesFinal := 'Abril';
      05 : mesFinal := 'Maio';
      06 : mesFinal := 'Junho';
      07 : mesFinal := 'Julho';
      08 : mesFinal := 'Agosto';
      09 : mesFinal := 'Setembro';
      10 : mesFinal := 'Outubro';
      11 : mesFinal := 'Novembro';
      12 : mesFinal := 'Dezembro';
    end;

    if anoInicial = anoFinal then
    Begin

      if mesInicial = mesFinal then
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' de ' + anoInicial + '.</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End
      else
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' até ' + mesFinal + ' de ' + anoInicial + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End;

    End
    else
    Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + '/' + anoInicial + '  até ' + mesFinal + '/' + anoFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

    End;

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);
  End
  Else
  Begin

    mesInicial := DateToStr(Now).Substring(3,2);
    anoInicial := DateToStr(Now).Substring(6,4);
    mesFinal := DateToStr(Now).Substring(3,2);
    anoFinal := DateToStr(Now).Substring(6,4);

    monthIHTML := anoInicial + '-' + mesInicial;
    monthIIHTML := anoFinal + '-' + mesFinal;

    dtMYI := '01/' + mesInicial + '/' + anoInicial;
    dtMYII := '01/' + mesFinal + '/' + anoFinal;

    Case StrToInt(mesInicial) Of
      01 : mesInicial := 'Janeiro';
      02 : mesInicial := 'Fevereiro';
      03 : mesInicial := 'Março';
      04 : mesInicial := 'Abril';
      05 : mesInicial := 'Maio';
      06 : mesInicial := 'Junho';
      07 : mesInicial := 'Julho';
      08 : mesInicial := 'Agosto';
      09 : mesInicial := 'Setembro';
      10 : mesInicial := 'Outubro';
      11 : mesInicial := 'Novembro';
      12 : mesInicial := 'Dezembro';
    end;

      //RETORNA NOME DO MÊS FINAL
    Case StrToInt(mesFinal) Of
      01 : mesFinal := 'Janeiro';
      02 : mesFinal := 'Fevereiro';
      03 : mesFinal := 'Março';
      04 : mesFinal := 'Abril';
      05 : mesFinal := 'Maio';
      06 : mesFinal := 'Junho';
      07 : mesFinal := 'Julho';
      08 : mesFinal := 'Agosto';
      09 : mesFinal := 'Setembro';
      10 : mesFinal := 'Outubro';
      11 : mesFinal := 'Novembro';
      12 : mesFinal := 'Dezembro';
    end;

    montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' de ' + anoInicial + '.</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';
  End;

  qryRO.Active := true;

  roID := qryRO.FieldByName('ID').AsString;
  resumoID := roID;
  resumoNome := qryRO.FieldByName('NOME').AsString;

  navBarFunction := MontaNavBarRO('AttGraphTwo', dateInicial, dateFinal, true, monthIHTML, monthIIHTML, true, campanhasID, false, true);

  campanhaCodigo := caCodigo;

  //MONTA PORQUE PRECISA DE DATASET PARA ATUALIZAR
  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' if exists(select * ' +
        ' from INFORMATION_SCHEMA.TABLES ' +
        ' where TABLE_NAME = ''TempTable_MonthOP'') ' +
        ' drop table TempTable_MonthOP ' +

        ' create table TempTable_MonthOP ' +
        ' (id int IDENTITY(1,1) NOT NULL, Nome varchar(100), QuantidadeTotal VARCHAR(10)');

        for I := 1 to 31 do
        Begin

          Add(', Quantidade' + I.ToString + ' varchar(10)');

        End;

    Add(' ) ');

  End;

  qryGeral1.ExecSQL;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' declare @strSQL varchar(MAX) ' +
        ' declare @Nome varchar(100) ' +
        ' declare @QuantidadeDia int ' +
        ' declare @QuantidadeMes int ' +
        ' declare @CaCodigo varchar(10) ' +
        ' declare @UsuariosID int ' +
        ' declare @mes int = 5 ' +

        ' declare cursor1 cursor for ' +

        ' SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
        ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
        ' WHERE Usuarios_ID <> 13 AND ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
        ' AND ISNULL(UsAtivo,0) = 1 ' +
        ' GROUP BY Usuarios_ID ' +

        ' open cursor1 ' +

        ' fetch next from cursor1 into @UsuariosID ' +

        '  while @@FETCH_STATUS = 0 ' +
        '  Begin ' +

        '    declare @i int = 1 ' +
        '    declare @update varchar(8000) ' +
        '    declare @idReg int ' +

        '        insert into TempTable_MonthOP (Nome, QuantidadeTotal) ' +
        '        values(dbo.RetornaNomeUsuario(CONVERT(VARCHAR(2), @UsuariosID)), ' +
        '        (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
        '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND  ' +
        '          CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') AND ' +
        '          ISNULL(CoUsuariosID,0) = CONVERT(VARCHAR(2), @UsuariosID))) ' +

        '        select @idReg = IDENT_CURRENT(''TempTable_MonthOP'') ' +
        '        set @i = 1 ' +

        '        if (SELECT ISNULL(QuantidadeTotal,0) FROM TempTable_MonthOP WHERE id = @idReg) > 0' +
        '        Begin ' +
          '        set @update = '''' ' +



          '        while @i <= 31 ' +
          '        Begin ' +

          '          set @update = ''Update TempTable_MonthOP set Quantidade'' + convert(varchar(2),@i) + '' = '' ' +
          '          set @update = @update + '' ' +
          '          (SELECT convert(varchar(10), CASE WHEN COUNT(CoResumoOperacaoID) = 0 THEN '''''''' ELSE COUNT(CoResumoOperacaoID) END) ' +
          '            FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
          '            INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
          '            WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
          '            DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = '' + convert(varchar(2), @i) + '' ' +
          '            AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dtMYI + ''''' AND EOMONTH(''''' + dtMYII + ''''') ' +
          '            AND ISNULL(CoUsuariosID,0) = '' + convert(varchar(10), @UsuariosID) + '') ' +
          '          Where id='' + convert(varchar(100), @idReg) ' +
          '          set @i = @i + 1; ' +
          '          exec (@update) ' +

          '        End ' +
        ' End ' +
        ' Else ' +
        ' Begin ' +

          ' DELETE FROM TempTable_MonthOP WHERE id = @idReg ' +

        ' End ' +

        '  fetch next from cursor1 into @UsuariosID ' +
        '  End ' +

        ' close cursor1 ' +
        ' deallocate cursor1 ');

  End;

  dmDados.GravaMsgArquivoAppLocal('StringTeste.txt',' Graph II ' + #13 + qryGeral1.SQL.Text, false);
  qryGeral1.ExecSQL;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' SELECT UPPER(Nome) [OPERADOR] ');

        for I := 1 to 31 do
        Begin

          Add(
          ', CASE Quantidade' + I.ToString + ' WHEN 0 THEN '''' ELSE Quantidade' + I.ToString + ' END [' + I.ToString + '] ');

        End;

    Add(' , CASE QuantidadeTotal WHEN 0 THEN '''' ELSE QuantidadeTotal END [TOTAL DO MÊS] ' +
      ' from TempTable_MonthOP order by Nome ');

  End;


  try
    qryGeral1.Active := true;
  except
    on E: EFDDBEngineException do
    begin

    end;

  end;

  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral1.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';

  headerCharts_1 := WebCharts1.ContinuosProject;
  headerCharts_2 := WebCharts1.ContinuosProject;
  headerCharts_3 := WebCharts1.ContinuosProject;
  headerCharts_4 := WebCharts1.ContinuosProject;
  headerCharts_5 := WebCharts1.ContinuosProject;
  headerCharts_6 := WebCharts1.ContinuosProject;
  headerCharts_7 := WebCharts1.ContinuosProject;
  headerCharts_8 := WebCharts1.ContinuosProject;


  config_row_1_Charts := headerCharts_1
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes1')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_2_Charts := headerCharts_2
            .Charts
              ._ChartType(bar)
                .Attributes
                  .Name('BAROpXMes1')
                  .ColSpan(4)
                  .Heigth(340)
                  .Labelling
                    .Format('0')
                    .RGBColor('30,30,30')
                    .FontSize(20)
                    .FontStyle('normal')
                    .FontFamily(fontFamily)
                  .&End
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                    .Scales
                      .GeneratedAxes(false)
                      .Axes
                        .yAxe
                          .Ticks
                            .fontColor('#FFF')
                          .&End
                          .GridLines
                            .drawTicks(false)
                          .&End
                        .&End
                      .&End
                    .&End
                  .&End;

  config_row_3_Charts := headerCharts_3
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes2')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_4_Charts := headerCharts_4
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BAROpXMes2')
          .ColSpan(4)
          .Heigth(350)
          .Labelling
            .Format('0')
            .RGBColor('30,30,30')
            .FontSize(20)
            .FontStyle('normal')
            .FontFamily(fontFamily)
          .&End
          .Options
            .Legend
              .display(true)
              .position('bottom')
            .&End
            .Scales
              .GeneratedAxes(false)
              .Axes
                .yAxe
                  .Ticks
                    .fontColor('#FFF')
                  .&End
                  .GridLines
                    .drawTicks(false)
                  .&End
                .&End
              .&End
            .&End
          .&End;
//
  config_row_5_Charts := headerCharts_5
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes3')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_6_Charts := headerCharts_6
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BAROpXMes3')
          .ColSpan(4)
          .Heigth(350)
          .Labelling
            .Format('0')
            .RGBColor('30,30,30')
            .FontSize(20)
            .FontStyle('normal')
            .FontFamily(fontFamily)
          .&End
          .Options
            .Legend
              .display(true)
              .position('bottom')
            .&End
            .Scales
              .GeneratedAxes(false)
              .Axes
                .yAxe
                  .Ticks
                    .fontColor('#FFF')
                  .&End
                  .GridLines
                    .drawTicks(false)
                  .&End
                .&End
              .&End
            .&End
          .&End;

  config_row_7_Charts := headerCharts_7
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes4')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_8_Charts := headerCharts_8
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BAROpXMes4')
          .ColSpan(4)
          .Heigth(350)
          .Labelling
            .Format('0')
            .RGBColor('30,30,30')
            .FontSize(20)
            .FontStyle('normal')
            .FontFamily(fontFamily)
          .&End
          .Options
            .Legend
              .display(true)
              .position('bottom')
            .&End
            .Scales
              .GeneratedAxes(false)
              .Axes
                .yAxe
                  .Ticks
                    .fontColor('#FFF')
                  .&End
                  .GridLines
                    .drawTicks(false)
                  .&End
                .&End
              .&End
            .&End
          .&End;


  LimpaQuery(qryGeral1);

  qryGeral1.Active := false;
  qryGeral1.SQL.Add(' SELECT UPPER(dbo.RetornaNomeUsuario(Usuarios_ID)) Nome, Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
            ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
            ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
            ' WHERE ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
            ' AND ISNULL(UsAtivo,0) = 1 ' +
            ' GROUP BY Usuarios_ID ORDER BY Nome ');
  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;

  SetLength(arrayQry, tamanhoQry);
  SetLength(arrayQry2, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

     usuarioID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
     usuarioNome := qryGeral1.FieldByName('Nome').AsString;


     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(' if exists(select * ' +
      '  from INFORMATION_SCHEMA.TABLES ' +
      '  where TABLE_NAME = ''TempDash_Operador' + usuarioID.ToString + '_30Dias'') ' +
      '  drop table TempDash_Operador' + usuarioID.ToString + '_30Dias ' +

      '  create table TempDash_Operador' + usuarioID.ToString + '_30Dias ' +
      '  (id int IDENTITY(1,1) NOT NULL, nome int, quantidadeTotal int) ');

      arrayQry[I].ExecSQL;

      arrayQry[I].SQL.Add('  declare @i int = 1 ' +
      '    set @i = 1 ' +

      '    while @i <= 31 ' +
      '    Begin ' +

      '      insert into TempDash_Operador' + usuarioID.ToString + '_30Dias (nome, quantidadeTotal) ' +
      '      values((@i), ' +
      '      (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo +
      '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
      '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
      '          DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = convert(varchar(2), @i) ' +
      '          AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') ' +
      '          AND ISNULL(CoUsuariosID,0) = ' + usuarioID.ToString + ')) ' +
      '      set @i = @i + 1; ' +

      '    End ' +

      '  SELECT Nome Label, quantidadeTotal Value, ' +
      '  CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1)) as RGB ' +
      '  from TempDash_Operador' + usuarioID.ToString + '_30Dias ' +
      '  order by Label');

      arrayQry[I].Active := True;

      usuarioRGB := arrayQry[I].FieldByName('RGB').AsString;

      config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      arrayQry2[I] := TFDQuery.Create(self);
      arrayQry2[I].Connection := dmDados.Conn;

      arrayQry2[I].Active := false;

      arrayQry2[I].SQL.Clear;
      arrayQry2[I].SQL.Add(
        'SELECT ''Total'' label, SUM(QuantidadeTotal) value, '''' RGB FROM TempDash_Operador' + usuarioID.ToString + '_30Dias');
      arrayQry2[I].Active := true;


      config_row_2_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      config_row_3_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;
//
      config_row_4_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;
//
      config_row_5_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      config_row_6_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      config_row_7_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      config_row_8_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;


     finally


     end;

    qryGeral1.Next;

  End;

  config_row_1_Charts.&End
              .&End
            ;

  config_row_2_Charts.&End
              .&End
            ;

  config_row_3_Charts.&End
              .&End
            ;

  config_row_4_Charts.&End
              .&End
            ;
//
  config_row_5_Charts.&End
              .&End
            ;

  config_row_6_Charts.&End
              .&End
            ;

  config_row_7_Charts.&End
              .&End
            ;

  config_row_8_Charts.&End
              .&End
            ;



  WebCharts1
  .BackgroundColor('#FFFFFF')
     .FontColor('#000000')
     .Container(fluid)
     .AddResource('<link href="css/green.css" rel="stylesheet">')
     .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
     .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
     .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
     .AddResource('<script src="/scripts/snippet-javascript-console.min.js?v=1"></script>')

     .NewProject

     .Rows
      .HTML(navBarFunction)
     .&End

       .Jumpline

       .Rows
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-2" style="font-family: Product Sans">Comparativos por Contatos</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
       .&End

       .Jumpline

       .Rows
        .ID('infoDate')
        .HTML(montaInfo)
       .&End

      .Jumpline

      .Rows
        .ID('Att-RO1Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
       .&End

      .Rows
        .ID('RO1Table')
        .HTML(montaTableOPRO)
      .&End

      .Rows
        ._Div
          .ColSpan(1)
        .&End
        .Tag
          .Add(headerCharts_1.HTML)
        .&End
        .Tag
          .Add(headerCharts_2.HTML)
        .&End
        ._Div
          .ColSpan(1)
        .&End
      .&End

      .Jumpline

      .Rows
        .ID('Att-RO2Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
      .&End

      .Rows
        .ID('RO2Table')
        .HTML('')
      .&End

      .Rows
        ._Div
          .ColSpan(1)
        .&End
        .Tag
          .Add(headerCharts_3.HTML)
        .&End
        .Tag
          .Add(headerCharts_4.HTML)
        .&End
        ._Div
          .ColSpan(1)
        .&End
      .&End
      //RO2 FINALIZADO
      .Jumpline

      //DASH PARA O RO3
      .Rows
        .ID('Att-RO3Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
       .&End

      .Rows
        .ID('RO3Table')
        .HTML('')
      .&End

      .Rows
        ._Div
          .ColSpan(1)
        .&End
        .Tag
          .Add(headerCharts_5.HTML)
        .&End
        .Tag
          .Add(headerCharts_6.HTML)
        .&End
        ._Div
          .ColSpan(1)
        .&End
      .&End
      //RO3 FINALIZADO
      .Jumpline

      //DASH PARA O RO4
      .Rows
        .ID('Att-RO4Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
       .&End

      .Rows
        .ID('RO4Table')
        .HTML('')
      .&End

      .Rows
        ._Div
          .ColSpan(1)
        .&End
        .Tag
          .Add(headerCharts_7.HTML)
        .&End
        .Tag
          .Add(headerCharts_8.HTML)
        .&End
        ._Div
          .ColSpan(1)
        .&End
      .&End
      //RO4 FINALIZADO
      .Jumpline


    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .CallbackJS
      .ClassProvider(Self)
    .&End
    .Generated;

    for i := 0 to tamanhoQry - 1 do
    begin
      FreeAndNil(arrayQry[i]);
      FreeAndNil(arrayQry2[i]);
    end;


    form_esmaecer_fundo.Hide;

end;

procedure TunitHome.AttGraphTwo(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
var
  I, J, tamanhoQry, columnQry: Integer;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  dtMYI, dtMYII: String;
  montaInfo: String;
  montaTableOPRO: String;

  resumoID, resumoNome: String;

  usuarioID: Integer;
  usuarioNome, usuarioRGB: String;

  arrayQry, arrayQry2: array of TFDQuery;

  headerCharts, bodyCharts, footerCharts: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts: iModelHTMLChartsConfig;
begin

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if campanhaCodigo = '' then
  Begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(campanhaCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  if roID = '' then
  Begin

    qryRO.Active := true;

    roID := qryRO.FieldByName('ID').AsString;
    resumoID := roID;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  End
  else
  begin

    LimpaQuery(qryRO);
    qryRO.Open('SELECT ResumoDeOperacoes_ID ID, ReNome NOME FROM ResumoDeOperacoes WITH(NOLOCK) ' +
      ' WHERE ResumoDeOperacoes_ID = ' + roID);

    resumoID := qryRO.FieldByName('ID').AsString;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  end;

  if (monthStart <> '') and (monthFinish <> '') then
  Begin

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);

    anoInicial := monthStart.Substring(0,4);
    anoFinal := monthFinish.Substring(0,4);

    dtMYI := '01/' + mesInicial + '/' + anoInicial;
    dtMYII := '01/' + mesFinal + '/' + anoFinal;

    //RETORNA NOME DO MÊS INICIAL
    Case StrToInt(mesInicial) Of
      01 : mesInicial := 'Janeiro';
      02 : mesInicial := 'Fevereiro';
      03 : mesInicial := 'Março';
      04 : mesInicial := 'Abril';
      05 : mesInicial := 'Maio';
      06 : mesInicial := 'Junho';
      07 : mesInicial := 'Julho';
      08 : mesInicial := 'Agosto';
      09 : mesInicial := 'Setembro';
      10 : mesInicial := 'Outubro';
      11 : mesInicial := 'Novembro';
      12 : mesInicial := 'Dezembro';
    end;

      //RETORNA NOME DO MÊS FINAL
    Case StrToInt(mesFinal) Of
      01 : mesFinal := 'Janeiro';
      02 : mesFinal := 'Fevereiro';
      03 : mesFinal := 'Março';
      04 : mesFinal := 'Abril';
      05 : mesFinal := 'Maio';
      06 : mesFinal := 'Junho';
      07 : mesFinal := 'Julho';
      08 : mesFinal := 'Agosto';
      09 : mesFinal := 'Setembro';
      10 : mesFinal := 'Outubro';
      11 : mesFinal := 'Novembro';
      12 : mesFinal := 'Dezembro';
    end;

    if anoInicial = anoFinal then
    Begin

      if mesInicial = mesFinal then
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' de ' + anoInicial + '.</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End
      else
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' até ' + mesFinal + ' de ' + anoInicial + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End;

    End
    else
    Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + '/' + anoInicial + '  até ' + mesFinal + '/' + anoFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

    End;

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);
  End;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('Att-RO1Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
        .Update;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('infoDate')
        .HTML(montaInfo)
        .Update;

  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' if exists(select * ' +
        ' from INFORMATION_SCHEMA.TABLES ' +
        ' where TABLE_NAME = ''TempTable_MonthOP'') ' +
        ' drop table TempTable_MonthOP ' +

        ' create table TempTable_MonthOP ' +
        ' (id int IDENTITY(1,1) NOT NULL, Nome varchar(100), QuantidadeTotal VARCHAR(10)');

        for I := 1 to 31 do
        Begin

          Add(', Quantidade' + I.ToString + ' varchar(10)');

        End;

    Add(' ) ');

  End;

  qryGeral1.ExecSQL;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' declare @strSQL varchar(MAX) ' +
        ' declare @Nome varchar(100) ' +
        ' declare @QuantidadeDia int ' +
        ' declare @QuantidadeMes int ' +
        ' declare @CaCodigo varchar(10) ' +
        ' declare @UsuariosID int ' +
        ' declare @mes int = 5 ' +

        ' declare cursor1 cursor for ' +

        ' SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
        ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
        ' WHERE Usuarios_ID <> 13 AND ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
        ' AND ISNULL(UsAtivo,0) = 1 ' +
        ' GROUP BY Usuarios_ID ' +

        ' open cursor1 ' +

        ' fetch next from cursor1 into @UsuariosID ' +

        '  while @@FETCH_STATUS = 0 ' +
        '  Begin ' +

        '    declare @i int = 1 ' +
        '    declare @update varchar(8000) ' +
        '    declare @idReg int ' +

        '        insert into TempTable_MonthOP (Nome, QuantidadeTotal) ' +
        '        values(dbo.RetornaNomeUsuario(CONVERT(VARCHAR(2), @UsuariosID)), ' +
        '        (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
        '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND  ' +
        '           ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') AND ' +
        '          ISNULL(CoUsuariosID,0) = CONVERT(VARCHAR(2), @UsuariosID))) ' +

        '        select @idReg = IDENT_CURRENT(''TempTable_MonthOP'') ' +
        '        set @i = 1 ' +

        '        if (SELECT ISNULL(QuantidadeTotal,0) FROM TempTable_MonthOP WHERE id = @idReg) > 0' +
        '        Begin ' +
          '        set @update = '''' ' +



          '        while @i <= 31 ' +
          '        Begin ' +

          '          set @update = ''Update TempTable_MonthOP set Quantidade'' + convert(varchar(2),@i) + '' = '' ' +
          '          set @update = @update + '' ' +
          '          (SELECT convert(varchar(10), CASE WHEN COUNT(CoResumoOperacaoID) = 0 THEN '''''''' ELSE COUNT(CoResumoOperacaoID) END) ' +
          '            FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
          '            INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
          '            WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
          '            DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = '' + convert(varchar(2), @i) + '' ' +
          '            AND ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dtMYI + ''''' AND EOMONTH(''''' + dtMYII + ''''') ' +
          '            AND ISNULL(CoUsuariosID,0) = '' + convert(varchar(10), @UsuariosID) + '') ' +
          '          Where id='' + convert(varchar(100), @idReg) ' +
          '          set @i = @i + 1; ' +
          '          exec (@update) ' +

          '        End ' +
        ' End ' +
        ' Else ' +
        ' Begin ' +

          ' DELETE FROM TempTable_MonthOP WHERE id = @idReg ' +

        ' End ' +

        '  fetch next from cursor1 into @UsuariosID ' +
        '  End ' +

        ' close cursor1 ' +
        ' deallocate cursor1 ');

  End;

  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  //TOTAL POR DIA NO FINAL
  qryGeral1.SQL.Add(
    ' DECLARE @I INT = 1 ' +
    ' DECLARE @STR VARCHAR(MAX) = '''' ' +
    '  ' +
    ' SET @STR = ''INSERT INTO TempTable_MonthOP( Nome, QuantidadeTotal '' ' +
    '  ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', Quantidade'' + CONVERT(VARCHAR(5), @I) ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    ' SET @STR = @STR + '') SELECT ''''TOTAL'''', SUM(CONVERT(INT, QuantidadeTotal)) '' ' +
    ' SET @I = 1 ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', SUM(CONVERT(INT, Quantidade'' + CONVERT(VARCHAR(5), @I) + '')) '' ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    '  ' +
    ' SET @STR = @STR + '' FROM TempTable_MonthOP'' ' +
    '  ' +
    ' EXEC (@STR) ');
  //FINAL DO TOTAL POR DIA
  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' SELECT UPPER(Nome) [OPERADOR] ');

        for I := 1 to 31 do
        Begin

          Add(
          ', CASE Quantidade' + I.ToString + ' WHEN 0 THEN '''' ELSE Quantidade' + I.ToString + ' END [' + I.ToString + '] ');

        End;

    Add(' , CASE QuantidadeTotal WHEN 0 THEN '''' ELSE QuantidadeTotal END [TOTAL DO MÊS] ' +
      ' from TempTable_MonthOP order by Nome ');

  End;


  try
    qryGeral1.Active := true;
  except
    on E: EFDDBEngineException do
    begin

    end;

  end;

  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral1.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';


  WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('RO1Table')
        .HTML(montaTableOPRO)
        .Update;

  headerCharts := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes1')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_2_Charts := headerCharts
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BAROpXMes1')
          .ColSpan(4)
          .Heigth(350)
          .Options
            .Legend
              .display(true)
              .position('bottom')
            .&End
          .&End;


  LimpaQuery(qryGeral1);

  qryGeral1.Active := false;
  qryGeral1.SQL.Add(' SELECT UPPER(dbo.RetornaNomeUsuario(Usuarios_ID)) Nome, Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
            ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
            ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
            ' WHERE ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
            ' AND ISNULL(UsAtivo,0) = 1 ' +
            ' GROUP BY Usuarios_ID ORDER BY Nome ');
  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;

  SetLength(arrayQry, tamanhoQry);
  SetLength(arrayQry2, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

     usuarioID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
     usuarioNome := qryGeral1.FieldByName('Nome').AsString;


     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(' if exists(select * ' +
      '  from INFORMATION_SCHEMA.TABLES ' +
      '  where TABLE_NAME = ''TempDash_Operador' + usuarioID.ToString + '_30Dias'') ' +
      '  drop table TempDash_Operador' + usuarioID.ToString + '_30Dias ' +

      '  create table TempDash_Operador' + usuarioID.ToString + '_30Dias ' +
      '  (id int IDENTITY(1,1) NOT NULL, nome int, quantidadeTotal int) ');

      arrayQry[I].ExecSQL;

      arrayQry[I].SQL.Add('  declare @i int = 1 ' +
      '    set @i = 1 ' +

      '    while @i <= 31 ' +
      '    Begin ' +

      '      insert into TempDash_Operador' + usuarioID.ToString + '_30Dias (nome, quantidadeTotal) ' +
      '      values((@i), ' +
      '      (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo +
      '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
      '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
      '          DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = convert(varchar(2), @i) ' +
      '          AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') ' +
      '          AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ISNULL(CoUsuariosID,0) = ' + usuarioID.ToString + ')) ' +
      '      set @i = @i + 1; ' +

      '    End ' +

      '  SELECT Nome Label, quantidadeTotal Value, ' +
      '  CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1)) as RGB ' +
      '  from TempDash_Operador' + usuarioID.ToString + '_30Dias ' +
      '  order by Label');

      arrayQry[I].Active := True;

      usuarioRGB := arrayQry[I].FieldByName('RGB').AsString;

      config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      arrayQry2[I] := TFDQuery.Create(self);
      arrayQry2[I].Connection := dmDados.Conn;

      arrayQry2[I].Active := false;

      arrayQry2[I].SQL.Clear;
      arrayQry2[I].SQL.Add(
        'SELECT ''Total'' label, SUM(QuantidadeTotal) value, '''' RGB FROM TempDash_Operador' + usuarioID.ToString + '_30Dias');
      arrayQry2[I].Active := true;


      config_row_2_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;


     finally


     end;

    qryGeral1.Next;

  End;

  config_row_1_Charts.&End
              .UpdateChart;

  config_row_2_Charts.&End
              .UpdateChart;


  for i := 0 to tamanhoQry - 1 do
  begin
    FreeAndNil(arrayQry[i]);
    FreeAndNil(arrayQry2[i]);
  end;

end;

procedure TunitHome.GraphThree(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: string;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');

     caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
     campanhaCodigo := caCodigo;
     campanhaNome := qryCamp.FieldByName('CaNome').AsString;
     campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  navBarFunction := MontaNavBarHTML('AttGraphThree', dateStart, dateFinish, false, '', '');

  WebCharts1
  .BackgroundColor('#FFFFFF')
     .FontColor('#000000')
     .Container(fluid)
     .AddResource('<link href="css/green.css" rel="stylesheet">')
     .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
     .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
     .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
     .AddResource('<script src="/scripts/snippet-javascript-console.min.js?v=1"></script>')

     .NewProject

     .Rows
      .HTML(navBarFunction)
     .&End

     .Container(true)

       .Jumpline

       .Rows
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-2" style="font-family: Product Sans">Produção Geral</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
       .&End

       .Jumpline

       .Rows
        .ID('infoDate')
        .HTML(
          ' <figure class="text-center"> ' +
            ' <blockquote class="blockquote"> ' +
              ' <p style=" font-family: Product Sans;">Faça um filtro para buscar os dados gerais.</p> ' +
            ' </blockquote> ' +
            ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' </figcaption> ' +
          ' </figure> '
        )
       .&End

       .Jumpline

       .Rows
        .ID('CardsGeral')
        .HTML('')
       .&End


    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .CallbackJS
      .ClassProvider(Self)
    .&End
    .Generated;

    form_esmaecer_fundo.Hide;
end;

procedure TunitHome.AttGraphTwoII(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
var
  I, J, tamanhoQry, columnQry: Integer;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  dtMYI, dtMYII: String;
  montaInfo: String;
  montaTableOPRO: String;

  resumoID, resumoNome: String;

  usuarioID: Integer;
  usuarioNome, usuarioRGB: String;

  arrayQry, arrayQry2: array of TFDQuery;

  headerCharts, headerCharts_2, bodyCharts, footerCharts: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts: iModelHTMLChartsConfig;
begin

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if campanhaCodigo = '' then
  Begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(campanhaCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  if roID = '' then
  Begin

    qryRO.Active := true;

    roID := qryRO.FieldByName('ID').AsString;
    resumoID := roID;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  End
  else
  begin

    LimpaQuery(qryRO);
    qryRO.Open('SELECT ResumoDeOperacoes_ID ID, ReNome NOME FROM ResumoDeOperacoes WITH(NOLOCK) ' +
      ' WHERE ResumoDeOperacoes_ID = ' + roID);

    resumoID := qryRO.FieldByName('ID').AsString;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  end;

  if (monthStart <> '') and (monthFinish <> '') then
  Begin

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);

    anoInicial := monthStart.Substring(0,4);
    anoFinal := monthFinish.Substring(0,4);

    dtMYI := '01/' + mesInicial + '/' + anoInicial;
    dtMYII := '01/' + mesFinal + '/' + anoFinal;

    //RETORNA NOME DO MÊS INICIAL
    Case StrToInt(mesInicial) Of
      01 : mesInicial := 'Janeiro';
      02 : mesInicial := 'Fevereiro';
      03 : mesInicial := 'Março';
      04 : mesInicial := 'Abril';
      05 : mesInicial := 'Maio';
      06 : mesInicial := 'Junho';
      07 : mesInicial := 'Julho';
      08 : mesInicial := 'Agosto';
      09 : mesInicial := 'Setembro';
      10 : mesInicial := 'Outubro';
      11 : mesInicial := 'Novembro';
      12 : mesInicial := 'Dezembro';
    end;

      //RETORNA NOME DO MÊS FINAL
    Case StrToInt(mesFinal) Of
      01 : mesFinal := 'Janeiro';
      02 : mesFinal := 'Fevereiro';
      03 : mesFinal := 'Março';
      04 : mesFinal := 'Abril';
      05 : mesFinal := 'Maio';
      06 : mesFinal := 'Junho';
      07 : mesFinal := 'Julho';
      08 : mesFinal := 'Agosto';
      09 : mesFinal := 'Setembro';
      10 : mesFinal := 'Outubro';
      11 : mesFinal := 'Novembro';
      12 : mesFinal := 'Dezembro';
    end;

    if anoInicial = anoFinal then
    Begin

      if mesInicial = mesFinal then
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' de ' + anoInicial + '.</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End
      else
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' até ' + mesFinal + ' de ' + anoInicial + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End;

    End
    else
    Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + '/' + anoInicial + '  até ' + mesFinal + '/' + anoFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

    End;

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);
  End;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('Att-RO2Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
        .Update;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('infoDate')
        .HTML(montaInfo)
        .Update;

  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' if exists(select * ' +
        ' from INFORMATION_SCHEMA.TABLES ' +
        ' where TABLE_NAME = ''TempTable_MonthOP2'') ' +
        ' drop table TempTable_MonthOP2 ' +

        ' create table TempTable_MonthOP2 ' +
        ' (id int IDENTITY(1,1) NOT NULL, Nome varchar(100), QuantidadeTotal VARCHAR(10)');

        for I := 1 to 31 do
        Begin

          Add(', Quantidade' + I.ToString + ' varchar(10)');

        End;

    Add(' ) ');

  End;

  qryGeral1.ExecSQL;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' declare @strSQL varchar(MAX) ' +
        ' declare @Nome varchar(100) ' +
        ' declare @QuantidadeDia int ' +
        ' declare @QuantidadeMes int ' +
        ' declare @CaCodigo varchar(10) ' +
        ' declare @UsuariosID int ' +
        ' declare @mes int = 5 ' +

        ' declare ccursor2 cursor for ' +

        ' SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
        ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
        ' WHERE Usuarios_ID <> 13 AND ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
        ' AND ISNULL(UsAtivo,0) = 1 ' +
        ' GROUP BY Usuarios_ID ' +

        ' open ccursor2 ' +

        ' fetch next from ccursor2 into @UsuariosID ' +

        '  while @@FETCH_STATUS = 0 ' +
        '  Begin ' +

        '    declare @i int = 1 ' +
        '    declare @update varchar(8000) ' +
        '    declare @idReg int ' +

        '        insert into TempTable_MonthOP2 (Nome, QuantidadeTotal) ' +
        '        values(dbo.RetornaNomeUsuario(CONVERT(VARCHAR(2), @UsuariosID)), ' +
        '        (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
        '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND  ' +
        '           ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') AND ' +
        '          ISNULL(CoUsuariosID,0) = CONVERT(VARCHAR(2), @UsuariosID))) ' +

        '        select @idReg = IDENT_CURRENT(''TempTable_MonthOP2'') ' +
        '        set @i = 1 ' +

        '        if (SELECT ISNULL(QuantidadeTotal,0) FROM TempTable_MonthOP2 WHERE id = @idReg) > 0' +
        '        Begin ' +
          '        set @update = '''' ' +



          '        while @i <= 31 ' +
          '        Begin ' +

          '          set @update = ''Update TempTable_MonthOP2 set Quantidade'' + convert(varchar(2),@i) + '' = '' ' +
          '          set @update = @update + '' ' +
          '          (SELECT convert(varchar(10), CASE WHEN COUNT(CoResumoOperacaoID) = 0 THEN '''''''' ELSE COUNT(CoResumoOperacaoID) END) ' +
          '            FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
          '            INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
          '            WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
          '            DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = '' + convert(varchar(2), @i) + '' ' +
          '            AND ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dtMYI + ''''' AND EOMONTH(''''' + dtMYII + ''''') ' +
          '            AND ISNULL(CoUsuariosID,0) = '' + convert(varchar(10), @UsuariosID) + '') ' +
          '          Where id='' + convert(varchar(100), @idReg) ' +
          '          set @i = @i + 1; ' +
          '          exec (@update) ' +

          '        End ' +
        ' End ' +
        ' Else ' +
        ' Begin ' +

          ' DELETE FROM TempTable_MonthOP2 WHERE id = @idReg ' +

        ' End ' +

        '  fetch next from ccursor2 into @UsuariosID ' +
        '  End ' +

        ' close ccursor2 ' +
        ' deallocate ccursor2 ');

  End;

  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  //TOTAL POR DIA NO FINAL
  qryGeral1.SQL.Add(
    ' DECLARE @I INT = 1 ' +
    ' DECLARE @STR VARCHAR(MAX) = '''' ' +
    '  ' +
    ' SET @STR = ''INSERT INTO TempTable_MonthOP2( Nome, QuantidadeTotal '' ' +
    '  ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', Quantidade'' + CONVERT(VARCHAR(5), @I) ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    ' SET @STR = @STR + '') SELECT ''''TOTAL'''', SUM(CONVERT(INT, QuantidadeTotal)) '' ' +
    ' SET @I = 1 ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', SUM(CONVERT(INT, Quantidade'' + CONVERT(VARCHAR(5), @I) + '')) '' ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    '  ' +
    ' SET @STR = @STR + '' FROM TempTable_MonthOP2'' ' +
    '  ' +
    ' EXEC (@STR) ');
  //FINAL DO TOTAL POR DIA
  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' SELECT UPPER(Nome) [OPERADOR] ');

        for I := 1 to 31 do
        Begin

          Add(
          ', CASE Quantidade' + I.ToString + ' WHEN 0 THEN '''' ELSE Quantidade' + I.ToString + ' END [' + I.ToString + '] ');

        End;

    Add(' , CASE QuantidadeTotal WHEN 0 THEN '''' ELSE QuantidadeTotal END [TOTAL DO MÊS] ' +
      ' from TempTable_MonthOP2 order by Nome ');

  End;


  try
    qryGeral1.Active := true;
  except
    on E: EFDDBEngineException do
    begin

    end;

  end;

  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral1.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';


  WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('RO2Table')
        .HTML(montaTableOPRO)
        .Update;

  headerCharts := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  headerCharts_2 := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes2')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_2_Charts := headerCharts_2
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BAROpXMes2')
          .ColSpan(4)
          .Heigth(350)
          .Options
            .Legend
              .display(true)
              .position('bottom')
            .&End
          .&End;


  LimpaQuery(qryGeral1);

  qryGeral1.Active := false;
  qryGeral1.SQL.Add(' SELECT UPPER(dbo.RetornaNomeUsuario(Usuarios_ID)) Nome, Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
            ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
            ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
            ' WHERE ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
            ' AND ISNULL(UsAtivo,0) = 1 ' +
            ' GROUP BY Usuarios_ID ORDER BY Nome ');
  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;

  SetLength(arrayQry, tamanhoQry);
  SetLength(arrayQry2, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

     usuarioID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
     usuarioNome := qryGeral1.FieldByName('Nome').AsString;


     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(' if exists(select * ' +
      '  from INFORMATION_SCHEMA.TABLES ' +
      '  where TABLE_NAME = ''TempDash_Operador' + usuarioID.ToString + '_30Dias2'') ' +
      '  drop table TempDash_Operador' + usuarioID.ToString + '_30Dias2 ' +

      '  create table TempDash_Operador' + usuarioID.ToString + '_30Dias2 ' +
      '  (id int IDENTITY(1,1) NOT NULL, nome int, quantidadeTotal int) ');

      arrayQry[I].ExecSQL;

      arrayQry[I].SQL.Add('  declare @i int = 1 ' +
      '    set @i = 1 ' +

      '    while @i <= 31 ' +
      '    Begin ' +

      '      insert into TempDash_Operador' + usuarioID.ToString + '_30Dias2 (nome, quantidadeTotal) ' +
      '      values((@i), ' +
      '      (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo +
      '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
      '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
      '          DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = convert(varchar(2), @i) ' +
      '          AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') ' +
      '          AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ISNULL(CoUsuariosID,0) = ' + usuarioID.ToString + ')) ' +
      '      set @i = @i + 1; ' +

      '    End ' +

      '  SELECT Nome Label, quantidadeTotal Value, ' +
      '  CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1)) as RGB ' +
      '  from TempDash_Operador' + usuarioID.ToString + '_30Dias2 ' +
      '  order by Label');

      arrayQry[I].Active := True;

      usuarioRGB := arrayQry[I].FieldByName('RGB').AsString;

      config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      arrayQry2[I] := TFDQuery.Create(self);
      arrayQry2[I].Connection := dmDados.Conn;

      arrayQry2[I].Active := false;

      arrayQry2[I].SQL.Clear;
      arrayQry2[I].SQL.Add(
        'SELECT ''Total'' label, SUM(QuantidadeTotal) value, '''' RGB FROM TempDash_Operador' + usuarioID.ToString + '_30Dias2');
      arrayQry2[I].Active := true;


      config_row_2_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;


     finally


     end;

    qryGeral1.Next;

  End;

  config_row_1_Charts.&End
              .UpdateChart;

  config_row_2_Charts.&End
              .UpdateChart;


  for i := 0 to tamanhoQry - 1 do
  begin
    FreeAndNil(arrayQry[i]);
    FreeAndNil(arrayQry2[i]);
  end;

end;

procedure TunitHome.AttGraphTwoIII(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
var
  I, J, tamanhoQry, columnQry: Integer;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  dtMYI, dtMYII: String;
  montaInfo: String;
  montaTableOPRO: String;

  resumoID, resumoNome: String;

  usuarioID: Integer;
  usuarioNome, usuarioRGB: String;

  arrayQry, arrayQry2: array of TFDQuery;

  headerCharts, headerCharts_2, bodyCharts, footerCharts: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts: iModelHTMLChartsConfig;
begin

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if campanhaCodigo = '' then
  Begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(campanhaCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  if roID = '' then
  Begin

    qryRO.Active := true;

    roID := qryRO.FieldByName('ID').AsString;
    resumoID := roID;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  End
  else
  begin

    LimpaQuery(qryRO);
    qryRO.Open('SELECT ResumoDeOperacoes_ID ID, ReNome NOME FROM ResumoDeOperacoes WITH(NOLOCK) ' +
      ' WHERE ResumoDeOperacoes_ID = ' + roID);

    resumoID := qryRO.FieldByName('ID').AsString;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  end;

  if (monthStart <> '') and (monthFinish <> '') then
  Begin

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);

    anoInicial := monthStart.Substring(0,4);
    anoFinal := monthFinish.Substring(0,4);

    dtMYI := '01/' + mesInicial + '/' + anoInicial;
    dtMYII := '01/' + mesFinal + '/' + anoFinal;

    //RETORNA NOME DO MÊS INICIAL
    Case StrToInt(mesInicial) Of
      01 : mesInicial := 'Janeiro';
      02 : mesInicial := 'Fevereiro';
      03 : mesInicial := 'Março';
      04 : mesInicial := 'Abril';
      05 : mesInicial := 'Maio';
      06 : mesInicial := 'Junho';
      07 : mesInicial := 'Julho';
      08 : mesInicial := 'Agosto';
      09 : mesInicial := 'Setembro';
      10 : mesInicial := 'Outubro';
      11 : mesInicial := 'Novembro';
      12 : mesInicial := 'Dezembro';
    end;

      //RETORNA NOME DO MÊS FINAL
    Case StrToInt(mesFinal) Of
      01 : mesFinal := 'Janeiro';
      02 : mesFinal := 'Fevereiro';
      03 : mesFinal := 'Março';
      04 : mesFinal := 'Abril';
      05 : mesFinal := 'Maio';
      06 : mesFinal := 'Junho';
      07 : mesFinal := 'Julho';
      08 : mesFinal := 'Agosto';
      09 : mesFinal := 'Setembro';
      10 : mesFinal := 'Outubro';
      11 : mesFinal := 'Novembro';
      12 : mesFinal := 'Dezembro';
    end;

    if anoInicial = anoFinal then
    Begin

      if mesInicial = mesFinal then
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' de ' + anoInicial + '.</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End
      else
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' até ' + mesFinal + ' de ' + anoInicial + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End;

    End
    else
    Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + '/' + anoInicial + '  até ' + mesFinal + '/' + anoFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

    End;

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);
  End;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('Att-RO3Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
        .Update;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('infoDate')
        .HTML(montaInfo)
        .Update;

  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' if exists(select * ' +
        ' from INFORMATION_SCHEMA.TABLES ' +
        ' where TABLE_NAME = ''TempTable_MonthOP3'') ' +
        ' drop table TempTable_MonthOP3 ' +

        ' create table TempTable_MonthOP3 ' +
        ' (id int IDENTITY(1,1) NOT NULL, Nome varchar(100), QuantidadeTotal VARCHAR(10)');

        for I := 1 to 31 do
        Begin

          Add(', Quantidade' + I.ToString + ' varchar(10)');

        End;

    Add(' ) ');

  End;

  qryGeral1.ExecSQL;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' declare @strSQL varchar(MAX) ' +
        ' declare @Nome varchar(100) ' +
        ' declare @QuantidadeDia int ' +
        ' declare @QuantidadeMes int ' +
        ' declare @CaCodigo varchar(10) ' +
        ' declare @UsuariosID int ' +
        ' declare @mes int = 5 ' +

        ' declare ccursor3 cursor for ' +

        ' SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
        ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
        ' WHERE Usuarios_ID <> 13 AND ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
        ' AND ISNULL(UsAtivo,0) = 1 ' +
        ' GROUP BY Usuarios_ID ' +

        ' open ccursor3 ' +

        ' fetch next from ccursor3 into @UsuariosID ' +

        '  while @@FETCH_STATUS = 0 ' +
        '  Begin ' +

        '    declare @i int = 1 ' +
        '    declare @update varchar(8000) ' +
        '    declare @idReg int ' +

        '        insert into TempTable_MonthOP3 (Nome, QuantidadeTotal) ' +
        '        values(dbo.RetornaNomeUsuario(CONVERT(VARCHAR(2), @UsuariosID)), ' +
        '        (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
        '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND  ' +
        '          ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') AND ' +
        '          ISNULL(CoUsuariosID,0) = CONVERT(VARCHAR(2), @UsuariosID))) ' +

        '        select @idReg = IDENT_CURRENT(''TempTable_MonthOP3'') ' +
        '        set @i = 1 ' +

        '        if (SELECT ISNULL(QuantidadeTotal,0) FROM TempTable_MonthOP3 WHERE id = @idReg) > 0' +
        '        Begin ' +
          '        set @update = '''' ' +

          '        while @i <= 31 ' +
          '        Begin ' +

          '          set @update = ''Update TempTable_MonthOP3 set Quantidade'' + convert(varchar(2),@i) + '' = '' ' +
          '          set @update = @update + '' ' +
          '          (SELECT convert(varchar(10), CASE WHEN COUNT(CoResumoOperacaoID) = 0 THEN '''''''' ELSE COUNT(CoResumoOperacaoID) END) ' +
          '            FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
          '            INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
          '            WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
          '            DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = '' + convert(varchar(2), @i) + '' ' +
          '            AND ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dtMYI + ''''' AND EOMONTH(''''' + dtMYII + ''''') ' +
          '            AND ISNULL(CoUsuariosID,0) = '' + convert(varchar(10), @UsuariosID) + '') ' +
          '          Where id='' + convert(varchar(100), @idReg) ' +
          '          set @i = @i + 1; ' +
          '          exec (@update) ' +

          '        End ' +
        ' End ' +
        ' Else ' +
        ' Begin ' +

          ' DELETE FROM TempTable_MonthOP3 WHERE id = @idReg ' +

        ' End ' +

        '  fetch next from ccursor3 into @UsuariosID ' +
        '  End ' +

        ' close ccursor3 ' +
        ' deallocate ccursor3 ');

  End;

  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  //TOTAL POR DIA NO FINAL
  qryGeral1.SQL.Add(
    ' DECLARE @I INT = 1 ' +
    ' DECLARE @STR VARCHAR(MAX) = '''' ' +
    '  ' +
    ' SET @STR = ''INSERT INTO TempTable_MonthOP3( Nome, QuantidadeTotal '' ' +
    '  ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', Quantidade'' + CONVERT(VARCHAR(5), @I) ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    ' SET @STR = @STR + '') SELECT ''''TOTAL'''', SUM(CONVERT(INT, QuantidadeTotal)) '' ' +
    ' SET @I = 1 ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', SUM(CONVERT(INT, Quantidade'' + CONVERT(VARCHAR(5), @I) + '')) '' ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    '  ' +
    ' SET @STR = @STR + '' FROM TempTable_MonthOP3'' ' +
    '  ' +
    ' EXEC (@STR) ');
  //FINAL DO TOTAL POR DIA
  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' SELECT UPPER(Nome) [OPERADOR] ');

        for I := 1 to 31 do
        Begin

          Add(
          ', CASE Quantidade' + I.ToString + ' WHEN 0 THEN '''' ELSE Quantidade' + I.ToString + ' END [' + I.ToString + '] ');

        End;

    Add(' , CASE QuantidadeTotal WHEN 0 THEN '''' ELSE QuantidadeTotal END [TOTAL DO MÊS] ' +
      ' from TempTable_MonthOP3 order by Nome ');

  End;


  try
    qryGeral1.Active := true;
  except
    on E: EFDDBEngineException do
    begin

    end;

  end;

  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral1.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';


  WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('RO3Table')
        .HTML(montaTableOPRO)
        .Update;

  headerCharts := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  headerCharts_2 := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes3')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_2_Charts := headerCharts_2
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BAROpXMes3')
          .ColSpan(4)
          .Heigth(350)
          .Options
            .Legend
              .display(true)
              .position('bottom')
            .&End
          .&End;


  LimpaQuery(qryGeral1);

  qryGeral1.Active := false;
  qryGeral1.SQL.Add(' SELECT UPPER(dbo.RetornaNomeUsuario(Usuarios_ID)) Nome, Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
            ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
            ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
            ' WHERE ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
            ' AND ISNULL(UsAtivo,0) = 1 ' +
            ' GROUP BY Usuarios_ID ORDER BY Nome ');
  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;

  SetLength(arrayQry, tamanhoQry);
  SetLength(arrayQry2, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

     usuarioID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
     usuarioNome := qryGeral1.FieldByName('Nome').AsString;


     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(' if exists(select * ' +
      '  from INFORMATION_SCHEMA.TABLES ' +
      '  where TABLE_NAME = ''TempDash_Operador' + usuarioID.ToString + '_30Dias3'') ' +
      '  drop table TempDash_Operador' + usuarioID.ToString + '_30Dias3 ' +

      '  create table TempDash_Operador' + usuarioID.ToString + '_30Dias3 ' +
      '  (id int IDENTITY(1,1) NOT NULL, nome int, quantidadeTotal int) ');

      arrayQry[I].ExecSQL;

      arrayQry[I].SQL.Add('  declare @i int = 1 ' +
      '    set @i = 1 ' +

      '    while @i <= 31 ' +
      '    Begin ' +

      '      insert into TempDash_Operador' + usuarioID.ToString + '_30Dias3 (nome, quantidadeTotal) ' +
      '      values((@i), ' +
      '      (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo +
      '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
      '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
      '          DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = convert(varchar(2), @i) ' +
      '          AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') ' +
      '          AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ISNULL(CoUsuariosID,0) = ' + usuarioID.ToString + ')) ' +
      '      set @i = @i + 1; ' +

      '    End ' +

      '  SELECT Nome Label, quantidadeTotal Value, ' +
      '  CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1)) as RGB ' +
      '  from TempDash_Operador' + usuarioID.ToString + '_30Dias3 ' +
      '  order by Label');

      arrayQry[I].Active := True;

      usuarioRGB := arrayQry[I].FieldByName('RGB').AsString;

      config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      arrayQry2[I] := TFDQuery.Create(self);
      arrayQry2[I].Connection := dmDados.Conn;

      arrayQry2[I].Active := false;

      arrayQry2[I].SQL.Clear;
      arrayQry2[I].SQL.Add(
        'SELECT ''Total'' label, SUM(QuantidadeTotal) value, '''' RGB FROM TempDash_Operador' + usuarioID.ToString + '_30Dias3 ');
      arrayQry2[I].Active := true;


      config_row_2_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;


     finally


     end;

    qryGeral1.Next;

  End;

  config_row_1_Charts.&End
              .UpdateChart;

  config_row_2_Charts.&End
              .UpdateChart;


  for i := 0 to tamanhoQry - 1 do
  begin
    FreeAndNil(arrayQry[i]);
    FreeAndNil(arrayQry2[i]);
  end;

end;

procedure TunitHome.AttGraphTwoIV(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = ''; roID: String = '');
var
  I, J, tamanhoQry, columnQry: Integer;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  dtMYI, dtMYII: String;
  montaInfo: String;
  montaTableOPRO: String;

  resumoID, resumoNome: String;

  usuarioID: Integer;
  usuarioNome, usuarioRGB: String;

  arrayQry, arrayQry2: array of TFDQuery;

  headerCharts, headerCharts_2, bodyCharts, footerCharts: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts: iModelHTMLChartsConfig;
begin

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if campanhaCodigo = '' then
  Begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(campanhaCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  if roID = '' then
  Begin

    qryRO.Active := true;

    roID := qryRO.FieldByName('ID').AsString;
    resumoID := roID;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  End
  else
  begin

    LimpaQuery(qryRO);
    qryRO.Open('SELECT ResumoDeOperacoes_ID ID, ReNome NOME FROM ResumoDeOperacoes WITH(NOLOCK) ' +
      ' WHERE ResumoDeOperacoes_ID = ' + roID);

    resumoID := qryRO.FieldByName('ID').AsString;
    resumoNome := qryRO.FieldByName('NOME').AsString;

  end;

  if (monthStart <> '') and (monthFinish <> '') then
  Begin

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);

    anoInicial := monthStart.Substring(0,4);
    anoFinal := monthFinish.Substring(0,4);

    dtMYI := '01/' + mesInicial + '/' + anoInicial;
    dtMYII := '01/' + mesFinal + '/' + anoFinal;

    //RETORNA NOME DO MÊS INICIAL
    Case StrToInt(mesInicial) Of
      01 : mesInicial := 'Janeiro';
      02 : mesInicial := 'Fevereiro';
      03 : mesInicial := 'Março';
      04 : mesInicial := 'Abril';
      05 : mesInicial := 'Maio';
      06 : mesInicial := 'Junho';
      07 : mesInicial := 'Julho';
      08 : mesInicial := 'Agosto';
      09 : mesInicial := 'Setembro';
      10 : mesInicial := 'Outubro';
      11 : mesInicial := 'Novembro';
      12 : mesInicial := 'Dezembro';
    end;

      //RETORNA NOME DO MÊS FINAL
    Case StrToInt(mesFinal) Of
      01 : mesFinal := 'Janeiro';
      02 : mesFinal := 'Fevereiro';
      03 : mesFinal := 'Março';
      04 : mesFinal := 'Abril';
      05 : mesFinal := 'Maio';
      06 : mesFinal := 'Junho';
      07 : mesFinal := 'Julho';
      08 : mesFinal := 'Agosto';
      09 : mesFinal := 'Setembro';
      10 : mesFinal := 'Outubro';
      11 : mesFinal := 'Novembro';
      12 : mesFinal := 'Dezembro';
    end;

    if anoInicial = anoFinal then
    Begin

      if mesInicial = mesFinal then
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' de ' + anoInicial + '.</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End
      else
      Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + ' até ' + mesFinal + ' de ' + anoInicial + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

      End;

    End
    else
    Begin

        montaInfo :=
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + mesInicial + '/' + anoInicial + '  até ' + mesFinal + '/' + anoFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> ';

    End;

    mesInicial := monthStart.Substring(5,2);
    mesFinal := monthFinish.Substring(5,2);
  End;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('Att-RO4Nome')
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-4" style="font-family: Product Sans">' + resumoNome + '</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
        .Update;

    WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('infoDate')
        .HTML(montaInfo)
        .Update;

  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' if exists(select * ' +
        ' from INFORMATION_SCHEMA.TABLES ' +
        ' where TABLE_NAME = ''TempTable_MonthOP4'') ' +
        ' drop table TempTable_MonthOP4 ' +

        ' create table TempTable_MonthOP4 ' +
        ' (id int IDENTITY(1,1) NOT NULL, Nome varchar(100), QuantidadeTotal VARCHAR(10)');

        for I := 1 to 31 do
        Begin

          Add(', Quantidade' + I.ToString + ' varchar(10)');

        End;

    Add(' ) ');

  End;

  qryGeral1.ExecSQL;

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' declare @strSQL varchar(MAX) ' +
        ' declare @Nome varchar(100) ' +
        ' declare @QuantidadeDia int ' +
        ' declare @QuantidadeMes int ' +
        ' declare @CaCodigo varchar(10) ' +
        ' declare @UsuariosID int ' +
        ' declare @mes int = 5 ' +

        ' declare ccursor4 cursor for ' +

        ' SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
        ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
        ' WHERE Usuarios_ID <> 13 AND ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
        ' AND ISNULL(UsAtivo,0) = 1 ' +
        ' GROUP BY Usuarios_ID ' +

        ' open ccursor4 ' +

        ' fetch next from ccursor4 into @UsuariosID ' +

        '  while @@FETCH_STATUS = 0 ' +
        '  Begin ' +

        '    declare @i int = 1 ' +
        '    declare @update varchar(8000) ' +
        '    declare @idReg int ' +

        '        insert into TempTable_MonthOP4 (Nome, QuantidadeTotal) ' +
        '        values(dbo.RetornaNomeUsuario(CONVERT(VARCHAR(2), @UsuariosID)), ' +
        '        (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
        '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND  ' +
        '          ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') AND ' +
        '          ISNULL(CoUsuariosID,0) = CONVERT(VARCHAR(2), @UsuariosID))) ' +

        '        select @idReg = IDENT_CURRENT(''TempTable_MonthOP4'') ' +
        '        set @i = 1 ' +

        '        if (SELECT ISNULL(QuantidadeTotal,0) FROM TempTable_MonthOP4 WHERE id = @idReg) > 0' +
        '        Begin ' +
          '        set @update = '''' ' +



          '        while @i <= 31 ' +
          '        Begin ' +

          '          set @update = ''Update TempTable_MonthOP4 set Quantidade'' + convert(varchar(2),@i) + '' = '' ' +
          '          set @update = @update + '' ' +
          '          (SELECT convert(varchar(10), CASE WHEN COUNT(CoResumoOperacaoID) = 0 THEN '''''''' ELSE COUNT(CoResumoOperacaoID) END) ' +
          '            FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
          '            INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
          '            WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
          '            DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = '' + convert(varchar(2), @i) + '' ' +
          '            AND ISNULL(CoCampanhaOrigemID,0) = 0 AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dtMYI + ''''' AND EOMONTH(''''' + dtMYII + ''''') ' +
          '            AND ISNULL(CoUsuariosID,0) = '' + convert(varchar(10), @UsuariosID) + '') ' +
          '          Where id='' + convert(varchar(100), @idReg) ' +
          '          set @i = @i + 1; ' +
          '          exec (@update) ' +

          '        End ' +
        ' End ' +
        ' Else ' +
        ' Begin ' +

          ' DELETE FROM TempTable_MonthOP4 WHERE id = @idReg ' +

        ' End ' +

        '  fetch next from ccursor4 into @UsuariosID ' +
        '  End ' +

        ' close ccursor4 ' +
        ' deallocate ccursor4 ');

  End;

  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  //TOTAL POR DIA NO FINAL
  qryGeral1.SQL.Add(
    ' DECLARE @I INT = 1 ' +
    ' DECLARE @STR VARCHAR(MAX) = '''' ' +
    '  ' +
    ' SET @STR = ''INSERT INTO TempTable_MonthOP4( Nome, QuantidadeTotal '' ' +
    '  ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', Quantidade'' + CONVERT(VARCHAR(5), @I) ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    ' SET @STR = @STR + '') SELECT ''''TOTAL'''', SUM(CONVERT(INT, QuantidadeTotal)) '' ' +
    ' SET @I = 1 ' +
    ' WHILE @I <= 31 ' +
    ' BEGIN ' +
    ' 	SET @STR = @STR + '', SUM(CONVERT(INT, Quantidade'' + CONVERT(VARCHAR(5), @I) + '')) '' ' +
    ' 	SET @I = @I + 1 ' +
    ' END ' +
    '  ' +
    ' SET @STR = @STR + '' FROM TempTable_MonthOP4'' ' +
    '  ' +
    ' EXEC (@STR) ');
  //FINAL DO TOTAL POR DIA
  qryGeral1.ExecSQL;

  LimpaQuery(qryGeral1);

  with qryGeral1.SQL do
  Begin

    Clear;
    Add(' SELECT UPPER(Nome) [OPERADOR] ');

        for I := 1 to 31 do
        Begin

          Add(
          ', CASE Quantidade' + I.ToString + ' WHEN 0 THEN '''' ELSE Quantidade' + I.ToString + ' END [' + I.ToString + '] ');

        End;

    Add(' , CASE QuantidadeTotal WHEN 0 THEN '''' ELSE QuantidadeTotal END [TOTAL DO MÊS] ' +
      ' from TempTable_MonthOP4 order by Nome ');

  End;


  try
    qryGeral1.Active := true;
  except
    on E: EFDDBEngineException do
    begin

    end;

  end;

  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral1.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral1.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';


  WebCharts1
      .ContinuosProject
      .WindowParent(FMXWindowParent)
      .WebBrowser(FMXChromium1)
      .DOMElement
        .ID('RO4Table')
        .HTML(montaTableOPRO)
        .Update;

  headerCharts := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  headerCharts_2 := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINEOpXMes4')
                  .ColSpan(6)
                  .Heigth(200)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                  .&End;

  config_row_2_Charts := headerCharts_2
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BAROpXMes4')
          .ColSpan(4)
          .Heigth(350)
          .Options
            .Legend
              .display(true)
              .position('bottom')
            .&End
          .&End;


  LimpaQuery(qryGeral1);

  qryGeral1.Active := false;
  qryGeral1.SQL.Add(' SELECT UPPER(dbo.RetornaNomeUsuario(Usuarios_ID)) Nome, Usuarios_ID FROM Usuarios WITH(NOLOCK) ' +
            ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
            ' INNER JOIN Campanhas WITH(NOLOCK) ON Campanhas_ID = UsCampanhasID ' +
            ' WHERE ISNULL(CaCodigo,'''') = ''' + caCodigo + ''' AND ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(CaEncerrada,0) = 0 ' +
            ' AND ISNULL(UsAtivo,0) = 1 ' +
            ' GROUP BY Usuarios_ID ORDER BY Nome ');
  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;

  SetLength(arrayQry, tamanhoQry);
  SetLength(arrayQry2, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

     usuarioID := qryGeral1.FieldByName('Usuarios_ID').AsInteger;
     usuarioNome := qryGeral1.FieldByName('Nome').AsString;


     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(' if exists(select * ' +
      '  from INFORMATION_SCHEMA.TABLES ' +
      '  where TABLE_NAME = ''TempDash_Operador' + usuarioID.ToString + '_30Dias4'') ' +
      '  drop table TempDash_Operador' + usuarioID.ToString + '_30Dias4 ' +

      '  create table TempDash_Operador' + usuarioID.ToString + '_30Dias4 ' +
      '  (id int IDENTITY(1,1) NOT NULL, nome int, quantidadeTotal int) ');

      arrayQry[I].ExecSQL;

      arrayQry[I].SQL.Add('  declare @i int = 1 ' +
      '    set @i = 1 ' +

      '    while @i <= 31 ' +
      '    Begin ' +

      '      insert into TempDash_Operador' + usuarioID.ToString + '_30Dias4 (nome, quantidadeTotal) ' +
      '      values((@i), ' +
      '      (SELECT COUNT(CoResumoOperacaoID) FROM Fone_List_' + caCodigo +
      '          INNER JOIN ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
      '          WHERE ISNULL(CoResumoOperacaoID,0) IN (' + roID + ') AND ' +
      '          DATEPART(DAY, CONVERT(DATE, CoDataInicioFicha)) = convert(varchar(2), @i) ' +
      '          AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dtMYI + ''' AND EOMONTH(''' + dtMYII + ''') ' +
      '          AND ISNULL(CoCampanhaOrigemID,0) = 0 AND ISNULL(CoUsuariosID,0) = ' + usuarioID.ToString + ')) ' +
      '      set @i = @i + 1; ' +

      '    End ' +

      '  SELECT Nome Label, quantidadeTotal Value, ' +
      '  CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1)) as RGB ' +
      '  from TempDash_Operador' + usuarioID.ToString + '_30Dias4 ' +
      '  order by Label');

      arrayQry[I].Active := True;

      usuarioRGB := arrayQry[I].FieldByName('RGB').AsString;

      config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;

      arrayQry2[I] := TFDQuery.Create(self);
      arrayQry2[I].Connection := dmDados.Conn;

      arrayQry2[I].Active := false;

      arrayQry2[I].SQL.Clear;
      arrayQry2[I].SQL.Add(
        'SELECT ''Total'' label, SUM(QuantidadeTotal) value, '''' RGB FROM TempDash_Operador' + usuarioID.ToString + '_30Dias4 ');
      arrayQry2[I].Active := true;


      config_row_2_Charts
              .DataSet
                .DataSet(arrayQry2[I])
                .textLabel(usuarioNome)
                .BackgroundColor(usuarioRGB)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(usuarioRGB)
              .&End;


     finally


     end;

    qryGeral1.Next;

  End;

  config_row_1_Charts.&End
              .UpdateChart;

  config_row_2_Charts.&End
              .UpdateChart;


  for i := 0 to tamanhoQry - 1 do
  begin
    FreeAndNil(arrayQry[i]);
    FreeAndNil(arrayQry2[i]);
  end;

end;

procedure TunitHome.AttGraphThree(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  grROID, grRONome: String;

  qryString: String;

  gridCards: String;
  mudaCor: Boolean;

  tamanhoQry, columnQry: Integer;
  totalVirgens, ativos, inativos, marcPeriodo: String;
  cliRec, cliCad, cliT, cliV: String;
  ligTotal, ligTMedio, ligTTotal: String;
  ficTotal, ficTMedio, ficTTotal: String;
  fichasTrabalhadas: String;
  I: Integer;
  J: Integer;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = '+ dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
      .ID('infoDate')
      .HTML(
          ' <figure class="text-center"> ' +
            ' <blockquote class="blockquote"> ' +
              ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
            ' </blockquote> ' +
            ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' </figcaption> ' +
          ' </figure> '
        )
      .Update;

  qryGrRO.Active := true;

  tamanhoQry := qryGrRO.RecordCount;
  columnQry := qryGrRO.FieldCount;

  qryString :=
    ' IF EXISTS(SELECT *  ' +
    ' 	FROM INFORMATION_SCHEMA.TABLES ' +
    ' 	WHERE TABLE_NAME = ''TempTable_OP_Camp_Geral'') ' +
    ' DROP TABLE TempTable_OP_Camp_Geral ' +
    '  ' +
    ' CREATE TABLE TempTable_OP_Camp_Geral ' +
    ' 	(TempTable_OP_Camp_Geral_ID INT NOT NULL IDENTITY(1,1), usID VARCHAR(10), campID VARCHAR(10), virgensTotal INT, ' +
    '     ativos INT, inativos INT, marcPeriodo INT, cRec INT, cCad INT, fT INT, fV INT, ' +
    '     ligTotal INT, ligTM TIME, ligTT TIME, ficTotal INT, ficTM TIME, ficTT TIME, pesT INT ';

  qryGrRO.First;

  for I := 0 to tamanhoQry - 1 do
  begin

    grROID := qryGrRO.FieldByName('ID').AsString;

    qryString := qryString +
      ' , grupoRO' + grROID + ' INT ';

    qryGrRO.Next;
  end;

  qryString := qryString + ' ) ';

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(qryString);

  qryString :=
    ' DECLARE @usID VARCHAR(10), @campID VARCHAR(10), @campCodigo VARCHAR(10) ' +
    ' DECLARE @strSQL VARCHAR(MAX) ' +
    '  DECLARE @ID INT ' +
    '  ' +
    ' DECLARE ccursor CURSOR FOR ' +
    '  ' +
    ' 	SELECT DISTINCT UsUsuariosID, UsCampanhasID, CaCodigo ' +
    ' 	FROM UsuariosCampanhas WITH(NOLOCK) ' +
    ' 	INNER JOIN Usuarios WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
    ' 	INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
    ' 	WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(UsAtivo,0) = 1 ' +
    ' 	ORDER BY UsUsuariosID ' +
    '  ' +
    ' OPEN ccursor ' +
    '  ' +
    ' FETCH NEXT FROM ccursor INTO @usID, @campID, @campCodigo ' +
    ' 	 ' +
    ' WHILE @@FETCH_STATUS = 0 ' +
    ' BEGIN ' +
    '  ' +
    ' 	SET @strSQL = '' ' +
    '  ' +
    '   INSERT INTO TempTable_OP_Camp_Geral(usID, campID, virgensTotal, ativos, inativos, marcPeriodo, cRec, cCad, fT, fV, ' +
    '     ligTotal, ligTM, ligTT, ficTotal, ficTM, ficTT, pesT) ' +
    ' 	VALUES('' + @usID + '', '' + @campID + '', ' +
    ' 		/*virgens total*/ ' +
    ' 		(SELECT COUNT(FN_PessoasID) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE ISNULL(FN_OperadorProprietarioID,0) = '' + @usID + '' AND ISNULL(FN_StatusFicha,'''''''') = ''''V''''), ' +
    ' 		/*ATIVOS*/ ' +
    ' 		(SELECT COUNT(FN_PessoasID) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE ISNULL(FN_OperadorProprietarioID,0) = '' + @usID + '' AND ISNULL(FN_AtivoNaLista,0) = 1 AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*INATIVOS*/ ' +
    ' 		(SELECT COUNT(FN_PessoasID) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE ISNULL(FN_OperadorProprietarioID,0) = '' + @usID + '' AND ISNULL(FN_AtivoNaLista,0) = 0 AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*MARCADOS NO PERÍODO*/ ' +
    ' 		(SELECT COUNT(AG_Agendas_'' + @campCodigo + ''_ID) FROM Agendas_'' + @campCodigo + '' WITH(NOLOCK) ' +
    ' 			INNER JOIN Fone_List_'' + @campCodigo + '' WITH(NOLOCK) ON AG_FN_FoneList_'' + @campCodigo + ''ID = FN_FoneList_'' + @campCodigo + ''_ID ' +
    ' 			WHERE ISNULL(FN_OperadorProprietarioID,0) = '' + @usID + '' AND CONVERT(DATE, AG_DataProximaLigacao) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''' AND ISNULL(AG_Status,'''''''') = ''''ABERTO''''), ' +
    ' 		/*CADASTROS RECEBIDOS*/ ' +
    ' 		(SELECT COUNT(FN_PessoasID) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) ' +
    ' 			INNER JOIN Pessoas WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
    ' 			WHERE ISNULL(FN_OperadorProprietarioID,0) = '' + @usID + '' AND ISNULL(PesUsuarioCadastrouID,0) NOT IN (0, '' + @usID + '') AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*CADASTROS CADASTRADOS*/ ' +
    ' 		(SELECT COUNT(FN_PessoasID) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) ' +
    ' 			INNER JOIN Pessoas WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
    ' 			WHERE ISNULL(FN_OperadorProprietarioID,0) = '' + @usID + '' AND ISNULL(PesUsuarioCadastrouID,0) IN ('' + @usID + '') AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*FICHAS TRABALHADOS*/ ' +
    ' 		(SELECT COUNT(DISTINCT FN_PessoasID) FROM ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK) ' +
    ' 			INNER JOIN Fone_List_'' + @campCodigo + '' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_'' + @campCodigo + ''_ID ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*FICHAS VIRGENS*/ ' +
    ' 		(SELECT COUNT(DISTINCT FN_PessoasID) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			WHERE ISNULL(FN_OperadorProprietarioID,0) = '' + @usID + '' AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''' ' +
    ' 			AND ISNULL(FN_StatusFicha,'''''''') = ''''V''''), ' +
    ' 		/*LIGAÇÕES QUANTIDADE TOTAL*/ ' +
    ' 		(SELECT COUNT(*) qntdLigacoes ' +
    ' 			FROM Contatos_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			INNER JOIN ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK) ON CO_ContatosFichaID = ContatosFichas_'' + @campCodigo + ''_ID  ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*LIGAÇÕES TEMPO MEDIO*/ ' +
    ' 		(SELECT ISNULL(CONVERT(TIME(0), DATEADD(SECOND, AVG(DATEDIFF(SECOND, ''''00:00:00.000'''', ISNULL(CO_Duracao,''''''''))), ''''00:00:00.000'''')), ''''00:00:00'''') ' +
    ' 			FROM Contatos_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			INNER JOIN ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK) ON CO_ContatosFichaID = ContatosFichas_'' + @campCodigo + ''_ID  ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*LIGAÇÕES TEMPO TOTAL*/ ' +
    ' 		(SELECT ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''''00:00:00.000'''', ISNULL(CO_Duracao,''''''''))), ''''00:00:00.000'''')), ''''00:00:00'''') ' +
    ' 			FROM Contatos_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			INNER JOIN ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK) ON CO_ContatosFichaID = ContatosFichas_'' + @campCodigo + ''_ID  ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*FICHAS TOTAIS*/ ' +
    ' 		(SELECT COUNT(*) qntdFichas ' +
    ' 			FROM ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*FICHAS TEMPO MEDIO*/ ' +
    ' 		(SELECT ISNULL(CONVERT(TIME(0), DATEADD(SECOND, AVG(DATEDIFF(SECOND, ''''00:00:00.000'''', ISNULL(CoDuracaoFicha,''''''''))), ''''00:00:00'''')), ''''00:00:00'''') ' +
    ' 			FROM ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*FICAHS TEMPO TOTAL*/ ' +
    ' 		(SELECT ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''''00:00:00.000'''', ISNULL(CoDuracaoFicha,''''''''))), ''''00:00:00'''')), ''''00:00:00'''') ' +
    ' 			FROM ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		/*CLIENTES TRABALHADOS*/ ' +
    ' 		(SELECT COUNT(DISTINCT CoFoneListsID) FROM ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK)  ' +
    ' 			WHERE ISNULL(CoUsuariosID,0) = '' + @usID + '' AND CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''') ';

    qryString := qryString + ' ) '' ';
    qryString := qryString + ' EXEC (@strSQL) ';
    qryString := qryString + ' SET @ID = IDENT_CURRENT(''TempTable_OP_Camp_Geral'') ';
    qryString := qryString + ' SET @strSQL = '' ';
    qryString := qryString + ' UPDATE TempTable_OP_Camp_Geral SET ';

    qryGrRO.First;

    for I := 0 to tamanhoQry - 1 do
    begin

      grROID := qryGrRO.FieldByName('ID').AsString;


      qryString := qryString +
        ' grupoRO' + grROID + ' = ' +
        ' (SELECT COUNT(*) FROM ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK) ' +
        '     INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = CoResumoOperacaoID ' +
        '     WHERE ReGrupoResumoOperacoesID = ''''' + grROID + ''''' AND ISNULL(CoUsuariosID,0) = '' + @usID + '' AND ' +
        '      CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''') ';

      if i <> tamanhoQry - 1 then
        qryString := qryString + ' , ';


      qryGrRO.Next;
    end;


    qryString := qryString +
    '  WHERE TempTable_OP_Camp_Geral_ID = '' + CONVERT(VARCHAR(10), @ID) + '' '' ' +
    '  ' +
    ' 		EXEC (@strSQL) ' +
    ' FETCH NEXT FROM ccursor INTO @usID, @campID, @campCodigo ' +

    ' END ' +

    ' CLOSE ccursor ' +
    ' DEALLOCATE ccursor ';

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(qryString);

  qryString :=
    ' SELECT UPPER(dbo.RetornaNomeUsuario(usID)) [Operador], SUM(virgensTotal) [Total de Virgens],  ' +
    ' 	SUM(ativos) [Ativos], SUM(inativos) [Inativos], SUM(marcPeriodo) [Marcados no Período],  ' +
    ' 	SUM(cRec) [Clientes Recebidos], SUM(cCad) [Clientes Cadastrados], SUM(fT) [Trabalhados],  ' +
    ' 	SUM(fV) [Virgens], SUM(ligTotal) [Ligações], ' +
    ' 	ISNULL(CONVERT(TIME(0), DATEADD(SECOND, AVG(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(ligTM,''''))), ''00:00:00'')), ''00:00:00'') [Tempo Médio], ' +
    ' 	ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(ligTT,''''))), ''00:00:00'')), ''00:00:00'') [Tempo Total] ';

  qryGrRO.First;


  for I := 0 to tamanhoQry - 1 do
  begin

    grROID := qryGrRO.FieldByName('ID').AsString;
    grRONome := qryGrRO.FieldByName('NOME').AsString;

    qryString := qryString +
      ' , SUM(grupoRO' + grROID + ') [' + grRONome + '] ';

    qryGrRO.Next;
  end;

  qryString := qryString +
    ' ,	SUM(ficTotal) [Fichas], ' +
    ' 	ISNULL(CONVERT(TIME(0), DATEADD(SECOND, AVG(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(ficTM,''''))), ''00:00:00'')), ''00:00:00'') [Tempo Médio], ' +
    ' 	ISNULL(CONVERT(TIME(0), DATEADD(SECOND, SUM(DATEDIFF(SECOND, ''00:00:00.000'', ISNULL(ficTT,''''))), ''00:00:00'')), ''00:00:00'') [Tempo Total], ' +
    ' 	SUM(pesT) [Clientes Trabalhados] ' +
    ' FROM TempTable_OP_Camp_Geral ' +
    ' GROUP BY usID ' +
    ' ORDER BY Operador ';

  dmDados.GravaMsgArquivoAppLocal('StringTeste.txt', 'Graph III' + #13 + qryString, false);


  LimpaQuery(qryGeral1);
  qryGeral1.Open(qryString);

  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  if tamanhoQry < 4 then
    gridCards := ' <div class="row row-cols-1 row-cols-md-' + tamanhoQry.ToString + ' g-4" style=" width: 80%; margin-left: 10%"> '
  else
    gridCards := ' <div class="row row-cols-1 row-cols-md-3 g-4" style=" width: 80%; margin-left: 10%"> ';


  qryGeral1.First;

  for I := 0 to tamanhoQry - 1 do
  Begin

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      begin

        gridCards := gridCards +
        ' <div class="col"> ' +
          ' <div class="card shadow p-3 mb-5 bg-body rounded" style=" background-color: #FFFFFF; margin-bottom: 10px; "> ' +
            ' <div class="card-header">' + qryGeral1.Fields[J].AsString + '</div>' +
            ' <div class="card-body"> ';

      end
      else if J = 1 then
      Begin

        gridCards := gridCards +
          ' <h5 class="card-title"> Total de Virgens: <span style=" display:block; float:right; margin-left:10px;">'
             + qryGeral1.Fields[J].AsString + '</span></h5> ';

      End
      else if J = 2 then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Ativos: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 3 then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Inativos: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 4 then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Marcados no Período: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 5 then
      Begin

        gridCards := gridCards +
           ' <hr> ' +
           ' <h4 class="card-title">Clientes Novos</h4> ' +
           ' <hr> ' +
           ' <p class="card-text">Clientes Recebidos: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 6 then
      Begin

        gridCards := gridCards +
           ' <p class="card-text">Clientes Cadastrados: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 7 then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Trabalhadas: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 8 then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Virgens: <span style=" display:block; float:right; margin-left:10px;">' +
          qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 9 then
      Begin

        gridCards := gridCards +
          ' <hr> ' +
          ' <h4 class="card-title">Fichas Trabalhadas</h4> ' +
          ' <hr> ' +
          ' <h5 class="card-title"><i class="fas fa-phone-volume"></i> Ligações</h5> ' +
          ' <hr> ' +
          ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">'
            + qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 10 then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if J = 11 then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ' +
          ' <hr> ';

      End
      else if qryGeral1.Fields[J].FieldName = 'Fichas' then
      Begin

        gridCards := gridCards +
          ' <hr> ' +
          ' <h5 class="card-title"><i class="fas fa-user"></i> Fichas</h5> ' +
          ' <hr> ' +
          ' <p class="card-text">Quantidade: <span style=" display:block; float:right; margin-left:10px;">'
            + qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if qryGeral1.Fields[J].FieldName = 'Tempo Médio' then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Tempo Médio: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if qryGeral1.Fields[J].FieldName = 'Tempo Total' then
      Begin

        gridCards := gridCards +
          ' <p class="card-text">Tempo Total: <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End
      else if qryGeral1.Fields[J].FieldName = 'Clientes Trabalhados' then
      Begin

        gridCards := gridCards +
          ' <hr> ' +
          ' <h5 class="card-title"> Clientes Trabalhados: <span style=" display:block; float:right; margin-left:10px;">' +
          qryGeral1.Fields[J].AsString + '</span></h5> ';

      End
      else
      Begin

        gridCards := gridCards +
          ' <p class="card-text">' + qryGeral1.Fields[J].FieldName + ': <span style=" display:block; float:right; margin-left:10px;">' +
            qryGeral1.Fields[J].AsString + '</span></p> ';

      End;

    End;

    gridCards := gridCards +
        ' </div> ' +
      ' </div> ' +
    ' </div> ';

    qryGeral1.Next;

  End;

  gridCards := gridCards + ' </div> ';

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
      .ID('CardsGeral')
      .HTML(gridCards)
      .Update;

  form_esmaecer_fundo.Hide;
end;

procedure TunitHome.GraphFour(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  dateIHTML, dateIIHTML: String;

  navBarFunction: String;
  montaSlide, montaCards, montaCardsI, montaCardsII: String;

  usuariosID: String;

  tamanhoQry, columnQry: Integer;
  I: Integer;
  J: Integer;

  arrayQry: array of TFDQuery;

  grupoROID, grupoRONome: String;

  rgbColor: string;

  headerCharts, bodyCharts, footerCharts: iModelHTML;
  config_row_1_Charts, config_row_2_Charts, config_row_3_Charts: iModelHTMLChartsConfig;

  userID: Integer;
  userNome: String;

  PesQntd, Fn1Qntd, Fn2Qntd, Fn3Qntd: integer;
  FnPesV1, FnPesV2, FnPesV3, FnCR1, FnCR2, FnCR3: integer;
  CaPes: integer;
  tempoTotalFicha, tempoTotalLig, tempoMedioFicha, tempoMedioLig, tempoTotalGeral: String;
  qntdLigacoes, qntdFichas: String;
  ativos, inativos, marcPeriodo: string;
  totalMensagens: String;

  qntdV, qntdT, qntdVT: String;
  clientesTrab: String;

  mudaCor: Boolean;
  montaTableOPRO, gridCards: String;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

     caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
     campanhaCodigo := caCodigo;
     campanhaNome := qryCamp.FieldByName('CaNome').AsString;
     campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
     qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
     ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
     ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  navBarFunction := MontaNavBarHTML('GraphFour', dateInicial, dateFinal, false, '', '', true, campanhasID);

  if (dmDados.GB_UsuarioID = '13') or (dmDados.GB_UsuarioID = '14')
    or (dmDados.GB_UsuarioID = '15') or (dmDados.GB_UsuarioID = '59') then
  Begin

    LimpaQuery(qryGeral1);
    qryGeral1.Open(
    ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE CONVERT(DATE, DataCad) ' +
    '   BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ' +
    ' 	ISNULL(PesCadastroImportado,0) = 1 ');

    montaCardsI :=
      ' <div class="col"> ' +
          ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 18rem; max-height: 8rem;"> ' +
            ' <div class="card-header"> Cadastro pela Importação </div>' +
            ' <div class="card-body"> ' +
            ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
            qryGeral1.FieldByName('QNTD').AsString + ' </p> ' +
            ' </div> ' +
          ' </div> ' +
        ' </div> ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(
    ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE CONVERT(DATE, DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ' +
    ' 	ISNULL(PesCadastroImportado,0) = 0 AND ISNULL(PesOrigemCadastro,'''') IN (''CS'', ''OP'', ''SU'') AND  ' +
    ' 	ISNULL(PesUsuarioCadastrouID,0) NOT IN  ' +
    ' 	(SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE UsNome LIKE ''%CRIARTE%'') ');

    montaCardsI :=  montaCardsI +
      ' <div class="col"> ' +
          ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 18rem; max-height: 8rem;"> ' +
            ' <div class="card-header"> Cadastro Manual Loja </div>' +
            ' <div class="card-body"> ' +
            ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
            qryGeral1.FieldByName('QNTD').AsString + ' </p> ' +
            ' </div> ' +
          ' </div> ' +
        ' </div> ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(
    ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE CONVERT(DATE, DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND ' +
    ' 	ISNULL(PesOrigemCadastro,'''') = ''CS'' AND ISNULL(PesUsuarioCadastrouID,0) IN  ' +
    ' 	(SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE UsNome LIKE ''%CRIARTE%'') ');

    montaCardsI :=  montaCardsI +
      ' <div class="col"> ' +
          ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 18rem; max-height: 8rem;"> ' +
            ' <div class="card-header"> Cadastro Manual Marketing </div>' +
            ' <div class="card-body"> ' +
            ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
            qryGeral1.FieldByName('QNTD').AsString + ' </p> ' +
            ' </div> ' +
          ' </div> ' +
        ' </div> ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(
    ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE CONVERT(DATE, DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' AND  ' +
    ' 	ISNULL(PesRDuuid,'''') <> '''' ');

    montaCardsI :=  montaCardsI +
      ' <div class="col"> ' +
          ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 18rem; max-height: 8rem;"> ' +
            ' <div class="card-header"> Cadastro RD Station </div>' +
            ' <div class="card-body"> ' +
            ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
            qryGeral1.FieldByName('QNTD').AsString + ' </p> ' +
            ' </div> ' +
          ' </div> ' +
        ' </div> ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(
    ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE Pessoas_ID NOT IN  ' +
    ' 	(SELECT CaPessoasID FROM CampanhasPessoas WITH(NOLOCK)) ');

    montaCardsI :=  montaCardsI +
      ' <div class="col"> ' +
          ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 18rem; max-height: 8rem;"> ' +
            ' <div class="card-header"> Pessoas sem Campanha </div>' +
            ' <div class="card-body"> ' +
            ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
            qryGeral1.FieldByName('QNTD').AsString + ' </p> ' +
            ' </div> ' +
          ' </div> ' +
        ' </div> ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(
    ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE CONVERT(DATE, DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ');

     montaCardsI :=  montaCardsI +
      ' <div class="col"> ' +
          ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 18rem; max-height: 8rem;"> ' +
            ' <div class="card-header"> Total de Leads </div>' +
            ' <div class="card-body"> ' +
            ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
            qryGeral1.FieldByName('QNTD').AsString + ' </p> ' +
            ' </div> ' +
          ' </div> ' +
        ' </div> ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open('SELECT Lojas_ID, LoNome FROM Lojas WITH(NOLOCK)');

    tamanhoQry := qryGeral1.RecordCount;

    for I := 0 to tamanhoQry do
    Begin
      if I = tamanhoQry then
      Begin

        LimpaQuery(qryGeral2);
        qryGeral2.Open(
        ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE CONVERT(DATE, DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ' +
        ' 	AND ISNULL(PesLojasID,0) = 0');

        montaCardsII := montaCardsII +
        ' <div class="col"> ' +
            ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 10rem; max-height: 8rem;"> ' +
              ' <div class="card-header"> Sem Loja </div>' +
              ' <div class="card-body"> ' +
              ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
              qryGeral2.FieldByName('QNTD').AsString + ' </p> ' +
              ' </div> ' +
            ' </div> ' +
          ' </div> ';

      End
      Else
      Begin

        LimpaQuery(qryGeral2);
        qryGeral2.Open(
        ' SELECT COUNT(*) QNTD FROM Pessoas WITH(NOLOCK) WHERE CONVERT(DATE, DataCad) BETWEEN '' ' + dateInicial + ' '' AND '' ' + dateFinal + ' '' ' +
        ' 	AND ISNULL(PesLojasID,0) = ' + qryGeral1.FieldByName('Lojas_ID').AsString);

        montaCardsII := montaCardsII +
        ' <div class="col"> ' +
            ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 10rem; max-height: 8rem;"> ' +
              ' <div class="card-header"> ' + qryGeral1.FieldByName('LoNome').AsString + ' </div>' +
              ' <div class="card-body"> ' +
              ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
              qryGeral2.FieldByName('QNTD').AsString + ' </p> ' +
              ' </div> ' +
            ' </div> ' +
          ' </div> ';

      End;

      qryGeral1.Next;
    End;

    montaCards :=
      ' <div class="row row-cols-1 row-cols-md-3 g-4" style=" width: 80%; margin-left: 10%"> ' +
        montaCardsI +
      ' </div> ' +

      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-5" style="font-family: Product Sans">Leads por Loja</p> ' +
        ' </blockquote> ' +
      ' </figure> ' +

      ' <div class="row row-cols-1 row-cols-md-5 g-4" style="width: 80%; margin-left: 10%"> ' +
        montaCardsII +
      ' </div> ';

  End;

  LimpaQuery(qryGeral1);
  qryGeral1.Active := false;
  qryGeral1.SQL.Add(
    ' IF EXISTS(SELECT *  ' +
    ' 	FROM INFORMATION_SCHEMA.TABLES ' +
    ' 	WHERE TABLE_NAME = ''TempTable_Info_Camps'') ' +
    ' DROP TABLE TempTable_Info_Camps ' +
    '  ' +
    ' CREATE TABLE TempTable_Info_Camps ' +
    ' 	(TempTable_Info_Camps_ID INT NOT NULL IDENTITY(1,1), campID VARCHAR(10), campCod VARCHAR(10), campName VARCHAR(100),  ' +
    ' 		virgensTotal INT, virgensProp INT, virgensSProp INT, fichasTrab INT, concR INT, opCamp INT) ');
  qryGeral1.ExecSQL;

  qryGeral1.SQL.Add(
    ' DECLARE @campID VARCHAR(10), @campCodigo VARCHAR(10), @campName VARCHAR(100) ' +
    ' DECLARE @strSQL VARCHAR(MAX) ' +
    '  ' +
    ' DECLARE ccurCamp CURSOR FOR ' +
    '  ' +
    '   SELECT Campanhas_ID, CaCodigo FROM Campanhas WITH(NOLOCK) ' +
    '   Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    '   WHERE CaEncerrada = 0 ORDER BY CaCodigo ' +
    '  ' +
    ' OPEN ccurCamp ' +
    '  ' +
    ' FETCH NEXT FROM ccurCamp INTO @campID, @campCodigo ' +
    ' 	 ' +
    ' WHILE @@FETCH_STATUS = 0 ' +
    ' BEGIN ' +
    '  ' +
    ' 	SET @strSQL = ''  ' +
    '  ' +
    ' 	INSERT INTO TempTable_Info_Camps ' +
    ' 	VALUES('' + @campID + '', '' + @campCodigo + '',  ' +
    ' 		(SELECT CONCAT(CaCodigo, '''' - '''', CaNome) FROM Campanhas WITH(NOLOCK) WHERE Campanhas_ID = '' + @campID + ''),  ' +
    ' 		(SELECT COUNT(*) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE ISNULL(FN_AtivoNaLista,0) = 1 AND ISNULL(FN_StatusFicha,'''''''') = ''''V'''' ' +
    ' 			AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''), ' +
    ' 		(SELECT COUNT(*) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE ISNULL(FN_AtivoNaLista,0) = 1 AND ISNULL(FN_StatusFicha,'''''''') = ''''V'''' ' +
    ' 			AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''' AND ISNULL(FN_OperadorProprietarioID,0) <> 0), ' +
    ' 		(SELECT COUNT(*) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE ISNULL(FN_AtivoNaLista,0) = 1 AND ISNULL(FN_StatusFicha,'''''''') = ''''V'''' ' +
    ' 			AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''' AND ISNULL(FN_OperadorProprietarioID,0) = 0), ' +
    ' 		(SELECT COUNT(*) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE ISNULL(FN_StatusFicha,'''''''') = ''''T'''' ' +
    ' 			AND CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''' AND FN_FoneList_'' + @campCodigo + ''_ID IN  ' +
    ' 				(SELECT CoFoneListsID FROM ContatosFichas_'' + @campCodigo + '' WITH(NOLOCK) WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + ''''')), ' +
    ' 		(SELECT COUNT(*) FROM Fone_List_'' + @campCodigo + '' WITH(NOLOCK) WHERE CONVERT(DATE, FN_DataCad) BETWEEN ''''' + dateInicial + ''''' AND ''''' + dateFinal + '''''  ' +
    ' 			AND ISNULL(FN_UltimoROID,0) IN (SELECT ResumoDeOperacoes_ID FROM ResumoDeOperacoes WITH(NOLOCK) WHERE ISNULL(ReAbreviacao,'''''''') = ''''CONC_RAP'''')), ' +
    ' 		(SELECT COUNT(DISTINCT UsUsuariosID) FROM UsuariosCampanhas WITH(NOLOCK) WHERE ISNULL(UsCampanhasID,0) = '' + @campID + '')) ' +
    '  ' +
    ' 		'' ' +
    ' 		 ' +
    ' 	EXEC (@strSQL) ' +
    '  ' +
    ' FETCH NEXT FROM ccurCamp INTO @campID, @campCodigo ' +
    '  ' +
    ' END ' +
    '  ' +
    ' CLOSE ccurCamp ' +
    ' DEALLOCATE ccurCamp ' +

    ' SELECT campName, virgensTotal [Fichas Virgens Totais], virgensProp [Fichas Virgens com Propriedade], ' +
    ' 	virgensSProp [Fichas Virgens sem Propriedade], fichasTrab [Fichas Trabalhadas], ' +
    ' 	concR [Concluídas Rápido], opCamp [Operadores na Campanha] ' +
    ' FROM TempTable_Info_Camps ');

  qryGeral1.Active := true;

  tamanhoQry := qryGeral1.RecordCount;
  columnQry := qryGeral1.FieldCount;

  montaSlide :=
    '<div id="carouselCampanhas" class="carousel carousel-dark slide" data-bs-ride="carousel"> ' +
    ' <div class="carousel-indicators"> ' +
    '   <button type="button" data-bs-target="#carouselCampanhas" data-bs-slide-to="0" class="active" ' +
    '     aria-current="true" aria-label="Slide 0"></button> ';

  for I := 1 to tamanhoQry do
  Begin
    montaSlide := montaSlide +
    '   <button type="button" data-bs-target="#carouselCampanhas" data-bs-slide-to="' + I.ToString + '" ' +
    '     class="active" aria-label="Slide ' + I.ToString + '"></button> ';

    qryGeral1.Next;
  End;

  montaSlide := montaSlide +
    ' </div > ' +
    ' <div class="carousel-inner"> ';

   if (dmDados.GB_UsuarioID = '13') or (dmDados.GB_UsuarioID = '14')
    or (dmDados.GB_UsuarioID = '15') or (dmDados.GB_UsuarioID = '59') then
  Begin
    montaSlide := montaSlide +
    '   <div class="carousel-item active" data-bs-interval="60000"> ' + montaCards +
    '   </div> ';
  End;


  qryGeral1.First;
  for I := 0 to tamanhoQry - 1 do
  Begin

    if (dmDados.GB_UsuarioID = '13') or (dmDados.GB_UsuarioID = '14')
    or (dmDados.GB_UsuarioID = '15') or (dmDados.GB_UsuarioID = '59') then
    Begin
      montaSlide := montaSlide +
      '   <div class="carousel-item" data-bs-interval="60000"> ' +
      '     <div class="row row-cols-1 row-cols-md-3 g-4" style=" width: 80%; margin-left: 15%; margin-bottom: 40px;"> ';
    End
    else if I = 0 then
    Begin
      montaSlide := montaSlide +
      '   <div class="carousel-item active" data-bs-interval="60000"> ' +
      '     <div class="row row-cols-1 row-cols-md-3 g-4" style=" width: 80%; margin-left: 15%; margin-bottom: 40px;"> ';
    End
    else
    Begin
      montaSlide := montaSlide +
      '   <div class="carousel-item" data-bs-interval="60000"> ' +
      '     <div class="row row-cols-1 row-cols-md-3 g-4" style=" width: 80%; margin-left: 15%; margin-bottom: 40px;"> ';
    End;


    for J := 1 to columnQry - 1 do
    Begin

     montaSlide := montaSlide +
      ' <div class="col"> ' +
          ' <div class="card border-dark mb-3 shadow p-3 mb-5 bg-body rounded" style="background-color: #FFFFFF; margin-bottom: 10px; max-width: 18rem; max-height: 8rem;"> ' +
            ' <div class="card-header"> ' + qryGeral1.Fields[J].FieldName + ' </div>' +
            ' <div class="card-body"> ' +
            ' <p class="card-text" style="font-family: Roboto; text-align: center; font-size: 1.5rem;">' +
            qryGeral1.Fields[J].AsString + ' </p> ' +
            ' </div> ' +
          ' </div> ' +
        ' </div> ';

    End;

    montaSlide := montaSlide +
    '     </div> ' +
    '     <div class="carousel-caption d-none d-md-block"> ' +
    '       <h5>' + qryGeral1.FieldByName('campName').AsString + '</h5> ' +
    '     </div> ' +
    '   </div> ';

    qryGeral1.Next;

  End;


  montaSlide := montaSlide +
    '  </div> ' +
    '  <button class="carousel-control-prev" type="button" data-bs-target="#carouselCampanhas" data-bs-slide="prev"> ' +
    '   <span class="carousel-control-prev-icon" aria-hidden="true"></span> ' +
    '   <span class="visually-hidden">Previous</span> ' +
    '  </button> ' +
    '  <button class="carousel-control-next" type="button" data-bs-target="#carouselCampanhas" data-bs-slide="next"> ' +
    '   <span class="carousel-control-next-icon" aria-hidden="true"></span> ' +
    '   <span class="visually-hidden">Next</span> ' +
    '  </button> ' +
    '</div> ';

  LimpaQuery(qryGeral3);
  qryGeral3.Active := false;
  qryGeral3.SQL.Add(
    ' WITH pessoas_1 AS (  ' +
    '  ' +
    ' SELECT DBO.RetornaNomeUsuario(FN_OperadorProprietarioID) OP,  ' +
    ' 	COUNT(DISTINCT FN_PessoasID) COFN,  ' +
    ' 	CAST(CAST(COUNT(DISTINCT FN_PessoasID) AS decimal(18,2)) AS decimal(18,2)) / CAST((SELECT COUNT(DISTINCT FN_PessoasID)  ' +
    ' 		FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' 		WHERE CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ' AND  ' +
    ' 			FN_OperadorProprietarioID IN (SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S''))  AS DECIMAL(18,2)) AS TOTAL,  ' +
    ' 	''RGB'' AS RGB, ROW_NUMBER() OVER (ORDER BY COUNT(*) desc) AS classificacao  ' +
    ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
    ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON FN_OperadorProprietarioID = UsUsuariosID ' +
    ' WHERE CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ' AND  ' +
    ' 	FN_OperadorProprietarioID IN (SELECT Usuarios_ID FROM Usuarios WITH(NOLOCK) WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S'') ' +
    '   AND UsCampanhasID = ' + campanhasID +
    ' GROUP BY FN_OperadorProprietarioID),  ' +
    '  ' +
    ' pessoas_2 AS (  ' +
    '  ' +
    ' SELECT OP AS Label,	 ' +
    ' TOTAL as Value, ' +
    ' 	RGB, classificacao   ' +
    ' FROM pessoas_1 )  ' +
    '  ' +
    '          SELECT Label, Value, GrRGB AS RGB  ' +
    '          FROM pessoas_2 AS p  ' +
    '          JOIN GraficosRGB AS g ON p.classificacao = g.GraficosRGB_ID  ' +
    '          ORDER BY p.Value DESC ');

  qryGeral3.Active := true;

  qryGrRO.Active := false;
  qryGrRO.Active := true;

  tamanhoQry := qryGrRO.RecordCount;


  qryGeral2.Active := false;
  qryGeral2.SQL.Clear;
  qryGeral2.SQL.Add(' if exists(select * ' +
      ' from INFORMATION_SCHEMA.TABLES ' +
      ' where TABLE_NAME = ''Temp_Table_FoneList'') ' +
      ' drop table Temp_Table_FoneList ' +

      ' create table Temp_Table_FoneList ' +
      ' (nomeUser varchar(100)');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      grupoROID := qryGrRO.FieldByName('ID').AsString;

      qryGeral2.SQL.Add(', gRO' + grupoROID + ' varchar(10) ');


      qryGrRO.Next;

    End;

    qryGeral2.SQL.Add(' , gROTotal varchar(10)) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , gROTotal varchar(10)) ');

  End;

  grupoROID := '';
  qryGrRO.First;

  qryGeral2.ExecSQL;

  qryGeral2.SQL.Add(
        ' declare @idUser int ' +

        ' declare cursor1 cursor for ' +

        ' SELECT DISTINCT FN_OperadorProprietarioID ' +
        ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON FN_OperadorProprietarioID = UsUsuariosID ' +
        ' WHERE CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ' AND ' +
        ' ISNULL(FN_OperadorProprietarioID,0) <> 0 ' +
        ' AND UsCampanhasID = ' + campanhasID +

        ' open cursor1 ' +
        ' fetch next from cursor1 into @idUser ' +
        ' while @@FETCH_STATUS = 0 ' +
        ' Begin ');

  //INSERT DENTRO DO CURSOR
  qryGeral2.SQL.Add('INSERT INTO Temp_Table_FoneList (nomeUser ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      grupoROID := qryGrRO.FieldByName('ID').AsString;

      qryGeral2.SQL.Add(', gRO' + grupoROID + ' ');


      qryGrRO.Next;

    End;

    qryGeral2.SQL.Add(' , gROTotal ) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , gROTotal ) ');

  End;

  grupoROID := '';
  qryGrRO.First;

  qryGeral2.SQL.Add('VALUES ( dbo.RetornaNomeUsuario(@idUser) ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      grupoROID := qryGrRO.FieldByName('ID').AsString;

      qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = FN_UltimoROID ' +
        '  WHERE ISNULL(FN_OperadorProprietarioID,0) = @idUser AND ' +
        '  ReGrupoResumoOperacoesID = ' + grupoROID + ' AND ' +
        '  CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ') ');

      qryGrRO.Next;

    End;

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = FN_UltimoROID ' +
        '  WHERE ISNULL(FN_OperadorProprietarioID,0) = @idUser AND ' +
        '  CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ') ) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = FN_UltimoROID ' +
        '  WHERE ISNULL(FN_OperadorProprietarioID,0) = @idUser AND ' +
        '  CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ') ) ');

  End;

  grupoROID := '';
  qryGrRO.First;

  qryGeral2.SQL.Add(' fetch next from cursor1 into @idUser' +
        ' End ' +
        ' close cursor1 ' +
        ' deallocate cursor1 ');

  dmDados.GravaMsgArquivoAppLocal('StringTeste.txt', 'Graph IV - 3' + #13 + qryGeral2.SQL.Text, false);

  qryGeral2.ExecSQL;
  qryGeral2.SQL.Clear;

  qryGeral2.SQL.Add('INSERT INTO Temp_Table_FoneList (nomeUser ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      grupoROID := qryGrRO.FieldByName('ID').AsString;

      qryGeral2.SQL.Add(', gRO' + grupoROID + ' ');


      qryGrRO.Next;

    End;

    qryGeral2.SQL.Add(' , gROTotal ) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , gROTotal ) ');

  End;

  grupoROID := '';
  qryGrRO.First;

  qryGeral2.SQL.Add('VALUES ( ''TOTAL'' ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      grupoROID := qryGrRO.FieldByName('ID').AsString;

      qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = FN_UltimoROID ' +
        '  WHERE ReGrupoResumoOperacoesID = ' + grupoROID + ' AND ' +
        '  CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ') ');

      qryGrRO.Next;

    End;

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = FN_UltimoROID ' +
        '  WHERE CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ') ) ');

  End
  else
  Begin

    qryGeral2.SQL.Add(', (SELECT COUNT(*) ' +
        '  FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
        '  INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = FN_UltimoROID ' +
        '  WHERE CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) + ') ) ');

  End;

  grupoROID := '';
  qryGrRO.First;

  qryGeral2.ExecSQL;
  qryGeral2.SQL.Clear;

  //SELECT FINAL DA TABELA
  qryGeral2.SQL.Add('SELECT UPPER(nomeUser) [OPERADORES] ');

  if (tamanhoQry <> 0) and (tamanhoQry > 0) then
  Begin

    for I := 0 to tamanhoQry - 1 do
    Begin

      grupoROID := qryGrRO.FieldByName('ID').AsString;
      grupoRONome := qryGrRO.FieldByName('Nome').AsString;

      qryGeral2.SQL.Add(', CASE ISNULL(gRO' + grupoROID + ','''') WHEN 0 ' +
        ' THEN '''' ELSE ISNULL(gRO' + grupoROID + ','''') END [' + grupoRONome + '] ');


      qryGrRO.Next;

    End;

    qryGeral2.SQL.Add(' , CASE ISNULL(gROTotal,'''') WHEN 0 THEN '''' ELSE ISNULL(gROTotal,'''') END [TOTAL] ');

  End
  else
  Begin

    qryGeral2.SQL.Add(' , CASE ISNULL(gROTotal,'''') WHEN 0 THEN '''' ELSE ISNULL(gROTotal,'''') END [TOTAL] ');

  End;

  qryGeral2.SQL.Add(' FROM Temp_Table_FoneList');

  try
    qryGeral2.Active := true;
  except
    on E: EFDDBEngineException do
    begin
         //dmdados.fdErrorDialog.Execute(E);
    end;

  end;

  tamanhoQry := qryGeral2.RecordCount;
  columnQry := qryGeral2.FieldCount;

  montaTableOPRO :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; ' +
        ' margin-left: 10%;"> ' +
      ' <thead> ' +
        ' <tr> ';

  for J := 0 to columnQry - 1 do
  Begin

     montaTableOPRO := montaTableOPRO +
      ' <th scope="col" style="text-align: center">' + qryGeral2.Fields[J].FieldName + ' </th> ' ;

  End;

  montaTableOPRO := montaTableOPRO +
        ' </tr> ' +
    ' </thead> ' +
    ' <tbody> ';

  for I := 0 to tamanhoQry - 1 do
  Begin

    montaTableOPRO := montaTableOPRO +
      ' <tr> ';

    for J := 0 to columnQry - 1 do
    Begin

      if J = 0 then
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <th scope="col" style="text-align: center">' + qryGeral2.Fields[J].AsString + ' </th> ' ;
      End
      else
      Begin
        montaTableOPRO := montaTableOPRO +
          ' <td style="text-align: center">' + qryGeral2.Fields[J].AsString + ' </td> ' ;
      End;

    End;

    montaTableOPRO := montaTableOPRO +
      ' </tr> ';

    qryGeral2.Next;

  End;

  montaTableOPRO := montaTableOPRO +
      ' </tbody> ' +
    ' </table> ';


  headerCharts := WebCharts1
          .ContinuosProject;

  config_row_1_Charts := headerCharts
            .Charts
              ._ChartType(horizontalBar)
                .Attributes
                  .Name('BAR_GrupoRO_OP')
                  .ColSpan(12)
                  .Heigth(180)
                  .Labelling
                    .Format('0')
                    .RGBColor('30,30,30')
                    .FontSize(20)
                    .FontStyle('normal')
                    .FontFamily(fontFamily)
                  .&End
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                    .Tooltip
                      .Format('0')
                      .DisplayTitle(true)
                    .&End
                  .&End;


  qryGrRO.Active := false;
  qryGrRO.Active := true;

  tamanhoQry := qryGrRO.RecordCount;

  SetLength(arrayQry, tamanhoQry);

  for I := 0 to tamanhoQry -1 do
  Begin

    grupoROID := qryGrRO.FieldByName('ID').AsString;
    grupoRONome := qryGrRO.FieldByName('NOME').AsString;


     try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].SQL.Add(
        ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDashConta_GrupoRO_OP' + grupoROID + ''') ' +
        ' 	DROP TABLE TempDashConta_GrupoRO_OP' + grupoROID + '  ' +
        '  ' +
        '  CREATE TABLE TempDashConta_GrupoRO_OP' + grupoROID + '  ' +
        '     (id INT IDENTITY(1,1) NOT NULL, UsuariosID VARCHAR(10), UsNome VARCHAR(100), countRO VARCHAR(10), RGB VARCHAR(15)) ' +
        '  ' +
        ' INSERT INTO TempDashConta_GrupoRO_OP' + grupoROID + '(UsuariosID, UsNome) ' +
        ' SELECT Usuarios_ID, UsNome ' +
        ' FROM Usuarios WITH(NOLOCK) ' +
        ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = Usuarios_ID ' +
        ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
        ' WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(UsAtivo,0) = 1 AND CaCodigo = ' + caCodigo + ' ' +
        ' ORDER BY UsNome ' +
        '  ' +
        ' SELECT UsuariosID ' +
        ' FROM TempDashConta_GrupoRO_OP' + grupoROID);

       arrayQry[I].Active := True;

       for J := 0 to arrayQry[I].RecordCount - 1 do
       Begin

        usuariosID := arrayQry[I].FieldByName('UsuariosID').AsString;

        dmDados.qryGeral.ExecSQL(
          ' UPDATE  TempDashConta_GrupoRO_OP' + grupoROID +
          ' SET countRO = (' +
          ' SELECT COUNT(FN_PessoasID) Value ' +
          ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
          ' INNER JOIN ResumoGrupos WITH(NOLOCK) ON ReResumoOperacoesID = FN_UltimoROID  ' +
          ' INNER JOIN UsuariosCampanhas WITH(NOLOCK) ON UsUsuariosID = FN_OperadorProprietarioID  ' +
          ' WHERE ReGrupoResumoOperacoesID IN (' + QuotedStr(grupoROID) + ') AND  ' +
          ' CONVERT(DATE, FN_DataCad) BETWEEN ' + QuotedStr(dateInicial) + ' AND ' + QuotedStr(dateFinal) +
          ' AND UsCampanhasID = ' + campanhasID +
          ' AND FN_OperadorProprietarioID IN (' + usuariosID + ') AND  ' +
          ' ReResumoOperacoesID IN (SELECT ResumoDeOperacoes_ID FROM ResumoDeOperacoes WHERE ISNULL(ReOcultarRelDash,0) = 0) ) ' +
          ' WHERE UsuariosID = ' + usuariosID);

        arrayQry[I].Next;
       End;

       LimpaQuery(arrayQry[I]);
       arrayQry[I].Open('SELECT DBO.RetornaNomeUsuario(UsuariosID) Label, countRO Value, '''' RGB ' +
        ' FROM TempDashConta_GrupoRO_OP' + grupoROID);

       rgbColor := IntToStr(Random(256)) + ',' + IntToStr(Random(256)) + ',' + IntToStr(Random(256));


       if arrayQry[I].RecordCount > 0 then
       Begin

        config_row_1_Charts
              .DataSet
                .textLabel(grupoRONome)
                .DataSet(arrayQry[I])
                .BackgroundColor(rgbColor)
                .BackgroundOpacity(8)
              .&End;
       End;

     finally


     end;


     qryGrRO.Next;

  End;

  config_row_1_Charts.&End.&End;

    WebCharts1
    .BackgroundColor('#FFFFFF')
    .FontColor('#000000')
    .Container(fluid)
    .AddResource('<link href="css/green.css" rel="stylesheet">')
    .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
    .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
    .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
    .AddResource('<script src="/scripts/snippet-javascript-console.min.js?v=1"></script>')

     .NewProject

     .Rows
      .HTML(navBarFunction)
     .&End

     .Container(true)

       .Jumpline

       .Rows
        .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p class="fs-2" style="font-family: Product Sans">Conta Corrente</p> ' +
          ' </blockquote> ' +
        ' </figure> ')
       .&End

       .Jumpline

       .Rows
        .ID('infoDate')
        .HTML(
          ' <figure class="text-center"> ' +
            ' <blockquote class="blockquote"> ' +
              ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
            ' </blockquote> ' +
            ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
              ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
            ' </figcaption> ' +
          ' </figure> '
        )
       .&End

       .Jumpline

       .Rows
        .ID('slide')
        .Tag
          .Add(montaSlide)
        .&End
       .&End

       .Jumpline

       .Rows
        ._Div
          .ColSpan(12)
          .Add(
            WebCharts1
            .ContinuosProject
              .Charts
                ._ChartType(doughnut)
                  .Attributes
                    .Name('DoughnutROsTotal')
                    .Labelling
                      .Format('0.00%')// Consultar em http://numeraljs.com/#format
                      .RGBColor('30,30,30') // Cor RGB separado por Virgula
                      .FontSize(20) // Tamanho da Fonte
                      .FontStyle('normal') // normal, bold, italic
                      .FontFamily(fontFamily) // Open Sans, Arial, Helvetica e etc..
                      .Padding(0) // Numeros negativos e positivos
                      .PaddingX(0)
                    .&End
                    .ColSpan(12)
                    .Heigth(130)
                    .Options
                      .Legend
                        .display(true)
                        .position('bottom')
                        .Labels
                          .padding(30)
                          .fontSize(16)
                          .fontColorHEX('#000000')
                          .fontFamily(fontFamily)
                        .&End
                      .&End
                      .Title
                        .FontSize(18)
                        .fontColorHEX('#000000')
                        .text('TOTAL DE LEADS RECEBIDOS')
                        .display(true)
                        .position('top')
                        .fontFamily(fontFamily)
                        .fontStyle('normal')
                        .padding(20)
                      .&End
                      .Tooltip
                        .DisplayTitle(true)
                        .Format('0.00%')
                        .ToolTipNoScales
                        .Intersect(true)
                      .&End
                    .&End
                    .DataSet
                      .textLabel('R.O.s Total por Período %')
                      .DataSet(qryGeral3)
                      .BackgroundOpacity(8)
                    .&End
                  .&End
                .&End
              .&End
            .HTML
            )
        .&End
      .&End

      .Jumpline

      .Rows
        ._Div
          .ColSpan(1)
        .&End
        ._Div
          .ColSpan(10)
          .Add(
              headerCharts.HTML)
        .&End
        ._Div
          .ColSpan(1)
        .&End
      .&End

      .Jumpline
      .Jumpline

      .Rows
        .ID('ROTable')
        .HTML(montaTableOPRO)
      .&End


    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .CallbackJS
      .ClassProvider(Self)
    .&End
    .Generated;

  for i := 0 to tamanhoQry - 1 do
  begin
    FreeAndNil(arrayQry[i]);
  end;

  form_esmaecer_fundo.Hide;

end;

procedure TunitHome.AttGraphFour;
begin

end;

procedure TunitHome.GraphFiveI(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: String;
  tableColumnCount, tableRecordCount: Integer;
  I, J: Integer;
  montaTableVendas: String;
  montaSQL, montaSQLI: String;
  idOrigem, nomeOrigem: String;
  qntdPesOrigens: Integer;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  navBarFunction := MontaNavBar_Marketing('GraphFive', dateStart, dateFinish, false, '', '', true, campanhasID);

  LimpaQuery(qryGeral);
  qryGeral.Open('SELECT GruposResumoOperacoes_ID ID, ReGrupoNome NOME FROM GruposResumoOperacoes WHERE ReAtivo = 1 ORDER BY NOME');

  montaSQL := ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_Marketing'') ' +
  ' DROP TABLE TempDash_Marketing ' +
  '  ' +
  ' create table TempDash_Marketing ' +
  ' (id int IDENTITY(1,1) NOT NULL, origem varchar(100), qntdOrigens varchar(10) ';

  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' VARCHAR(10) ';
    qryGeral.Next;
  End;


  montaSQL := montaSQL + ', rgb VARCHAR(15) ) ';

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(montaSQL);

  LimpaQuery(qryGeral1);
  qryGeral1.Open(
  ' SELECT DISTINCT DBO.RETORNAORNOME(PesOrigensID) ORIGEM, PesOrigensID ID, COUNT(*) QNTD ' +
  ' FROM Pessoas WITH(NOLOCK) ' +
  ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
  ' WHERE CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' AND ISNULL(PesOrigensID,0) <> 0 ' +
  ' GROUP BY PesOrigensID ' +
  ' ORDER BY ORIGEM ');

  montaSQL := '';
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    idOrigem := qryGeral1.FieldByName('ID').AsString;
    nomeOrigem := qryGeral1.FieldByName('ORIGEM').AsString;
    qntdPesOrigens := qryGeral1.FieldByName('QNTD').AsInteger;

    montaSQL := ' INSERT INTO TempDash_Marketing(origem, qntdOrigens';

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    Begin
      montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
      qryGeral.Next;
    End;


    montaSQL := montaSQL + ' , rgb) ';
    montaSQL := montaSQL + ' VALUES(' + QuotedStr(nomeOrigem) + ', ' + qntdPesOrigens.ToString;

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    begin
      montaSQL := montaSQL + ', (';
      montaSQL := montaSQL +
      ' SELECT COUNT(DISTINCT CoFoneListsID) ' +
      ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
      ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID' +
      ' INNER JOIN Pessoas WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
      ' WHERE ISNULL(PesOrigensID,0) = ' + idOrigem + ' AND CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' ' +
      ' AND ISNULL(CoResumoOperacaoID,0) IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) ' +
      ' INNER JOIN GruposResumoOperacoes WITH(NOLOCK) ON ReGrupoResumoOperacoesID = GruposResumoOperacoes_ID ' +
      ' WHERE ReGrupoResumoOperacoesID IN (' + qryGeral.FieldByName('ID').AsString + ') AND ReAtivo = 1) ';
      montaSQL := montaSQL + ' ) ';

      qryGeral.Next;
    end;
    montaSQL := montaSQL + ', (CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1))) ) ';

    LimpaQuery(qryGeral2);
    qryGeral2.ExecSQL(montaSQL);

    qryGeral1.Next;
  End;


  montaSQL := 'SELECT X.Origem [Origem], X.qntdOrigens [Quantidade Origens] ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' [' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ' FROM ( ';

  montaSQL := montaSQL + 'SELECT Origem, qntdOrigens ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')) / CONVERT(FLOAT, qntdOrigens)), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 0 AS UNIONORDER FROM TempDash_Marketing UNION ';

  montaSQL := montaSQL + 'SELECT ''TOTAL'', SUM(CONVERT(INT, qntdOrigens)) ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', SUM(CONVERT(INT, grRO' + qryGeral.FieldByName('ID').AsString + ')) ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, SUM(CONVERT(INT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')))) / CONVERT(FLOAT, SUM(CONVERT(INT, qntdOrigens)))), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 1 AS UNIONORDER FROM TempDash_Marketing';
  montaSQL := montaSQL + ' ) X ';

  montaSQL := montaSQL + 'GROUP BY X.Origem, X.qntdOrigens ';
  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', X.UNIONORDER ORDER BY X.UNIONORDER';

  LimpaQuery(qryGeral1);
  qryGeral1.Open(montaSQL);

  tableRecordCount := qryGeral1.RecordCount;
  tableColumnCount := qryGeral1.FieldCount;

  montaTableVendas :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 12px; width: 50%; margin-left: 10px; margin-right: 10px;"> ' +
    '   <thead> ' +
    '     <tr> ';

  for I := 0 to qryGeral1.FieldCount - 1 do
    montaTableVendas := montaTableVendas + ' <th scope="col">' + qryGeral1.Fields[I].FieldName + '</th> ';

  montaTableVendas := montaTableVendas +
    '     </tr> ' +
    '   </thead> ' +
    '   <tbody> ' ;

  qryGeral1.First;
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    montaTableVendas := montaTableVendas + ' <tr> ';
    for J := 0 to qryGeral1.FieldCount - 1 do
    Begin
      if J = 0 then
        montaTableVendas := montaTableVendas + ' <th scope="col"> ' + qryGeral1.Fields[J].AsString + ' </th> '
      else
        montaTableVendas := montaTableVendas + ' <td> ' + qryGeral1.Fields[J].AsString + ' </td> ';
    End;
    montaTableVendas := montaTableVendas + ' </tr> ';
    qryGeral1.Next;
  End;

  montaTableVendas := montaTableVendas + ' </tbody> ';
  montaTableVendas := montaTableVendas + ' </table> ';

  LimpaQuery(qryGeral);
  qryGeral.Open(
  ' SELECT Origem Label, CONVERT(DECIMAL(18, 5), (CONVERT(FLOAT, qntdOrigens) * 100) ' +
  '   / (SELECT SUM(CONVERT(INT, qntdOrigens)) FROM TempDash_Marketing)) / 100 Value, ' +
  ' RGB ' +
  ' FROM TempDash_Marketing ' +
  ' ORDER BY Value DESC ');

  WebCharts1
  .BackgroundColor('#FFFFFF')
  .FontColor('#000000')
  .Container(fluid)
  .AddResource('<link href="css/green.css" rel="stylesheet">')
  .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
  .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
  .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
  .NewProject

   .Rows
    .HTML(navBarFunction)
   .&End

   .Container(true)

     .Jumpline

     .Rows
      .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Vendas por Origem</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Por Data de Cadastro da Pessoa & Contatos Feitos ' +
        ' </figcaption> ' +
      ' </figure> ')
     .&End

     .Jumpline

     .Rows
      .ID('infoDate')
      .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> '
      )
     .&End

     .Jumpline

    .Rows
      .HTML(montaTableVendas)
    .&End

    .Jumpline

    .Rows
      ._Div
        .ColSpan(12)
        .Add(WebCharts1
              .ContinuosProject
                .Charts
                  ._ChartType(pie)
                    .Attributes
                      .Labelling
                        .Format('0.00%')// Consultar em http://numeraljs.com/#format
                        .RGBColor('255,255,255') // Cor RGB separado por Virgula
                        .FontSize(20) // Tamanho da Fonte
                        .FontStyle('normal') // normal, bold, italic
                        .FontFamily('Open Sans') // Open Sans, Arial, Helvetica e etc..
                        .Padding(0) // Numeros negativos e positivos
                        .PaddingX(0)
                      .&End
                      .Name('PIE_Origens_Cadastro')
                      .ColSpan(10)
                      .Heigth(200)
                      .Options
                        .Legend
                          .display(true)
                          .position('right')
                          .Labels
                            .fontSize(16)
                            .fontColorHEX('#000000')
                            .fontFamily('Open Sans')
                          .&End
                        .&End
                        .Title
                          .FontSize(18)
                          .fontColorHEX('#000000')
                          .text('Leads Dividos por Origens %')
                          .display(true)
                          .position('top')
                          .fontFamily('Product Sans')
                          .fontStyle('normal')
                          .padding(20)
                        .&End
                        .Tooltip
                          .Format('0.00%')
                        .&End
                      .&End
                      .DataSet
                        .DataSet(qryGeral)
                      .&End
                    .&End
                  .&End
                .&End
              .HTML)
      .&End
    .&End


  .WindowParent(FMXWindowParent)
  .WebBrowser(FMXChromium1)
  .CallbackJS
    .ClassProvider(Self)
  .&End
  .Generated;

  form_esmaecer_fundo.Hide;

end;

procedure TunitHome.GraphFive(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: String;
  tableColumnCount, tableRecordCount: Integer;
  I, J: Integer;
  montaTableVendas: String;
  montaSQL, montaSQLI: String;
  idOrigem, nomeOrigem: String;
  qntdPesOrigens: Integer;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  navBarFunction := MontaNavBar_Marketing('GraphFive', dateStart, dateFinish, false, '', '', true, campanhasID);

  LimpaQuery(qryGeral);
  qryGeral.Open('SELECT GruposResumoOperacoes_ID ID, ReGrupoNome NOME FROM GruposResumoOperacoes WHERE ReAtivo = 1 ORDER BY NOME');

  montaSQL := ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_Marketing'') ' +
  ' DROP TABLE TempDash_Marketing ' +
  '  ' +
  ' create table TempDash_Marketing ' +
  ' (id int IDENTITY(1,1) NOT NULL, origem varchar(100), qntdOrigens varchar(10) ';

  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' VARCHAR(10) ';
    qryGeral.Next;
  End;


  montaSQL := montaSQL + ', rgb VARCHAR(15) ) ';

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(montaSQL);

  LimpaQuery(qryGeral1);
  qryGeral1.Open(
  ' SELECT DISTINCT DBO.RETORNAORNOME(PesOrigensID) ORIGEM, PesOrigensID ID, COUNT(*) QNTD ' +
  ' FROM Pessoas WITH(NOLOCK) ' +
  ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
  ' WHERE CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' AND ISNULL(PesOrigensID,0) <> 0 ' +
  ' GROUP BY PesOrigensID ' +
  ' ORDER BY ORIGEM ');

  montaSQL := '';
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    idOrigem := qryGeral1.FieldByName('ID').AsString;
    nomeOrigem := qryGeral1.FieldByName('ORIGEM').AsString;
    qntdPesOrigens := qryGeral1.FieldByName('QNTD').AsInteger;

    montaSQL := ' INSERT INTO TempDash_Marketing(origem, qntdOrigens';

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    Begin
      montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
      qryGeral.Next;
    End;


    montaSQL := montaSQL + ' , rgb) ';
    montaSQL := montaSQL + ' VALUES(' + QuotedStr(nomeOrigem) + ', ' + qntdPesOrigens.ToString;

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    begin
      montaSQL := montaSQL + ', (';
      montaSQL := montaSQL +
      ' SELECT COUNT(DISTINCT FN_PessoasID) ' +
      ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
      ' INNER JOIN Pessoas WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
      ' WHERE ISNULL(PesOrigensID,0) = ' + idOrigem + ' AND CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' ' +
      ' AND ISNULL(FN_UltimoROID,0) IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) ' +
      ' INNER JOIN GruposResumoOperacoes WITH(NOLOCK) ON ReGrupoResumoOperacoesID = GruposResumoOperacoes_ID ' +
      ' WHERE ReGrupoResumoOperacoesID IN (' + qryGeral.FieldByName('ID').AsString + ') AND ReAtivo = 1) ';
      montaSQL := montaSQL + ' ) ';

      qryGeral.Next;
    end;
    montaSQL := montaSQL + ', (CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1))) ) ';

    LimpaQuery(qryGeral2);
    qryGeral2.ExecSQL(montaSQL);

    qryGeral1.Next;
  End;


  montaSQL := 'SELECT X.Origem [Origem], X.qntdOrigens [Quantidade Origens] ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' [' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ' FROM ( ';

  montaSQL := montaSQL + 'SELECT Origem, qntdOrigens ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')) / CONVERT(FLOAT, qntdOrigens)), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 0 AS UNIONORDER FROM TempDash_Marketing UNION ';

  montaSQL := montaSQL + 'SELECT ''TOTAL'', SUM(CONVERT(INT, qntdOrigens)) ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', SUM(CONVERT(INT, grRO' + qryGeral.FieldByName('ID').AsString + ')) ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, SUM(CONVERT(INT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')))) / CONVERT(FLOAT, SUM(CONVERT(INT, qntdOrigens)))), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 1 AS UNIONORDER FROM TempDash_Marketing';
  montaSQL := montaSQL + ' ) X ';

  montaSQL := montaSQL + 'GROUP BY X.Origem, X.qntdOrigens ';
  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', X.UNIONORDER ORDER BY X.UNIONORDER';

  LimpaQuery(qryGeral1);
  qryGeral1.Open(montaSQL);

  tableRecordCount := qryGeral1.RecordCount;
  tableColumnCount := qryGeral1.FieldCount;

  montaTableVendas :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 12px; width: 50%; margin-left: 10px; margin-right: 10px;"> ' +
    '   <thead> ' +
    '     <tr> ';

  for I := 0 to qryGeral1.FieldCount - 1 do
    montaTableVendas := montaTableVendas + ' <th scope="col">' + qryGeral1.Fields[I].FieldName + '</th> ';

  montaTableVendas := montaTableVendas +
    '     </tr> ' +
    '   </thead> ' +
    '   <tbody> ' ;

  qryGeral1.First;
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    montaTableVendas := montaTableVendas + ' <tr> ';
    for J := 0 to qryGeral1.FieldCount - 1 do
    Begin
      if J = 0 then
        montaTableVendas := montaTableVendas + ' <th scope="col"> ' + qryGeral1.Fields[J].AsString + ' </th> '
      else
        montaTableVendas := montaTableVendas + ' <td> ' + qryGeral1.Fields[J].AsString + ' </td> ';
    End;
    montaTableVendas := montaTableVendas + ' </tr> ';
    qryGeral1.Next;
  End;

  montaTableVendas := montaTableVendas + ' </tbody> ';
  montaTableVendas := montaTableVendas + ' </table> ';

  LimpaQuery(qryGeral);
  qryGeral.Open(
  ' SELECT Origem Label, CONVERT(DECIMAL(18, 5), (CONVERT(FLOAT, qntdOrigens) * 100) ' +
  '   / (SELECT SUM(CONVERT(INT, qntdOrigens)) FROM TempDash_Marketing)) / 100 Value, ' +
  ' RGB ' +
  ' FROM TempDash_Marketing ' +
  ' ORDER BY Value DESC ');

  WebCharts1
  .BackgroundColor('#FFFFFF')
  .FontColor('#000000')
  .Container(fluid)
  .AddResource('<link href="css/green.css" rel="stylesheet">')
  .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
  .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
  .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
  .NewProject

   .Rows
    .HTML(navBarFunction)
   .&End

   .Container(true)

     .Jumpline

     .Rows
      .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Vendas por Origem</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Por Data de Cadastro da Pessoa & Último R.O. ' +
        ' </figcaption> ' +
      ' </figure> ')
     .&End

     .Jumpline

     .Rows
      .ID('infoDate')
      .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> '
      )
     .&End

     .Jumpline

    .Rows
      .HTML(montaTableVendas)
    .&End

    .Jumpline

    .Rows
      ._Div
        .ColSpan(12)
        .Add(WebCharts1
              .ContinuosProject
                .Charts
                  ._ChartType(pie)
                    .Attributes
                      .Labelling
                        .Format('0.00%')// Consultar em http://numeraljs.com/#format
                        .RGBColor('255,255,255') // Cor RGB separado por Virgula
                        .FontSize(20) // Tamanho da Fonte
                        .FontStyle('normal') // normal, bold, italic
                        .FontFamily('Open Sans') // Open Sans, Arial, Helvetica e etc..
                        .Padding(0) // Numeros negativos e positivos
                        .PaddingX(0)
                      .&End
                      .Name('PIE_Origens_Cadastro')
                      .ColSpan(10)
                      .Heigth(200)
                      .Options
                        .Legend
                          .display(true)
                          .position('right')
                          .Labels
                            .fontSize(16)
                            .fontColorHEX('#000000')
                            .fontFamily('Open Sans')
                          .&End
                        .&End
                        .Title
                          .FontSize(18)
                          .fontColorHEX('#000000')
                          .text('Leads Dividos por Origens %')
                          .display(true)
                          .position('top')
                          .fontFamily('Product Sans')
                          .fontStyle('normal')
                          .padding(20)
                        .&End
                        .Tooltip
                          .Format('0.00%')
                        .&End
                      .&End
                      .DataSet
                        .DataSet(qryGeral)
                      .&End
                    .&End
                  .&End
                .&End
              .HTML)
      .&End
    .&End


  .WindowParent(FMXWindowParent)
  .WebBrowser(FMXChromium1)
  .CallbackJS
    .ClassProvider(Self)
  .&End
  .Generated;

  form_esmaecer_fundo.Hide;

end;

procedure TunitHome.AttGraphFive(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: String;
  tableColumnCount, tableRecordCount: Integer;
  I, J: Integer;
  montaTableVendas: String;
  montaSQL, montaSQLI: String;
  idOrigem, nomeOrigem: String;
  qntdPesOrigens: Integer;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  navBarFunction := MontaNavBar_Marketing('GraphFive', dateStart, dateFinish, false, '', '', true, campanhasID);

  LimpaQuery(qryGeral);
  qryGeral.Open('SELECT GruposResumoOperacoes_ID ID, ReGrupoNome NOME FROM GruposResumoOperacoes WHERE ReAtivo = 1 ORDER BY NOME');

  montaSQL := ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_Marketing'') ' +
  ' DROP TABLE TempDash_Marketing ' +
  '  ' +
  ' create table TempDash_Marketing ' +
  ' (id int IDENTITY(1,1) NOT NULL, origem varchar(100), qntdOrigens varchar(10) ';

  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' VARCHAR(10) ';
    qryGeral.Next;
  End;


  montaSQL := montaSQL + ', rgb VARCHAR(15) ) ';

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(montaSQL);

  LimpaQuery(qryGeral1);
  qryGeral1.Open(
  ' SELECT DISTINCT EnNome ORIGEM, EndRegioes_ID ID, COUNT(*) QNTD ' +
  ' FROM Pessoas WITH(NOLOCK) ' +
  ' INNER JOIN EndRegioes WITH(NOLOCK) ON PesEndRegioesID = EndRegioes_ID ' +
  ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
  ' WHERE CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' AND ISNULL(PesEndRegioesID,0) <> 0 ' +
  ' GROUP BY EnNome, EndRegioes_ID ' +
  ' ORDER BY ORIGEM ');

  montaSQL := '';
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    idOrigem := qryGeral1.FieldByName('ID').AsString;
    nomeOrigem := qryGeral1.FieldByName('ORIGEM').AsString;
    qntdPesOrigens := qryGeral1.FieldByName('QNTD').AsInteger;

    montaSQL := ' INSERT INTO TempDash_Marketing(origem, qntdOrigens';

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    Begin
      montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
      qryGeral.Next;
    End;


    montaSQL := montaSQL + ' , rgb) ';
    montaSQL := montaSQL + ' VALUES(' + QuotedStr(nomeOrigem) + ', ' + qntdPesOrigens.ToString;

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    begin
      montaSQL := montaSQL + ', (';
      montaSQL := montaSQL +
      ' SELECT COUNT(DISTINCT FN_PessoasID) ' +
      ' FROM Fone_List_' + caCodigo + ' WITH(NOLOCK) ' +
      ' INNER JOIN Pessoas WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
      ' WHERE ISNULL(PesEndRegioesID,0) = ' + idOrigem + ' AND CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' ' +
      ' AND ISNULL(FN_UltimoROID,0) IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID IN ('
      + qryGeral.FieldByName('ID').AsString + ')) ';
      montaSQL := montaSQL + ' ) ';

      qryGeral.Next;
    end;
    montaSQL := montaSQL + ', (CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1))) ) ';

    LimpaQuery(qryGeral2);
    qryGeral2.ExecSQL(montaSQL);

    qryGeral1.Next;
  End;

  montaSQL := 'SELECT X.Origem [Origem], X.qntdOrigens [Quantidade Origens] ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' [' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ' FROM ( ';

  montaSQL := montaSQL + 'SELECT Origem, qntdOrigens ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')) / CONVERT(FLOAT, qntdOrigens)), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 0 AS UNIONORDER FROM TempDash_Marketing UNION ';

  montaSQL := montaSQL + 'SELECT ''TOTAL'', SUM(CONVERT(INT, qntdOrigens)) ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', SUM(CONVERT(INT, grRO' + qryGeral.FieldByName('ID').AsString + ')) ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, SUM(CONVERT(INT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')))) / CONVERT(FLOAT, SUM(CONVERT(INT, qntdOrigens)))), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 1 AS UNIONORDER FROM TempDash_Marketing';
  montaSQL := montaSQL + ' ) X ';

  montaSQL := montaSQL + 'GROUP BY X.Origem, X.qntdOrigens ';
  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', X.UNIONORDER ORDER BY X.UNIONORDER';

  LimpaQuery(qryGeral1);
  qryGeral1.Open(montaSQL);

  tableRecordCount := qryGeral1.RecordCount;
  tableColumnCount := qryGeral1.FieldCount;

  montaTableVendas :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; margin-left: 10%;"> ' +
    '   <thead> ' +
    '     <tr> ';

  for I := 0 to qryGeral1.FieldCount - 1 do
    montaTableVendas := montaTableVendas + ' <th scope="col">' + qryGeral1.Fields[I].FieldName + '</th> ';

  montaTableVendas := montaTableVendas +
    '     </tr> ' +
    '   </thead> ' +
    '   <tbody> ' ;

  qryGeral1.First;
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    montaTableVendas := montaTableVendas + ' <tr> ';
    for J := 0 to qryGeral1.FieldCount - 1 do
    Begin
      if J = 0 then
        montaTableVendas := montaTableVendas + ' <th scope="col"> ' + qryGeral1.Fields[J].AsString + ' </th> '
      else
        montaTableVendas := montaTableVendas + ' <td> ' + qryGeral1.Fields[J].AsString + ' </td> ';
    End;
    montaTableVendas := montaTableVendas + ' </tr> ';
    qryGeral1.Next;
  End;

  montaTableVendas := montaTableVendas + ' </tbody> ';
  montaTableVendas := montaTableVendas + ' </table> ';

  LimpaQuery(qryGeral);
  qryGeral.Open(
  ' SELECT Origem Label, CONVERT(DECIMAL(18, 5), (CONVERT(FLOAT, qntdOrigens) * 100) ' +
  '   / (SELECT SUM(CONVERT(INT, qntdOrigens)) FROM TempDash_Marketing)) / 100 Value, ' +
  ' RGB ' +
  ' FROM TempDash_Marketing ' +
  ' ORDER BY Value DESC ');

  WebCharts1
  .BackgroundColor('#FFFFFF')
  .FontColor('#000000')
  .Container(fluid)
  .AddResource('<link href="css/green.css" rel="stylesheet">')
  .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
  .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
  .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
  .NewProject

   .Rows
    .HTML(navBarFunction)
   .&End

   .Container(true)

     .Jumpline

     .Rows
      .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Vendas por Regiões</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Por Data de Cadastro da Pessoa & Último R.O. ' +
        ' </figcaption> ' +
      ' </figure> ')
     .&End

     .Jumpline

     .Rows
      .ID('infoDate')
      .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> '
      )
     .&End

     .Jumpline

    .Rows
      .HTML(montaTableVendas)
    .&End



    .Jumpline

    .Rows
      ._Div
        .ColSpan(12)
        .Add(WebCharts1
              .ContinuosProject
                .Charts
                  ._ChartType(pie)
                    .Attributes
                      .Labelling
                        .Format('0.00%')// Consultar em http://numeraljs.com/#format
                        .RGBColor('255,255,255') // Cor RGB separado por Virgula
                        .FontSize(20) // Tamanho da Fonte
                        .FontStyle('normal') // normal, bold, italic
                        .FontFamily('Open Sans') // Open Sans, Arial, Helvetica e etc..
                        .Padding(0) // Numeros negativos e positivos
                        .PaddingX(0)
                      .&End
                      .Name('PIE_Origens_Cadastro')
                      .ColSpan(10)
                      .Heigth(200)
                      .Options
                        .Legend
                          .display(true)
                          .position('right')
                          .Labels
                            .fontSize(16)
                            .fontColorHEX('#000000')
                            .fontFamily('Open Sans')
                          .&End
                        .&End
                        .Title
                          .FontSize(18)
                          .fontColorHEX('#000000')
                          .text('Leads Dividos por Origens %')
                          .display(true)
                          .position('top')
                          .fontFamily('Product Sans')
                          .fontStyle('normal')
                          .padding(20)
                        .&End
                        .Tooltip
                          .Format('0.00%')
                        .&End
                      .&End
                      .DataSet
                        .DataSet(qryGeral)
                      .&End
                    .&End
                  .&End
                .&End
              .HTML)
      .&End
    .&End


  .WindowParent(FMXWindowParent)
  .WebBrowser(FMXChromium1)
  .CallbackJS
    .ClassProvider(Self)
  .&End
  .Generated;

  form_esmaecer_fundo.Hide;

end;

procedure TunitHome.AttGraphFiveI(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: String;
  tableColumnCount, tableRecordCount: Integer;
  I, J: Integer;
  montaTableVendas: String;
  montaSQL, montaSQLI: String;
  idOrigem, nomeOrigem: String;
  qntdPesOrigens: Integer;
begin

  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  navBarFunction := MontaNavBar_Marketing('GraphFive', dateStart, dateFinish, false, '', '', true, campanhasID);

  LimpaQuery(qryGeral);
  qryGeral.Open('SELECT GruposResumoOperacoes_ID ID, ReGrupoNome NOME FROM GruposResumoOperacoes WHERE ReAtivo = 1 ORDER BY NOME');

  montaSQL := ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_Marketing'') ' +
  ' DROP TABLE TempDash_Marketing ' +
  '  ' +
  ' create table TempDash_Marketing ' +
  ' (id int IDENTITY(1,1) NOT NULL, origem varchar(100), qntdOrigens varchar(10) ';

  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' VARCHAR(10) ';
    qryGeral.Next;
  End;


  montaSQL := montaSQL + ', rgb VARCHAR(15) ) ';

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(montaSQL);

  LimpaQuery(qryGeral1);
  qryGeral1.Open(
  ' SELECT DISTINCT EnNome ORIGEM, EndRegioes_ID ID, COUNT(*) QNTD ' +
  ' FROM Pessoas WITH(NOLOCK) ' +
  ' INNER JOIN EndRegioes WITH(NOLOCK) ON PesEndRegioesID = EndRegioes_ID ' +
  ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
  ' WHERE CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' AND ISNULL(PesEndRegioesID,0) <> 0 ' +
  ' GROUP BY EnNome, EndRegioes_ID ' +
  ' ORDER BY ORIGEM ');

  montaSQL := '';
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    idOrigem := qryGeral1.FieldByName('ID').AsString;
    nomeOrigem := qryGeral1.FieldByName('ORIGEM').AsString;
    qntdPesOrigens := qryGeral1.FieldByName('QNTD').AsInteger;

    montaSQL := ' INSERT INTO TempDash_Marketing(origem, qntdOrigens';

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    Begin
      montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
      qryGeral.Next;
    End;


    montaSQL := montaSQL + ' , rgb) ';
    montaSQL := montaSQL + ' VALUES(' + QuotedStr(nomeOrigem) + ', ' + qntdPesOrigens.ToString;

    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    begin
      montaSQL := montaSQL + ', (';
      montaSQL := montaSQL +
      ' SELECT COUNT(DISTINCT CoFoneListsID) ' +
      ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
      ' INNER JOIN Fone_List_' + caCodigo + ' WITH(NOLOCK) ON CoFoneListsID = FN_FoneList_' + caCodigo + '_ID ' +
      ' INNER JOIN Pessoas WITH(NOLOCK) ON FN_PessoasID = Pessoas_ID ' +
      ' WHERE ISNULL(PesEndRegioesID,0) = ' + idOrigem + ' AND CONVERT(DATE, DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + ''' ' +
      ' AND ISNULL(CoResumoOperacaoID,0) IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID IN ('
      + qryGeral.FieldByName('ID').AsString + ')) ';
      montaSQL := montaSQL + ' ) ';

      qryGeral.Next;
    end;
    montaSQL := montaSQL + ', (CONCAT(CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1),'','',CONVERT(INT, RAND()*(255-1)+1))) ) ';

    LimpaQuery(qryGeral2);
    qryGeral2.ExecSQL(montaSQL);

    qryGeral1.Next;
  End;

  montaSQL := 'SELECT X.Origem [Origem], X.qntdOrigens [Quantidade Origens] ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' [' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ' FROM ( ';

  montaSQL := montaSQL + 'SELECT Origem, qntdOrigens ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')) / CONVERT(FLOAT, qntdOrigens)), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 0 AS UNIONORDER FROM TempDash_Marketing UNION ';

  montaSQL := montaSQL + 'SELECT ''TOTAL'', SUM(CONVERT(INT, qntdOrigens)) ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', SUM(CONVERT(INT, grRO' + qryGeral.FieldByName('ID').AsString + ')) ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', CONCAT(CONVERT(DECIMAL(18,2), (100 * CONVERT(FLOAT, SUM(CONVERT(INT, grRO'
      + qryGeral.FieldByName('ID').AsString + ')))) / CONVERT(FLOAT, SUM(CONVERT(INT, qntdOrigens)))), '' %'') [% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', 1 AS UNIONORDER FROM TempDash_Marketing';
  montaSQL := montaSQL + ' ) X ';

  montaSQL := montaSQL + 'GROUP BY X.Origem, X.qntdOrigens ';
  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.grRO' + qryGeral.FieldByName('ID').AsString + ' ';
    qryGeral.Next;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', X.[% ' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ', X.UNIONORDER ORDER BY X.UNIONORDER';

  LimpaQuery(qryGeral1);
  qryGeral1.Open(montaSQL);

  tableRecordCount := qryGeral1.RecordCount;
  tableColumnCount := qryGeral1.FieldCount;

  montaTableVendas :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 14px; width: 80%; margin-left: 10%;"> ' +
    '   <thead> ' +
    '     <tr> ';

  for I := 0 to qryGeral1.FieldCount - 1 do
    montaTableVendas := montaTableVendas + ' <th scope="col">' + qryGeral1.Fields[I].FieldName + '</th> ';

  montaTableVendas := montaTableVendas +
    '     </tr> ' +
    '   </thead> ' +
    '   <tbody> ' ;

  qryGeral1.First;
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    montaTableVendas := montaTableVendas + ' <tr> ';
    for J := 0 to qryGeral1.FieldCount - 1 do
    Begin
      if J = 0 then
        montaTableVendas := montaTableVendas + ' <th scope="col"> ' + qryGeral1.Fields[J].AsString + ' </th> '
      else
        montaTableVendas := montaTableVendas + ' <td> ' + qryGeral1.Fields[J].AsString + ' </td> ';
    End;
    montaTableVendas := montaTableVendas + ' </tr> ';
    qryGeral1.Next;
  End;

  montaTableVendas := montaTableVendas + ' </tbody> ';
  montaTableVendas := montaTableVendas + ' </table> ';

  LimpaQuery(qryGeral);
  qryGeral.Open(
  ' SELECT Origem Label, CONVERT(DECIMAL(18, 5), (CONVERT(FLOAT, qntdOrigens) * 100) ' +
  '   / (SELECT SUM(CONVERT(INT, qntdOrigens)) FROM TempDash_Marketing)) / 100 Value, ' +
  ' RGB ' +
  ' FROM TempDash_Marketing ' +
  ' ORDER BY Value DESC ');

  WebCharts1
  .BackgroundColor('#FFFFFF')
  .FontColor('#000000')
  .Container(fluid)
  .AddResource('<link href="css/green.css" rel="stylesheet">')
  .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
  .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
  .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
  .NewProject

   .Rows
    .HTML(navBarFunction)
   .&End

   .Container(true)

     .Jumpline

     .Rows
      .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Vendas por Regiões</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Por Data de Cadastro da Pessoa & Contatos Feitos ' +
        ' </figcaption> ' +
      ' </figure> ')
     .&End

     .Jumpline

     .Rows
      .ID('infoDate')
      .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
          ' </blockquote> ' +
          ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
            ' Em ' + campanhaCodigo + ' - ' + campanhaNome +
          ' </figcaption> ' +
        ' </figure> '
      )
     .&End

     .Jumpline

    .Rows
      .HTML(montaTableVendas)
    .&End

    .Jumpline

    .Rows
      ._Div
        .ColSpan(12)
        .Add(WebCharts1
              .ContinuosProject
                .Charts
                  ._ChartType(pie)
                    .Attributes
                      .Labelling
                        .Format('0.00%')// Consultar em http://numeraljs.com/#format
                        .RGBColor('255,255,255') // Cor RGB separado por Virgula
                        .FontSize(20) // Tamanho da Fonte
                        .FontStyle('normal') // normal, bold, italic
                        .FontFamily('Open Sans') // Open Sans, Arial, Helvetica e etc..
                        .Padding(0) // Numeros negativos e positivos
                        .PaddingX(0)
                      .&End
                      .Name('PIE_Origens_Cadastro')
                      .ColSpan(10)
                      .Heigth(200)
                      .Options
                        .Legend
                          .display(true)
                          .position('right')
                          .Labels
                            .fontSize(16)
                            .fontColorHEX('#000000')
                            .fontFamily('Open Sans')
                          .&End
                        .&End
                        .Title
                          .FontSize(18)
                          .fontColorHEX('#000000')
                          .text('Leads Dividos por Origens %')
                          .display(true)
                          .position('top')
                          .fontFamily('Product Sans')
                          .fontStyle('normal')
                          .padding(20)
                        .&End
                        .Tooltip
                          .Format('0.00%')
                        .&End
                      .&End
                      .DataSet
                        .DataSet(qryGeral)
                      .&End
                    .&End
                  .&End
                .&End
              .HTML)
      .&End
    .&End


  .WindowParent(FMXWindowParent)
  .WebBrowser(FMXChromium1)
  .CallbackJS
    .ClassProvider(Self)
  .&End
  .Generated;

  form_esmaecer_fundo.Hide;

end;

procedure TunitHome.GraphSeven(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: String;
  tableColumnCount, tableRecordCount: Integer;
  I, J: Integer;
  montaTable, montaSlide, montaSlideI: String;
  montaSQL, montaSQLI: String;
  idOrigem, nomeOrigem: String;
  idUser, nomeUser: String;
  somaValores: Integer;

  equipeID: String;
  X: Integer;
begin
  form_esmaecer_fundo.Show;

  if (dateStart <> '') and (dateFinish <> '') then
  Begin

    if (dateInicial = dateStart) and (dateFinal = dateFinish) then
    //faz algo
    else
    Begin

      dateInicial := dateStart.Substring(8,2) + '/' + dateStart.Substring(5,2) + '/' + dateStart.Substring(0,4);
      dateFinal := dateFinish.Substring(8,2) + '/' + dateFinish.Substring(5,2) + '/' + dateFinish.Substring(0,4);

    End;

  End
  else
  Begin

    if (dateInicial = '') and (dateFinal = '') then
    Begin

      dateInicial := DateToStr(Now);
      dateFinal := DateToStr(Now);

    End;

  End;

  if caCodigo = '' then
  Begin

    if campanhaCodigo <> '' then
    begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 AND CaCodigo = ' + QuotedStr(campanhaCodigo) + ' ORDER BY CaCodigo ');
    end
    else
    Begin
      LimpaQuery(qryCamp);
      qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
      ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
      ' WHERE CaEncerrada = 0 ORDER BY CaCodigo ');
    End;

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  End
  else
  begin

    LimpaQuery(qryCamp);
    qryCamp.Open('SELECT TOP 1 CaCodigo, CaNome, Campanhas_ID FROM Campanhas WITH(NOLOCK) ' +
    ' Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID = ' + dmdados.GB_UsuarioID +
    ' WHERE CaCodigo = ' + QuotedStr(caCodigo));

    caCodigo := qryCamp.FieldByName('CaCodigo').AsString;
    campanhaCodigo := caCodigo;
    campanhaNome := qryCamp.FieldByName('CaNome').AsString;
    campanhasID := qryCamp.FieldByName('Campanhas_ID').AsString;

  end;

  if opID <> '' then
    equipeID := opID
  else
  Begin
    LimpaQuery(qryGeral);
    qryGeral.Open('SELECT Equipes_ID ID, EqNomeEquipe NOME FROM Equipes WITH(NOLOCK) WHERE ISNULL(EqAtiva,1) = 1');

    equipeID := qryGeral.FieldByName('ID').AsString;
  End;

  navBarFunction := MontaNavBar_Ranking('GraphSeven', dateStart, dateFinish, false, '', '', true, campanhasID);

  LimpaQuery(qryGeral);
  qryGeral.Open('SELECT GruposResumoOperacoes_ID ID, ReGrupoNome NOME FROM GruposResumoOperacoes WHERE ReAtivo = 1 ORDER BY NOME');

  montaSQL := ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_Ranking'') ' +
  ' DROP TABLE TempDash_Ranking ' +
  '  ' +
  ' create table TempDash_Ranking ' +
  ' (id int IDENTITY(1,1) NOT NULL, op varchar(100), rNovos INT, rRecicla INT, tNovos INT, tRecicla INT ';

  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' INT ';
    qryGeral.Next;
  End;

  montaSQL := montaSQL + ' ) ';

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(montaSQL);

  LimpaQuery(qryGeral1);
  qryGeral1.Open('SELECT Usuarios_ID ID, UsNome NOME FROM Usuarios WITH(NOLOCK) WHERE ISNULL(UsEquipesID,0) = ' + equipeID);

  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    montaSQL := montaSQL + ' INSERT INTO TempDash_Ranking(op, rNovos, rRecicla, tNovos, tRecicla ';
    qryGeral.First;
    for J := 0 to qryGeral.RecordCount - 1 do
    Begin
      montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' ';
      qryGeral.Next;
    End;
    montaSQL := montaSQL + ' ) ';

    idUser := qryGeral1.FieldByName('ID').AsString;
    nomeUser := qryGeral1.FieldByName('NOME').AsString;

    LimpaQuery(qryGeral2);
    qryGeral2.Open(' SELECT CaCodigo, Campanhas_ID ID ' +
    ' FROM UsuariosCampanhas WITH(NOLOCK) ' +
    ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
    ' WHERE Campanhas_ID IN (3, 29, 33, 36) AND UsUsuariosID = ' + idUser);

    montaSQL := montaSQL + ' VALUES(' + idUser + ' ';

    somaValores := 0;
    for J := 0 to qryGeral2.RecordCount - 1 do
    Begin
      LimpaQuery(qryGeral3);
      qryGeral3.Open(' SELECT COUNT(DISTINCT FN_PessoasID) QNTD FROM Fone_List_' + qryGeral2.FieldByName('CaCodigo').AsString +
        ' WITH(NOLOCK) WHERE CONVERT(DATE, FN_DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + '''  ' +
        ' AND ISNULL(FN_OperadorProprietarioID,0) = ' + idUser);

      somaValores := somaValores + qryGeral3.FieldByName('QNTD').AsInteger;

      qryGeral2.Next;
    End;
    qryGeral2.First;
    montaSQL := montaSQL + ', ' + somaValores.ToString;

    LimpaQuery(qryGeral2);
    qryGeral2.Open(' SELECT CaCodigo, Campanhas_ID ID ' +
    ' FROM UsuariosCampanhas WITH(NOLOCK) ' +
    ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
    ' WHERE Campanhas_ID IN (11, 30, 34, 37) AND UsUsuariosID = ' + idUser);

    somaValores := 0;
    for J := 0 to qryGeral2.RecordCount - 1 do
    Begin
      LimpaQuery(qryGeral3);
      qryGeral3.Open(' SELECT COUNT(DISTINCT FN_PessoasID) QNTD FROM Fone_List_' + qryGeral2.FieldByName('CaCodigo').AsString +
        ' WITH(NOLOCK) WHERE CONVERT(DATE, FN_DataCad) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + '''  ' +
        ' AND ISNULL(FN_OperadorProprietarioID,0) = ' + idUser);

      somaValores := somaValores + qryGeral3.FieldByName('QNTD').AsInteger;

      qryGeral2.Next;
    End;
    qryGeral2.First;
    montaSQL := montaSQL + ', ' + somaValores.ToString;

    LimpaQuery(qryGeral2);
    qryGeral2.Open(' SELECT CaCodigo, Campanhas_ID ID ' +
    ' FROM UsuariosCampanhas WITH(NOLOCK) ' +
    ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
    ' WHERE Campanhas_ID IN (3, 29, 33, 36) AND UsUsuariosID = ' + idUser);

    somaValores := 0;
    for J := 0 to qryGeral2.RecordCount - 1 do
    Begin
      LimpaQuery(qryGeral3);
      qryGeral3.Open(' SELECT COUNT(DISTINCT CoFoneListsID) QNTD FROM ContatosFichas_' + qryGeral2.FieldByName('CaCodigo').AsString +
        ' WITH(NOLOCK) WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + '''  ' +
        ' AND ISNULL(CoUsuariosID,0) = ' + idUser);

      somaValores := somaValores + qryGeral3.FieldByName('QNTD').AsInteger;

      qryGeral2.Next;
    End;
    qryGeral2.First;

    montaSQL := montaSQL + ', ' + somaValores.ToString;

    LimpaQuery(qryGeral2);
    qryGeral2.Open(' SELECT CaCodigo, Campanhas_ID ID ' +
    ' FROM UsuariosCampanhas WITH(NOLOCK) ' +
    ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
    ' WHERE Campanhas_ID IN (11, 30, 34, 37) AND UsUsuariosID = ' + idUser);

    somaValores := 0;
    for J := 0 to qryGeral2.RecordCount - 1 do
    Begin
      LimpaQuery(qryGeral3);
      qryGeral3.Open(' SELECT COUNT(DISTINCT CoFoneListsID) QNTD FROM ContatosFichas_' + qryGeral2.FieldByName('CaCodigo').AsString +
        ' WITH(NOLOCK) WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + '''  ' +
        ' AND ISNULL(CoUsuariosID,0) = ' + idUser);

      somaValores := somaValores + qryGeral3.FieldByName('QNTD').AsInteger;

      qryGeral2.Next;
    End;
    qryGeral2.First;

    montaSQL := montaSQL + ', ' + somaValores.ToString;

    LimpaQuery(qryGeral2);
    qryGeral2.Open(' SELECT CaCodigo, Campanhas_ID ID ' +
    ' FROM UsuariosCampanhas WITH(NOLOCK) ' +
    ' INNER JOIN Campanhas WITH(NOLOCK) ON UsCampanhasID = Campanhas_ID ' +
    ' WHERE UsUsuariosID = ' + idUser);

    somaValores := 0;

    qryGeral.First;
    for X := 0 to qryGeral.RecordCount - 1 do
    Begin
      somaValores := 0;
      qryGeral2.First;
      for J := 0 to qryGeral2.RecordCount - 1 do
      Begin
        LimpaQuery(qryGeral3);
        qryGeral3.Open(' SELECT COUNT(*) QNTD FROM ContatosFichas_' + qryGeral2.FieldByName('CaCodigo').AsString +
          ' WITH(NOLOCK) WHERE CONVERT(DATE, CoDataInicioFicha) BETWEEN ''' + dateInicial + ''' AND ''' + dateFinal + '''  ' +
          ' AND ISNULL(CoUsuariosID,0) = ' + idUser + ' AND CoResumoOperacaoID IN ( ' +
          '   SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID IN ('
          + qryGeral.FieldByName('ID').AsString + ' ) )' );

        somaValores := somaValores + qryGeral3.FieldByName('QNTD').AsInteger;

        qryGeral2.Next;
      End;

      montaSQL := montaSQL + ', ' + somaValores.ToString;
      qryGeral.Next;
    End;

    montaSQL := montaSQL + ' ) ';
    qryGeral1.Next;
  End;

  LimpaQuery(qryGeral1);
  qryGeral1.ExecSQL(montaSQL);

  montaSQL := 'SELECT CONCAT(CONVERT(VARCHAR(10), ROW_NUMBER() OVER(ORDER BY rNovos + tNovos DESC)), ''º'') [Ranking], ';
  montaSQL := montaSQL + 'DBO.RetornaNomeUsuario(OP) [Usuário], rNovos + tNovos [Novos], rRecicla + tRecicla [Reciclagem] ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  begin
    montaSQL := montaSQL + ', grRO' + qryGeral.FieldByName('ID').AsString + ' [' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  end;
  montaSQL := montaSQL + 'FROM TempDash_Ranking ';

  LimpaQuery(qryGeral1);
  qryGeral1.Open(montaSQL);

  montaTable :=
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 12px; width: 80%; margin-left: 10%;"> ' +
    '   <thead> ' +
    '     <tr> ';

  for I := 0 to qryGeral1.FieldCount - 1 do
  begin
    montaTable := montaTable + ' <th scope="col">' + qryGeral1.Fields[i].FieldName + ' </th> ';
  end;

  montaTable := montaTable +
    '     </tr> ' +
    '   </thead> ' +
    '   <tbody> ';

  qryGeral1.First;
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    montaTable := montaTable + ' <tr> ';
    for J := 0 to qryGeral1.FieldCount - 1 do
    begin
      montaTable := montaTable + ' <td> ' + qryGeral1.Fields[J].AsString + ' </td> ';
    end;
    montaTable := montaTable + ' </tr> ';
    qryGeral1.Next;
  End;

  montaTable := montaTable +
    '   </tbody> ';

  montaSQL := 'SELECT '''' [Ranking], ';
  montaSQL := montaSQL + ' ''Total'' [Usuário], SUM(rNovos + tNovos) [Novos], SUM(rRecicla + tRecicla) [Reciclagem] ';

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  begin
    montaSQL := montaSQL + ', SUM(grRO' + qryGeral.FieldByName('ID').AsString + ') [' + qryGeral.FieldByName('NOME').AsString + '] ';
    qryGeral.Next;
  end;
  montaSQL := montaSQL + 'FROM TempDash_Ranking ';

  LimpaQuery(qryGeral1);
  qryGeral1.Open(montaSQL);

  qryGeral1.First;
  for I := 0 to qryGeral1.RecordCount - 1 do
  Begin
    montaTable := montaTable + ' <tr> ';
    for J := 0 to qryGeral1.FieldCount - 1 do
    begin
      montaTable := montaTable + ' <th scope="col"> ' + qryGeral1.Fields[J].AsString + ' </th> ';
    end;
    montaTable := montaTable + ' </tr> ';
    qryGeral1.Next;
  End;

  montaTable := montaTable +
    ' </table> ';

  montaSlide :=
    ' <div id="carouselExampleCaptions" class="carousel carousel-dark slide" data-bs-ride="carousel"> ' +
    '   <div class="carousel-indicators"> ';

  montaSlideI := montaSlideI +
    ' <div class="carousel-inner"> ';

  qryGeral1.First;
  for I := 0 to qryGeral1.FieldCount - 1 do
  Begin
    LimpaQuery(qryGeral2);
    if I = 0 then
    Begin
      montaSlide := montaSlide +
      '   <button type="button" data-bs-target="#carouselExampleCaptions" data-bs-slide-to="'+ I.ToString +'" class="active" aria-current="true" aria-label="Slide '+ I.ToString +'"></button> ';

      qryGeral2.Open('SELECT CONCAT(CONVERT(VARCHAR(10), ROW_NUMBER() OVER(ORDER BY rNovos DESC)), ''º'') [Ranking], ' +
      ' DBO.RetornaNomeUsuario(OP) [Usuário], ' +
      ' rNovos [Novos], rRecicla [Reciclagem] ' +
      ' FROM TempDash_Ranking ');

      montaSlideI := montaSlideI +
        '   <div class="carousel-item active" style="padding-bottom: 50px"> ';

      montaSlideI := montaSlideI +
      ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 12px; width: 80%; margin-left: 10%;"> ';

      for J := 0 to qryGeral2.RecordCount - 1 do
      Begin

        if J = 0 then
        Begin
          montaSlideI := montaSlideI + ' <tr> ';
          for x := 0 to qryGeral2.FieldCount - 1 do
          Begin
            montaSlideI := montaSlideI + ' <th scope="col"> ' + qryGeral2.Fields[x].FieldName + ' </th> ';
          End;
          montaSlideI := montaSlideI + ' </tr> ';
        End;

        montaSlideI := montaSlideI + ' <tr> ';
        for x := 0 to qryGeral2.FieldCount - 1 do
        Begin
          montaSlideI := montaSlideI + ' <td> ' + qryGeral2.Fields[x].AsString + ' </td> ';
        End;
        montaSlideI := montaSlideI + ' </tr> ';

        qryGeral2.Next;
      End;

      montaSlideI := montaSlideI + ' </table> ';
      montaSlideI := montaSlideI +
        '   <div class="carousel-caption d-none d-md-block"> ' +
        '     <h5>Clientes Recebidos</h5> ' +
        '   </div> ' +
        ' </div> ';
    End
    else if I = 1 then
    Begin
     montaSlide := montaSlide +
      '   <button type="button" data-bs-target="#carouselExampleCaptions" data-bs-slide-to="'+ I.ToString +'" aria-label="Slide '+ I.ToString +'"></button> ';

     qryGeral2.Open('SELECT CONCAT(CONVERT(VARCHAR(10), ROW_NUMBER() OVER(ORDER BY rNovos DESC)), ''º'') [Ranking], ' +
      ' DBO.RetornaNomeUsuario(OP) [Usuário], ' +
      ' tNovos [Novos], tRecicla [Reciclagem] ' +
      ' FROM TempDash_Ranking ');

      montaSlideI := montaSlideI +
        '   <div class="carousel-item" style="padding-bottom: 50px"> ';

      montaSlideI := montaSlideI +
      ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 12px; width: 80%; margin-left: 10%;"> ';

      for J := 0 to qryGeral2.RecordCount - 1 do
      Begin

        if J = 0 then
        Begin
          montaSlideI := montaSlideI + ' <tr> ';
          for x := 0 to qryGeral2.FieldCount - 1 do
          Begin
            montaSlideI := montaSlideI + ' <th scope="col"> ' + qryGeral2.Fields[x].FieldName + ' </th> ';
          End;
          montaSlideI := montaSlideI + ' </tr> ';
        End;

        montaSlideI := montaSlideI + ' <tr> ';
        for x := 0 to qryGeral2.FieldCount - 1 do
        Begin
          montaSlideI := montaSlideI + ' <td> ' + qryGeral2.Fields[x].AsString + ' </td> ';
        End;
        montaSlideI := montaSlideI + ' </tr> ';

        qryGeral2.Next;
      End;

      montaSlideI := montaSlideI + ' </table> ';
      montaSlideI := montaSlideI +
        '   <div class="carousel-caption d-none d-md-block"> ' +
        '     <h5>Clientes Trabalhados</h5> ' +
        '   </div> ' +
        ' </div> ';
    End
    else
    Begin
      montaSlide := montaSlide +
      '   <button type="button" data-bs-target="#carouselExampleCaptions" data-bs-slide-to="'+ I.ToString +'" aria-label="Slide '+ I.ToString +'"></button> ';
    End;
  End;

  qryGeral.First;
  for I := 0 to qryGeral.RecordCount - 1 do
  Begin
    qryGeral2.Open('SELECT CONCAT(CONVERT(VARCHAR(10), ROW_NUMBER() OVER(ORDER BY rNovos DESC)), ''º'') [Ranking], ' +
      ' DBO.RetornaNomeUsuario(OP) [Usuário], ' +
      ' grRO' + qryGeral.FieldByName('ID').AsString + ' [' + qryGeral.FieldByName('NOME').AsString + '] ' +
      ' FROM TempDash_Ranking ');

    montaSlideI := montaSlideI +
      '   <div class="carousel-item" style="padding-bottom: 50px"> ';

    montaSlideI := montaSlideI +
    ' <table class="table table-sm table-borderless table-striped table-hover" style="font-size: 12px; width: 80%; margin-left: 10%;"> ';

    for J := 0 to qryGeral2.RecordCount - 1 do
    Begin

      if J = 0 then
      Begin
        montaSlideI := montaSlideI + ' <tr> ';
        for x := 0 to qryGeral2.FieldCount - 1 do
        Begin
          montaSlideI := montaSlideI + ' <th scope="col"> ' + qryGeral2.Fields[x].FieldName + ' </th> ';
        End;
        montaSlideI := montaSlideI + ' </tr> ';
      End;

      montaSlideI := montaSlideI + ' <tr> ';
      for x := 0 to qryGeral2.FieldCount - 1 do
      Begin
        montaSlideI := montaSlideI + ' <td> ' + qryGeral2.Fields[x].AsString + ' </td> ';
      End;
      montaSlideI := montaSlideI + ' </tr> ';

      qryGeral2.Next;
    End;

    montaSlideI := montaSlideI + ' </table> ';
    montaSlideI := montaSlideI +
      '   <div class="carousel-caption d-none d-md-block"> ' +
      '     <h5>' + qryGeral.FieldByName('NOME').AsString + '</h5> ' +
      '   </div> ' +
      ' </div> ';

    qryGeral.Next;
  End;

  montaSlideI := montaSlideI +
  ' </div> ' +
  '   <button class="carousel-control-prev" type="button" data-bs-target="#carouselExampleCaptions" data-bs-slide="prev"> ' +
  '     <span class="carousel-control-prev-icon" aria-hidden="true"></span> ' +
  '     <span class="visually-hidden">Previous</span> ' +
  '   </button> ' +
  '   <button class="carousel-control-next" type="button" data-bs-target="#carouselExampleCaptions" data-bs-slide="next"> ' +
  '     <span class="carousel-control-next-icon" aria-hidden="true"></span> ' +
  '     <span class="visually-hidden">Next</span> ' +
  '   </button> ' +
  ' </div> ';

  montaSlide := montaSlide + ' </div> ' + montaSlideI;


  WebCharts1
  .BackgroundColor('#FFFFFF')
  .FontColor('#000000')
  .Container(fluid)
  .AddResource('<link href="css/green.css" rel="stylesheet">')
  .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
  .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
  .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
  .NewProject

   .Rows
    .HTML(navBarFunction)
   .&End

   .Jumpline

   .Rows
    .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Ranking</p> ' +
        ' </blockquote> ' +
      ' </figure> ')
   .&End

   .Jumpline

   .Rows
    .ID('infoDate')
    .HTML(
        ' <figure class="text-center"> ' +
          ' <blockquote class="blockquote"> ' +
            ' <p style=" font-family: Product Sans;">Filtro de busca: ' + dateInicial + ' até ' + dateFinal + ' .</p> ' +
          ' </blockquote> ' +
        ' </figure> '
      )
    .&End

   .Jumpline

   .Rows
    .HTML(montaSlide)
   .&End

   .Jumpline

   .Rows
    .HTML(montaTable)
   .&End

   .WindowParent(FMXWindowParent)
   .WebBrowser(FMXChromium1)
   .CallbackJS
    .ClassProvider(Self)
   .&End
   .Generated;


  form_esmaecer_fundo.Hide;
end;

procedure TunitHome.GraphSix(caCodigo: String = ''; campanhaID: String = ''; dateStart: String = ''; dateFinish: String = '';
  monthStart: String = ''; monthFinish: String = ''; opID: String = '');
var
  navBarFunction: String;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  mesExt, diaSemanaExt: String;
  diaSemana, controlaMes: Integer;
  I, controlaID: Integer;
  x: Integer;
  rgbColor: String;
  montaSQL, montaSQLI, montaSQLII, montaCards, strHTML: String;
  dataCompleta: TDate;

  arrayQry: array of TFDQuery;

  headerCharts_1: iModelHTML;
  config_row_1_Charts: iModelHTMLChartsConfig;
  y: Integer;
Begin

  form_esmaecer_fundo.Show;

  mesInicial := DateToStr(Now).Substring(3,2);
  anoInicial := DateToStr(Now).Substring(6,4);

  //RETORNA NOME DO MÊS INICIAL
  Case StrToInt(mesInicial) Of
    01 : mesExt := 'Janeiro';
    02 : mesExt := 'Fevereiro';
    03 : mesExt := 'Março';
    04 : mesExt := 'Abril';
    05 : mesExt := 'Maio';
    06 : mesExt := 'Junho';
    07 : mesExt := 'Julho';
    08 : mesExt := 'Agosto';
    09 : mesExt := 'Setembro';
    10 : mesExt := 'Outubro';
    11 : mesExt := 'Novembro';
    12 : mesExt := 'Dezembro';
  end;

  navBarFunction := MontaNavBar('GraphSix', '', '', true, '', '', true, '', false, false, true);

  PreencheTableDash_Semanas;

  dataCompleta := StrToDate('01/' + mesInicial + '/' + anoInicial);

  for x := 0 to 2 do
  Begin

    mesFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2);
    anoFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(6,4);

    controlaID := 1;

    for I := 1 to 31 do
    Begin

      LimpaQuery(qryGeral);
      qryGeral.Open(
        ' SELECT DISTINCT DATEPART(WEEKDAY, DataCad) DiaSemana, DATENAME(WEEKDAY, DataCad) DiaSemanaExt ' +
        ' FROM PESSOAS WITH(NOLOCK) ' +
        ' WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '   AND DATEPART(YEAR, DataCad) = ' + anoFinal);

      diaSemanaExt := qryGeral.FieldByName('DiaSemanaExt').AsString;
      diaSemana := qryGeral.FieldByName('DiaSemana').AsInteger;

      if controlaID = 1 then
        controlaID := diaSemana;

      if controlaID <= 7 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id <= 7 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 8) and (controlaID <= 14) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 8 AND 14 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 15) and (controlaID <= 21) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 15 AND 21 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 22) and (controlaID <= 28) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 22 AND 28 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 29) and (controlaID <= 35) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 29 AND 35 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if controlaID >= 36 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id >= 36 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End;


      controlaID := controlaID + 1;

    End;

  End;

  montaSQL := '';

  headerCharts_1 := WebCharts1.ContinuosProject;

  config_row_1_Charts := headerCharts_1
            .Charts
              ._ChartType(line)
                .Attributes
                  .Labelling
                    .Format('0,0')// Consultar em http://numeraljs.com/#format
                    .RGBColor('30,30,30') // Cor RGB separado por Virgula
                    .FontSize(20) // Tamanho da Fonte
                    .FontStyle('normal') // normal, bold, italic
                    .FontFamily(fontFamily) // Open Sans, Arial, Helvetica e etc..
                    .Padding(0) // Numeros negativos e positivos
                    .PaddingX(0)
                  .&End
                  .Name('LINESemana')
                  .ColSpan(6)
                  .Heigth(130)
                  .Options
                    .Legend
                      .display(true)
                      .position('bottom')
                    .&End
                    .Tooltip
                      .DisplayTitle(true)
                      .Format('0,0')
                      .ToolTipNoScales
                      .Intersect(true)
                    .&End
                  .&End;


  SetLength(arrayQry, 3);
  for I := 2 downto 0 do
  begin
    controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - I)).Substring(3,2));

    //RETORNA NOME DO MÊS INICIAL
    Case controlaMes Of
      01 : mesExt := 'Janeiro';
      02 : mesExt := 'Fevereiro';
      03 : mesExt := 'Março';
      04 : mesExt := 'Abril';
      05 : mesExt := 'Maio';
      06 : mesExt := 'Junho';
      07 : mesExt := 'Julho';
      08 : mesExt := 'Agosto';
      09 : mesExt := 'Setembro';
      10 : mesExt := 'Outubro';
      11 : mesExt := 'Novembro';
      12 : mesExt := 'Dezembro';
    end;

    if I = 2 then
      rgbColor := '56, 56, 223'
    else if I = 1 then
      rgbColor := '223, 56, 56'
    else
      rgbColor := '56, 223, 56';


    montaSQL := montaSQL + ' SELECT ' + QuotedStr(mesExt) + ' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '
      + QuotedStr(rgbColor) + ' RGB FROM TempDash_Semanas ';

    if I <> 0 then
      montaSQL := montaSQL + ' UNION ';

    try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].Open(
        ' SELECT ''SEMANA 01'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id <= 7 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 02'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 8 AND 14 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 03'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 15 AND 21 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 04'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 22 AND 28 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 05'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 29 AND 35 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 06'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id >= 36 ');

      config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .BackgroundColor(rgbColor)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(rgbColor)
              .&End;
    finally

    end;

  end;

  config_row_1_Charts.&End
              .&End;

  LimpaQuery(qryGeral);
  qryGeral.Open(montaSQL);

  montaSQL := '';
  montaSQLII := '';
  strHTML := '';

  montaCards :=
    ' <div ' +
    '   class="row row-cols-1 row-cols-md-3 g-4" ' +
    '   style="width: 80%; margin-left: 10%; font-size: 13px" ' +
    ' > ';
  for I := 1 to 6 do
  Begin
    montaSQL := ' SELECT DiaSemana Dias ';

    montaSQLII := montaSQLII + 'SELECT ''SEMANA 0' + i.ToString + ''' Semanas ';

    for x := 2 downto 0 do
    begin
     controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2));

      //RETORNA NOME DO MÊS INICIAL
      Case controlaMes Of
        01 : mesExt := 'Janeiro';
        02 : mesExt := 'Fevereiro';
        03 : mesExt := 'Março';
        04 : mesExt := 'Abril';
        05 : mesExt := 'Maio';
        06 : mesExt := 'Junho';
        07 : mesExt := 'Julho';
        08 : mesExt := 'Agosto';
        09 : mesExt := 'Setembro';
        10 : mesExt := 'Outubro';
        11 : mesExt := 'Novembro';
        12 : mesExt := 'Dezembro';
      end;

      montaSQL := montaSQL + ', ISNULL(Mes' + x.ToString + ', '''') [' + mesExt.Substring(0, 3) + '] ';

      montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes' + x.ToString + ',''''))) [' + mesExt.Substring(0, 3) + '] ';

    end;

    montaSQL := montaSQL + ', CONVERT(INT, ISNULL(Mes0,0)) - CONVERT(INT, ISNULL(Mes1,0)) Divergência FROM TempDash_Semanas ';

    montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) Divergência FROM TempDash_Semanas ';

    case I of
      1: begin
        montaSQLI := ' WHERE id <= 7';
      end;
      2: begin
        montaSQLI := ' WHERE id BETWEEN 8 AND 14 '
      end;
      3: begin
        montaSQLI := ' WHERE id BETWEEN 15 AND 21 '
      end;
      4: begin
        montaSQLI := ' WHERE id BETWEEN 22 AND 28 '
      end;
      5: begin
        montaSQLI := ' WHERE id BETWEEN 29 AND 35 '
      end;
      6: begin
        montaSQLI := ' WHERE id >= 36 ';
      end;
    end;

    montaSQL := montaSQL + montaSQLI +
    ' UNION ' +
    ' SELECT ''Total'' Dias, CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes2,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes1,'''')))) , CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0, '''')) - CONVERT(INT, ISNULL(Mes1, '''')))) ' +
    ' FROM TempDash_Semanas ' + montaSQLI;

    montaSQLII := montaSQLII + montaSQLI + ' UNION ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(montaSQL);

    montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semana 0' + I.ToString +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral1.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral1.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral1.First;
    for x := 0 to qryGeral1.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral1.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral1.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral1.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral1.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';
  End;

  montaSQLII := montaSQLII +
  ' SELECT ''Total'', SUM(CONVERT(INT, ISNULL(Mes2,''''))), SUM(CONVERT(INT, ISNULL(Mes1,''''))), SUM(CONVERT(INT, ISNULL(Mes0,''''))), SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) ' +
  ' FROM TempDash_Semanas ';

  LimpaQuery(qryGeral2);
  qryGeral2.Open(montaSQLII);

  //PREENCHE O ULTIMO CARD COM TODAS AS SEMANAS
  montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semanas ' +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral2.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral2.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral2.First;
    for x := 0 to qryGeral2.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral2.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral2.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral2.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral2.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';

  montaCards := montaCards +
  ' </div> ';

  WebCharts1
  .BackgroundColor('#FFFFFF')
    .FontColor('#000000')
    .Container(fluid)
    .AddResource('<link href="css/green.css" rel="stylesheet">')
    .AddResource('<link href="css/custom.min.css" rel="stylesheet">')
    .AddResource('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">')
    .AddResource('<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>')
    .AddResource('<script src="/scripts/snippet-javascript-console.min.js?v=1"></script>')

    .NewProject

    .Rows
     .HTML(navBarFunction)
    .&End

    .Jumpline

    .Rows
      .ID('InfoGeral')
      .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Comparativos por Leads Cadastrados</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
          ' Em ' + mesExt + '/' + anoInicial +
        ' </figcaption> ' +
      ' </figure> ')
    .&End

    .Jumpline

    .Rows
      ._Div
        .ColSpan(1)
      .&End
      .Tag
        .Add(headerCharts_1.HTML)
      .&End
      .Tag
        .Add(WebCharts1
            .ContinuosProject
              .Charts
                ._ChartType(bar)
                  .Attributes
                    .Name('BARSemanas')
                    .Labelling
                      .Format('0,0')// Consultar em http://numeraljs.com/#format
                      .RGBColor('30,30,30') // Cor RGB separado por Virgula
                      .FontSize(20) // Tamanho da Fonte
                      .FontStyle('normal') // normal, bold, italic
                      .FontFamily(fontFamily) // Open Sans, Arial, Helvetica e etc..
                      .Padding(0) // Numeros negativos e positivos
                      .PaddingX(0)
                    .&End
                    .ColSpan(4)
                    .Heigth(210)
                    .Options
                      .Legend
                        .display(true)
                        .position('bottom')
                        .Labels
                          .padding(30)
                          .fontSize(16)
                          .fontColorHEX('#000000')
                          .fontFamily(fontFamily)
                        .&End
                      .&End
                      .Tooltip
                        .DisplayTitle(true)
                        .Format('0,0')
                        .ToolTipNoScales
                        .Intersect(true)
                      .&End
                    .&End
                    .DataSet
                      .DataSet(qryGeral)
                      .BackgroundOpacity(8)
                    .&End
                  .&End
                .&End
              .&End
            .HTML)
      .&End
      ._Div
        .ColSpan(1)
      .&End
    .&End

    .Jumpline

    .Rows
      .ID('CardsSemanas')
      .HTML(montaCards)
    .&End

    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .CallbackJS
      .ClassProvider(Self)
    .&End
    .Generated;


    for i := 0 to 2 do
    begin
      FreeAndNil(arrayQry[i]);
    end;

    form_esmaecer_fundo.Hide;

End;

procedure TunitHome.AtualizaCadastros(mes: String = '');
var
  navBarFunction: String;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  mesExt, diaSemanaExt: String;
  diaSemana, controlaMes: Integer;
  I, controlaID: Integer;
  x: Integer;
  rgbColor: String;
  montaSQL, montaSQLI, montaSQLII, montaCards, strHTML: String;
  dataCompleta: TDate;

  arrayQry: array of TFDQuery;

  headerCharts_1: iModelHTML;
  config_row_1_Charts: iModelHTMLChartsConfig;
  y: Integer;
Begin

  form_esmaecer_fundo.Show;

  mesInicial := mes.Substring(5,2);
  anoInicial := mes.Substring(0,4);

  dataCompleta := StrToDate('01/' + mesInicial + '/' + anoInicial);

  //RETORNA NOME DO MÊS INICIAL
  Case StrToInt(mesInicial) Of
    01 : mesExt := 'Janeiro';
    02 : mesExt := 'Fevereiro';
    03 : mesExt := 'Março';
    04 : mesExt := 'Abril';
    05 : mesExt := 'Maio';
    06 : mesExt := 'Junho';
    07 : mesExt := 'Julho';
    08 : mesExt := 'Agosto';
    09 : mesExt := 'Setembro';
    10 : mesExt := 'Outubro';
    11 : mesExt := 'Novembro';
    12 : mesExt := 'Dezembro';
  end;

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
    .ID('InfoGeral')
    .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Comparativos por Leads Cadastrados</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
          ' Em ' + mesExt + '/' + anoInicial +
        ' </figcaption> ' +
      ' </figure> ')
    .Update;


  PreencheTableDash_Semanas;

  for x := 0 to 2 do
  Begin

    controlaID := 1;

    mesFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2);
    anoFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(6,4);

    for I := 1 to 31 do
    Begin

      LimpaQuery(qryGeral);
      qryGeral.Open(
        ' SELECT DISTINCT DATEPART(WEEKDAY, DataCad) DiaSemana, DATENAME(WEEKDAY, DataCad) DiaSemanaExt ' +
        ' FROM PESSOAS WITH(NOLOCK) ' +
        ' WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '   AND DATEPART(YEAR, DataCad) = ' + anoFinal);

      diaSemanaExt := qryGeral.FieldByName('DiaSemanaExt').AsString;
      diaSemana := qryGeral.FieldByName('DiaSemana').AsInteger;

      if controlaID = 1 then
        controlaID := diaSemana;

      if controlaID <= 7 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id <= 7 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 8) and (controlaID <= 14) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 8 AND 14 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 15) and (controlaID <= 21) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 15 AND 21 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 22) and (controlaID <= 28) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 22 AND 28 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 29) and (controlaID <= 35) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id BETWEEN 29 AND 35 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End
      else if controlaID >= 36 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ' +
        ' WHERE id >= 36 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, DataCad) ' +
        ' 	FROM PESSOAS WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, DataCad) = ' + I.ToString + ' AND DATEPART(MONTH, DataCad) = ' + mesFinal +
        '     AND DATEPART(YEAR, DataCad) = ' + anoFinal + ') ')

      End;


      controlaID := controlaID + 1;

    End;

  End;

  headerCharts_1 := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts_1
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINESemana')
                  .ColSpan(6)
                  .Heigth(130);


  SetLength(arrayQry, 3);
  for I := 2 downto 0 do
  begin

    controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - I)).Substring(3,2));

    //RETORNA NOME DO MÊS INICIAL
    Case controlaMes Of
      01 : mesExt := 'Janeiro';
      02 : mesExt := 'Fevereiro';
      03 : mesExt := 'Março';
      04 : mesExt := 'Abril';
      05 : mesExt := 'Maio';
      06 : mesExt := 'Junho';
      07 : mesExt := 'Julho';
      08 : mesExt := 'Agosto';
      09 : mesExt := 'Setembro';
      10 : mesExt := 'Outubro';
      11 : mesExt := 'Novembro';
      12 : mesExt := 'Dezembro';
    end;

    if I = 2 then
      rgbColor := '56, 56, 223'
    else if I = 1 then
      rgbColor := '223, 56, 56'
    else
      rgbColor := '56, 223, 56';


    montaSQL := montaSQL + ' SELECT ' + QuotedStr(mesExt) + ' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '
      + QuotedStr(rgbColor) + ' RGB FROM TempDash_Semanas ';

    if I <> 0 then
      montaSQL := montaSQL + ' UNION ';

    try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].Open(
        ' SELECT ''SEMANA 01'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id <= 7 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 02'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 8 AND 14 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 03'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 15 AND 21 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 04'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 22 AND 28 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 05'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 29 AND 35 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 06'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id >= 36 ');

        config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(mesExt)
                .BackgroundColor(rgbColor)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(rgbColor)
              .&End;

    finally

    end;

  end;

  config_row_1_Charts.&End
              .UpdateChart;

  LimpaQuery(qryGeral);
  qryGeral.Open(montaSQL);

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BARSemanas')
          .ColSpan(4)
          .Heigth(210)
          .DataSet
            .DataSet(qryGeral)
            .BackgroundOpacity(8)
          .&End
        .&End
      .UpdateChart;


  montaSQL := '';
  montaSQLII := '';
  strHTML := '';

  montaCards :=
    ' <div ' +
    '   class="row row-cols-1 row-cols-md-3 g-4" ' +
    '   style="width: 80%; margin-left: 10%; font-size: 13px" ' +
    ' > ';
  for I := 1 to 6 do
  Begin
    montaSQL := ' SELECT DiaSemana Dias ';

    montaSQLII := montaSQLII + 'SELECT ''SEMANA 0' + i.ToString + ''' Semanas ';

    for x := 2 downto 0 do
    begin
     controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2));

      //RETORNA NOME DO MÊS INICIAL
      Case controlaMes Of
        01 : mesExt := 'Janeiro';
        02 : mesExt := 'Fevereiro';
        03 : mesExt := 'Março';
        04 : mesExt := 'Abril';
        05 : mesExt := 'Maio';
        06 : mesExt := 'Junho';
        07 : mesExt := 'Julho';
        08 : mesExt := 'Agosto';
        09 : mesExt := 'Setembro';
        10 : mesExt := 'Outubro';
        11 : mesExt := 'Novembro';
        12 : mesExt := 'Dezembro';
      end;

      montaSQL := montaSQL + ', ISNULL(Mes' + x.ToString + ', '''') ' + mesExt.Substring(0, 3) + ' ';

      montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes' + x.ToString + ',''''))) ' + mesExt.Substring(0, 3) + ' ';

    end;

    montaSQL := montaSQL + ', CONVERT(INT, ISNULL(Mes0,0)) - CONVERT(INT, ISNULL(Mes1,0)) Divergência FROM TempDash_Semanas ';

    montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) Divergência FROM TempDash_Semanas ';

    case I of
      1: begin
        montaSQLI := ' WHERE id <= 7';
      end;
      2: begin
        montaSQLI := ' WHERE id BETWEEN 8 AND 14 '
      end;
      3: begin
        montaSQLI := ' WHERE id BETWEEN 15 AND 21 '
      end;
      4: begin
        montaSQLI := ' WHERE id BETWEEN 22 AND 28 '
      end;
      5: begin
        montaSQLI := ' WHERE id BETWEEN 29 AND 35 '
      end;
      6: begin
        montaSQLI := ' WHERE id >= 36 ';
      end;
    end;

    montaSQL := montaSQL + montaSQLI +
    ' UNION ' +
    ' SELECT ''Total'' Dias, CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes2,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes1,'''')))) , CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0, '''')) - CONVERT(INT, ISNULL(Mes1, '''')))) ' +
    ' FROM TempDash_Semanas ' + montaSQLI;

    montaSQLII := montaSQLII + montaSQLI + ' UNION ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(montaSQL);

    montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semana 0' + I.ToString +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral1.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral1.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral1.First;
    for x := 0 to qryGeral1.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral1.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral1.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral1.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral1.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';
  End;

  montaSQLII := montaSQLII +
  ' SELECT ''Total'', SUM(CONVERT(INT, ISNULL(Mes2,''''))), SUM(CONVERT(INT, ISNULL(Mes1,''''))), SUM(CONVERT(INT, ISNULL(Mes0,''''))), SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) ' +
  ' FROM TempDash_Semanas ';

  LimpaQuery(qryGeral2);
  qryGeral2.Open(montaSQLII);

  //PREENCHE O ULTIMO CARD COM TODAS AS SEMANAS
  montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semanas ' +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral2.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral2.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral2.First;
    for x := 0 to qryGeral2.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral2.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral2.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral2.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral2.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';

  montaCards := montaCards +
  ' </div> ';

   WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
    .ID('CardsSemanas')
    .HTML(montaCards)
    .Update;

   form_esmaecer_fundo.Hide;
End;

procedure TunitHome.AtualizaContatos(mes: String = ''; caCodigo: String = ''; campID: String = '');
var
  navBarFunction: String;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  mesExt, diaSemanaExt: String;
  diaSemana, controlaMes: Integer;
  I, controlaID: Integer;
  x: Integer;
  rgbColor: String;
  montaSQL, montaSQLI, montaSQLII, montaCards, strHTML: String;
  dataCompleta: TDate;

  arrayQry: array of TFDQuery;

  headerCharts_1: iModelHTML;
  config_row_1_Charts: iModelHTMLChartsConfig;
  y: Integer;
begin

 form_esmaecer_fundo.Show;

  mesInicial := mes.Substring(5,2);
  anoInicial := mes.Substring(0,4);

  dataCompleta := StrToDate('01/' + mesInicial + '/' + anoInicial);

  //RETORNA NOME DO MÊS INICIAL
  Case StrToInt(mesInicial) Of
    01 : mesExt := 'Janeiro';
    02 : mesExt := 'Fevereiro';
    03 : mesExt := 'Março';
    04 : mesExt := 'Abril';
    05 : mesExt := 'Maio';
    06 : mesExt := 'Junho';
    07 : mesExt := 'Julho';
    08 : mesExt := 'Agosto';
    09 : mesExt := 'Setembro';
    10 : mesExt := 'Outubro';
    11 : mesExt := 'Novembro';
    12 : mesExt := 'Dezembro';
  end;

  LimpaQuery(qryGeral);
  qryGeral.Open('SELECT CaNome FROM Campanhas WITH(NOLOCK) WHERE Campanhas_ID = ' + campID);

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
    .ID('InfoGeral')
    .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Comparativos por Leads Contatados</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
          ' Em ' + mesExt + '/' + anoInicial +  ' & ' + caCodigo + ' - ' + qryGeral.FieldByName('CaNome').AsString +
        ' </figcaption> ' +
      ' </figure> ')
    .Update;


  PreencheTableDash_Semanas;

  for x := 0 to 2 do
  Begin

    controlaID := 1;

    mesFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2);
    anoFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(6,4);

    for I := 1 to 31 do
    Begin

      LimpaQuery(qryGeral);
      qryGeral.Open(
        ' SELECT DISTINCT DATEPART(WEEKDAY, CoDataInicioFicha) DiaSemana, DATENAME(WEEKDAY, CoDataInicioFicha) DiaSemanaExt ' +
        ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        ' WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '   AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal);

      diaSemanaExt := qryGeral.FieldByName('DiaSemanaExt').AsString;
      diaSemana := qryGeral.FieldByName('DiaSemana').AsInteger;

      if controlaID = 1 then
        controlaID := diaSemana;

      if controlaID <= 7 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0) ' +
        ' WHERE id <= 7 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 8) and (controlaID <= 14) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0) ' +
        ' WHERE id BETWEEN 8 AND 14 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 15) and (controlaID <= 21) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0) ' +
        ' WHERE id BETWEEN 15 AND 21 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 22) and (controlaID <= 28) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0) ' +
        ' WHERE id BETWEEN 22 AND 28 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 29) and (controlaID <= 35) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0) ' +
        ' WHERE id BETWEEN 29 AND 35 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if controlaID >= 36 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0) ' +
        ' WHERE id >= 36 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End;


      controlaID := controlaID + 1;

    End;

  End;

  headerCharts_1 := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts_1
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINESemana')
                  .ColSpan(6)
                  .Heigth(130);


  SetLength(arrayQry, 3);
  for I := 2 downto 0 do
  begin

    controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - I)).Substring(3,2));

    //RETORNA NOME DO MÊS INICIAL
    Case controlaMes Of
      01 : mesExt := 'Janeiro';
      02 : mesExt := 'Fevereiro';
      03 : mesExt := 'Março';
      04 : mesExt := 'Abril';
      05 : mesExt := 'Maio';
      06 : mesExt := 'Junho';
      07 : mesExt := 'Julho';
      08 : mesExt := 'Agosto';
      09 : mesExt := 'Setembro';
      10 : mesExt := 'Outubro';
      11 : mesExt := 'Novembro';
      12 : mesExt := 'Dezembro';
    end;

    if I = 2 then
      rgbColor := '56, 56, 223'
    else if I = 1 then
      rgbColor := '223, 56, 56'
    else
      rgbColor := '56, 223, 56';


    montaSQL := montaSQL + ' SELECT ' + QuotedStr(mesExt) + ' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '
      + QuotedStr(rgbColor) + ' RGB FROM TempDash_Semanas ';

    if I <> 0 then
      montaSQL := montaSQL + ' UNION ';

    try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].Open(
        ' SELECT ''SEMANA 01'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id <= 7 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 02'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 8 AND 14 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 03'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 15 AND 21 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 04'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 22 AND 28 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 05'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 29 AND 35 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 06'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id >= 36 ');

        config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(mesExt)
                .BackgroundColor(rgbColor)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(rgbColor)
              .&End;

    finally

    end;

  end;

  config_row_1_Charts.&End
              .UpdateChart;

  LimpaQuery(qryGeral);
  qryGeral.Open(montaSQL);

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BARSemanas')
          .ColSpan(4)
          .Heigth(210)
          .DataSet
            .DataSet(qryGeral)
            .BackgroundOpacity(8)
          .&End
        .&End
      .UpdateChart;


  montaSQL := '';
  montaSQLII := '';
  strHTML := '';

  montaCards :=
    ' <div ' +
    '   class="row row-cols-1 row-cols-md-3 g-4" ' +
    '   style="width: 80%; margin-left: 10%; font-size: 13px" ' +
    ' > ';
  for I := 1 to 6 do
  Begin
    montaSQL := ' SELECT DiaSemana Dias ';

    montaSQLII := montaSQLII + 'SELECT ''SEMANA 0' + i.ToString + ''' Semanas ';

    for x := 2 downto 0 do
    begin
     controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2));

      //RETORNA NOME DO MÊS INICIAL
      Case controlaMes Of
        01 : mesExt := 'Janeiro';
        02 : mesExt := 'Fevereiro';
        03 : mesExt := 'Março';
        04 : mesExt := 'Abril';
        05 : mesExt := 'Maio';
        06 : mesExt := 'Junho';
        07 : mesExt := 'Julho';
        08 : mesExt := 'Agosto';
        09 : mesExt := 'Setembro';
        10 : mesExt := 'Outubro';
        11 : mesExt := 'Novembro';
        12 : mesExt := 'Dezembro';
      end;

      montaSQL := montaSQL + ', ISNULL(Mes' + x.ToString + ', '''') ''' + mesExt.Substring(0, 3) + ''' ';

      montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes' + x.ToString + ',''''))) ''' + mesExt.Substring(0, 3) + ''' ';

    end;

    montaSQL := montaSQL + ', CONVERT(INT, ISNULL(Mes0,0)) - CONVERT(INT, ISNULL(Mes1,0)) Divergência FROM TempDash_Semanas ';

    montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) Divergência FROM TempDash_Semanas ';

    case I of
      1: begin
        montaSQLI := ' WHERE id <= 7';
      end;
      2: begin
        montaSQLI := ' WHERE id BETWEEN 8 AND 14 '
      end;
      3: begin
        montaSQLI := ' WHERE id BETWEEN 15 AND 21 '
      end;
      4: begin
        montaSQLI := ' WHERE id BETWEEN 22 AND 28 '
      end;
      5: begin
        montaSQLI := ' WHERE id BETWEEN 29 AND 35 '
      end;
      6: begin
        montaSQLI := ' WHERE id >= 36 ';
      end;
    end;

    montaSQL := montaSQL + montaSQLI +
    ' UNION ' +
    ' SELECT ''Total'' Dias, CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes2,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes1,'''')))) , CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0, '''')) - CONVERT(INT, ISNULL(Mes1, '''')))) ' +
    ' FROM TempDash_Semanas ' + montaSQLI;

    montaSQLII := montaSQLII + montaSQLI + ' UNION ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(montaSQL);

    montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semana 0' + I.ToString +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral1.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral1.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral1.First;
    for x := 0 to qryGeral1.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral1.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral1.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral1.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral1.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';
  End;

  montaSQLII := montaSQLII +
  ' SELECT ''Total'', SUM(CONVERT(INT, ISNULL(Mes2,''''))), SUM(CONVERT(INT, ISNULL(Mes1,''''))), SUM(CONVERT(INT, ISNULL(Mes0,''''))), SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) ' +
  ' FROM TempDash_Semanas ';

  LimpaQuery(qryGeral2);
  qryGeral2.Open(montaSQLII);

  //PREENCHE O ULTIMO CARD COM TODAS AS SEMANAS
  montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semanas ' +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral2.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral2.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral2.First;
    for x := 0 to qryGeral2.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral2.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral2.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral2.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral2.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';

  montaCards := montaCards +
  ' </div> ';

   WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
    .ID('CardsSemanas')
    .HTML(montaCards)
    .Update;

   form_esmaecer_fundo.Hide;

end;

procedure TunitHome.AtualizaGrupos(mes: String = ''; caCodigo: String = ''; campID: String = ''; gruposID: String = '');
var
  navBarFunction: String;
  mesInicial, mesFinal, anoInicial, anoFinal: String;
  mesExt, diaSemanaExt: String;
  diaSemana, controlaMes: Integer;
  I, controlaID: Integer;
  x: Integer;
  rgbColor: String;
  montaSQL, montaSQLI, montaSQLII, montaCards, strHTML: String;
  dataCompleta: TDate;

  arrayQry: array of TFDQuery;

  headerCharts_1: iModelHTML;
  config_row_1_Charts: iModelHTMLChartsConfig;
  y: Integer;
begin

 form_esmaecer_fundo.Show;

  mesInicial := mes.Substring(5,2);
  anoInicial := mes.Substring(0,4);

  dataCompleta := StrToDate('01/' + mesInicial + '/' + anoInicial);

  //RETORNA NOME DO MÊS INICIAL
  Case StrToInt(mesInicial) Of
    01 : mesExt := 'Janeiro';
    02 : mesExt := 'Fevereiro';
    03 : mesExt := 'Março';
    04 : mesExt := 'Abril';
    05 : mesExt := 'Maio';
    06 : mesExt := 'Junho';
    07 : mesExt := 'Julho';
    08 : mesExt := 'Agosto';
    09 : mesExt := 'Setembro';
    10 : mesExt := 'Outubro';
    11 : mesExt := 'Novembro';
    12 : mesExt := 'Dezembro';
  end;

  LimpaQuery(qryGeral);
  qryGeral.Open('SELECT CaNome FROM Campanhas WITH(NOLOCK) WHERE Campanhas_ID = ' + campID);

  LimpaQuery(qryGeral1);
  qryGeral1.Open('SELECT ReGrupoNome FROM GruposResumoOperacoes WITH(NOLOCK) WHERE GruposResumoOperacoes_ID = ' + gruposID);

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
    .ID('InfoGeral')
    .HTML(
      ' <figure class="text-center"> ' +
        ' <blockquote class="blockquote"> ' +
          ' <p class="fs-2" style="font-family: Product Sans">Comparativos por Grupos de R.O.s</p> ' +
        ' </blockquote> ' +
        ' <figcaption class="blockquote-footer" style=" font-family: Product Sans;"> ' +
          ' Em ' + mesExt + '/' + anoInicial +  ' & ' + caCodigo + ' - ' + qryGeral.FieldByName('CaNome').AsString + ' com '
            + qryGeral1.FieldByName('ReGrupoNome').AsString +
        ' </figcaption> ' +
      ' </figure> ')
    .Update;

  PreencheTableDash_Semanas;

  for x := 0 to 2 do
  Begin

    controlaID := 1;

    mesFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2);
    anoFinal := DateToStr(IncMonth(dataCompleta, - x)).Substring(6,4);

    for I := 1 to 31 do
    Begin

      LimpaQuery(qryGeral);
      qryGeral.Open(
        ' SELECT DISTINCT DATEPART(WEEKDAY, CoDataInicioFicha) DiaSemana, DATENAME(WEEKDAY, CoDataInicioFicha) DiaSemanaExt ' +
        ' FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        ' WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '   AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal);

      diaSemanaExt := qryGeral.FieldByName('DiaSemanaExt').AsString;
      diaSemana := qryGeral.FieldByName('DiaSemana').AsInteger;

      if controlaID = 1 then
        controlaID := diaSemana;

      if controlaID <= 7 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
        '     AND CoResumoOperacaoID IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID = ' + gruposID + ')) ' +
        ' WHERE id <= 7 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 8) and (controlaID <= 14) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
        '     AND CoResumoOperacaoID IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID = ' + gruposID + ')) ' +
        ' WHERE id BETWEEN 8 AND 14 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 15) and (controlaID <= 21) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
        '     AND CoResumoOperacaoID IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID = ' + gruposID + ')) ' +
        ' WHERE id BETWEEN 15 AND 21 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 22) and (controlaID <= 28) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
        '     AND CoResumoOperacaoID IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID = ' + gruposID + ')) ' +
        ' WHERE id BETWEEN 22 AND 28 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if (controlaID >= 29) and (controlaID <= 35) then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
        '     AND CoResumoOperacaoID IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID = ' + gruposID + ')) ' +
        ' WHERE id BETWEEN 29 AND 35 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End
      else if controlaID >= 36 then
      Begin

        LimpaQuery(qryGeral1);
        qryGeral1.ExecSQL(
        ' UPDATE TempDash_Semanas ' +
        ' SET Mes' + x.ToString + ' = (SELECT DISTINCT COUNT(*) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ' AND ISNULL(CoCampanhaOrigemID,0) = 0 ' +
        '     AND CoResumoOperacaoID IN (SELECT ReResumoOperacoesID FROM ResumoGrupos WITH(NOLOCK) WHERE ReGrupoResumoOperacoesID = ' + gruposID + ')) ' +
        ' WHERE id >= 36 AND DiaSemana = (SELECT DISTINCT DATENAME(WEEKDAY, CoDataInicioFicha) ' +
        ' 	FROM ContatosFichas_' + caCodigo + ' WITH(NOLOCK) ' +
        '   WHERE DATEPART(DAY, CoDataInicioFicha) = ' + I.ToString + ' AND DATEPART(MONTH, CoDataInicioFicha) = ' + mesFinal +
        '     AND DATEPART(YEAR, CoDataInicioFicha) = ' + anoFinal + ') ')

      End;


      controlaID := controlaID + 1;

    End;

  End;

  headerCharts_1 := WebCharts1
          .ContinuosProject
          .WindowParent(FMXWindowParent)
          .WebBrowser(FMXChromium1);

  config_row_1_Charts := headerCharts_1
            .Charts
              ._ChartType(line)
                .Attributes
                  .Name('LINESemana')
                  .ColSpan(6)
                  .Heigth(130);


  SetLength(arrayQry, 3);
  for I := 2 downto 0 do
  begin

    controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - I)).Substring(3,2));

    //RETORNA NOME DO MÊS INICIAL
    Case controlaMes Of
      01 : mesExt := 'Janeiro';
      02 : mesExt := 'Fevereiro';
      03 : mesExt := 'Março';
      04 : mesExt := 'Abril';
      05 : mesExt := 'Maio';
      06 : mesExt := 'Junho';
      07 : mesExt := 'Julho';
      08 : mesExt := 'Agosto';
      09 : mesExt := 'Setembro';
      10 : mesExt := 'Outubro';
      11 : mesExt := 'Novembro';
      12 : mesExt := 'Dezembro';
    end;

    if I = 2 then
      rgbColor := '56, 56, 223'
    else if I = 1 then
      rgbColor := '223, 56, 56'
    else
      rgbColor := '56, 223, 56';


    montaSQL := montaSQL + ' SELECT ' + QuotedStr(mesExt) + ' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '
      + QuotedStr(rgbColor) + ' RGB FROM TempDash_Semanas ';

    if I <> 0 then
      montaSQL := montaSQL + ' UNION ';

    try

       arrayQry[I] := TFDQuery.Create(self);
       arrayQry[I].Connection := dmDados.Conn;

       arrayQry[I].Active := false;

       arrayQry[I].SQL.Clear;
       arrayQry[I].Open(
        ' SELECT ''SEMANA 01'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id <= 7 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 02'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 8 AND 14 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 03'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 15 AND 21 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 04'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 22 AND 28 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 05'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id BETWEEN 29 AND 35 ' +
        ' UNION ' +
        ' SELECT ''SEMANA 06'' LABEL, SUM(CONVERT(INT, ISNULL(Mes' + I.ToString + ',''''))) VALUE, '''' RGB ' +
        ' FROM TempDash_Semanas ' +
        ' WHERE id >= 36 ');

        config_row_1_Charts
              .DataSet
                .DataSet(arrayQry[I])
                .textLabel(mesExt)
                .BackgroundColor(rgbColor)
                .BorderWidth(3)
                .BackgroundOpacity(7)
                .BorderColor(rgbColor)
              .&End;

    finally

    end;

  end;

  config_row_1_Charts.&End
              .UpdateChart;

  LimpaQuery(qryGeral);
  qryGeral.Open(montaSQL);

  WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .Charts
      ._ChartType(bar)
        .Attributes
          .Name('BARSemanas')
          .ColSpan(4)
          .Heigth(210)
          .DataSet
            .DataSet(qryGeral)
            .BackgroundOpacity(8)
          .&End
        .&End
      .UpdateChart;


  montaSQL := '';
  montaSQLII := '';
  strHTML := '';

  montaCards :=
    ' <div ' +
    '   class="row row-cols-1 row-cols-md-3 g-4" ' +
    '   style="width: 80%; margin-left: 10%; font-size: 13px" ' +
    ' > ';
  for I := 1 to 6 do
  Begin
    montaSQL := ' SELECT DiaSemana Dias ';

    montaSQLII := montaSQLII + 'SELECT ''SEMANA 0' + i.ToString + ''' Semanas ';

    for x := 2 downto 0 do
    begin
     controlaMes := StrToInt(DateToStr(IncMonth(dataCompleta, - x)).Substring(3,2));

      //RETORNA NOME DO MÊS INICIAL
      Case controlaMes Of
        01 : mesExt := 'Janeiro';
        02 : mesExt := 'Fevereiro';
        03 : mesExt := 'Março';
        04 : mesExt := 'Abril';
        05 : mesExt := 'Maio';
        06 : mesExt := 'Junho';
        07 : mesExt := 'Julho';
        08 : mesExt := 'Agosto';
        09 : mesExt := 'Setembro';
        10 : mesExt := 'Outubro';
        11 : mesExt := 'Novembro';
        12 : mesExt := 'Dezembro';
      end;

      montaSQL := montaSQL + ', ISNULL(Mes' + x.ToString + ', '''') ''' + mesExt.Substring(0, 3) + ''' ';

      montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes' + x.ToString + ',''''))) ''' + mesExt.Substring(0, 3) + ''' ';

    end;

    montaSQL := montaSQL + ', CONVERT(INT, ISNULL(Mes0,0)) - CONVERT(INT, ISNULL(Mes1,0)) Divergência FROM TempDash_Semanas ';

    montaSQLII := montaSQLII + ', SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) Divergência FROM TempDash_Semanas ';

    case I of
      1: begin
        montaSQLI := ' WHERE id <= 7';
      end;
      2: begin
        montaSQLI := ' WHERE id BETWEEN 8 AND 14 '
      end;
      3: begin
        montaSQLI := ' WHERE id BETWEEN 15 AND 21 '
      end;
      4: begin
        montaSQLI := ' WHERE id BETWEEN 22 AND 28 '
      end;
      5: begin
        montaSQLI := ' WHERE id BETWEEN 29 AND 35 '
      end;
      6: begin
        montaSQLI := ' WHERE id >= 36 ';
      end;
    end;

    montaSQL := montaSQL + montaSQLI +
    ' UNION ' +
    ' SELECT ''Total'' Dias, CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes2,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes1,'''')))) , CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0,'''')))), ' +
    ' CONVERT(VARCHAR(10), SUM(CONVERT(INT, ISNULL(Mes0, '''')) - CONVERT(INT, ISNULL(Mes1, '''')))) ' +
    ' FROM TempDash_Semanas ' + montaSQLI;

    montaSQLII := montaSQLII + montaSQLI + ' UNION ';

    LimpaQuery(qryGeral1);
    qryGeral1.Open(montaSQL);

    montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semana 0' + I.ToString +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral1.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral1.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral1.First;
    for x := 0 to qryGeral1.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral1.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral1.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral1.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral1.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';
  End;

  montaSQLII := montaSQLII +
  ' SELECT ''Total'', SUM(CONVERT(INT, ISNULL(Mes2,''''))), SUM(CONVERT(INT, ISNULL(Mes1,''''))), SUM(CONVERT(INT, ISNULL(Mes0,''''))), SUM(CONVERT(INT, ISNULL(Mes0,'''')) - CONVERT(INT, ISNULL(Mes1,''''))) ' +
  ' FROM TempDash_Semanas ';

  LimpaQuery(qryGeral2);
  qryGeral2.Open(montaSQLII);

  //PREENCHE O ULTIMO CARD COM TODAS AS SEMANAS
  montaCards := montaCards +
    ' <div class="col"> ' +
    '   <div ' +
    '     class="card shadow p-3 mb-5 bg-body rounded" ' +
    '     style="background-color: #ffffff; margin-bottom: 10px" ' +
    '    > ' +
    '     <div class="card-header" style="text-align: center"> ' +
    '       Semanas ' +
    '     </div> ' +
    '     <table ' +
    '       class="table table-responsive-sm table-borderless table-striped" ' +
    '       style="margin-top: 5px" ' +
    '     > ';

    montaCards := montaCards + ' <tr> ';

    for X := 0 to qryGeral2.FieldCount - 1 do
      montaCards := montaCards + ' <th> ' + qryGeral2.Fields[X].FieldName + ' </th> ';

    montaCards := montaCards + ' </tr> ';

    qryGeral2.First;
    for x := 0 to qryGeral2.RecordCount - 1 do
    Begin
      montaCards := montaCards + ' <tr> ';
      for y := 0 to qryGeral2.FieldCount - 1 do
      Begin
        if y = 0 then
          montaCards := montaCards + ' <th> ' + qryGeral2.Fields[Y].AsString + ' </th> '
        else
          montaCards := montaCards + ' <td > ' + qryGeral2.Fields[Y].AsString + ' </td> ';
      End;
      montaCards := montaCards + ' </tr> ';
      qryGeral2.Next;
    End;

    montaCards := montaCards +
    '  </table> ' +
    ' </div> ' +
    '</div> ';

  montaCards := montaCards +
  ' </div> ';

   WebCharts1
    .ContinuosProject
    .WindowParent(FMXWindowParent)
    .WebBrowser(FMXChromium1)
    .DOMElement
    .ID('CardsSemanas')
    .HTML(montaCards)
    .Update;

   form_esmaecer_fundo.Hide;

end;

procedure TunitHome.Timer1Timer(Sender: TObject);
begin

  Timer1.Enabled := false;
  GraphOne;
//  GraphSeven;
end;

procedure TunitHome.NotifyMoveOrResizeStarted;
begin
  if (FMXChromium1 <> nil) then FMXChromium1.NotifyMoveOrResizeStarted;
end;

function TunitHome.PostCustomMessage(aMsg: cardinal; aWParam: WPARAM;
  aLParam: LPARAM): boolean;
{$IFDEF MSWINDOWS}
var
  TempHWND : HWND;
{$ENDIF}
begin
{$IFDEF MSWINDOWS}
  TempHWND := FmxHandleToHWND(Handle);
  Result   := (TempHWND <> 0) and WinApi.Windows.PostMessage(TempHWND, aMsg, aWParam, aLParam);
{$ELSE}
  Result   := False;
{$ENDIF}
end;

procedure TunitHome.ResizeChild;
begin
  if (FMXWindowParent <> nil) then
    FMXWindowParent.SetBounds(GetFMXWindowParentRect);
end;

procedure TunitHome.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
var
  PositionChanged: Boolean;
begin
  PositionChanged := (ALeft <> Left) or (ATop <> Top);
  inherited SetBounds(ALeft, ATop, AWidth, AHeight);
  if PositionChanged then
    NotifyMoveOrResizeStarted;
end;

end.
