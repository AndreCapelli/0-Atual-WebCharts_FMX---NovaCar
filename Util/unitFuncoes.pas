unit unitFuncoes;

interface

uses FireDAC.Comp.Client, System.SysUtils, Vcl.Dialogs, Vcl.StdCtrls,
  Winapi.Messages, System.UITypes, Vcl.CheckLst, FMX.ListBox;

var
  ContaID, URLCallback, ClientID, BasicToken: string;

   var BD_NovoCaminho: string;

function Retorna_IdentCurrent(Tabela: string): String;
procedure LimpaQuery(Query: TFDQuery);
function IifStr(Teste: Boolean; ValorTrue, ValorFalse: String): String;
  overload;
Function GeraCPJCNPJ(TipoPessoa: string): String;
Procedure MsgErro(Msg: String);
Procedure MsgInformacao(Msg: String);
procedure ScrollMemo(Memo: TMemo; Direction: Integer);
function RetiraNumerosString(Texto: String): string;
Function MsgYesNo(Msg: String): Boolean;
function TrocaCaracterEspecial_EmailUTF8(aTexto: string;
  aLimExt: Boolean): string;
procedure CarregaCheckListBox(Lst: TCheckListBox; Select: string);
procedure CarregaCombo(Select: string; Combo: TComboBox;
  IncluirNenhum: Boolean = false; ItemIndex: Integer = -1;
  IncluirTodos: Boolean = false; IncluirTodas: Boolean = false;
  Personalizado: string = ''; LimpaCombo: Boolean = false;
  IdAtual: Integer = 0);
Function RetornaIDCombo(Combo: TComboBox): String;
function NavBarHTML: String;
function MontaNavBar(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false; usaGruposRO: Boolean = false): String;
function MontaNavBarHTML(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
function MontaNavBar_Marketing(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
function MontaNavBar_Ranking(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
function MontaNavBarRO(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
function MontaDropGrupos(recebeID: String = ''; idHTML: String = ''; visible: String = ''): String;
function MontaDropCampanhas(recebeID: String = ''; idHTML: String = ''; visible: String = ''): String;
function MontaDropEquipes(recebeID: String = ''; idHTML: String = ''; visible: String = ''): String;
function MontaDropOperadores(recebeID: String = ''): String;
function MontaDropROs(recebeID: String = ''): String;
function MontaDropROs2(recebeID: String = ''): String;
function MontaDropROs3(recebeID: String = ''): String;
function MontaDropROs4(recebeID: String = ''): String;

procedure PreencheTableDash_Semanas;

implementation

uses unitDmdados;

Function MsgYesNo(Msg: String): Boolean;
Begin

  if MessageDlg(Msg, mtConfirmation, [mbYes, mbNo], 0, mbNo) <> mrYes then
    Result := false
  else
    Result := true;
End;

function RetiraNumerosString(Texto: String): string;
// Retorna só os numeros
var
  nI: Integer;
  TextoLimpo: String;
begin
  TextoLimpo := '';
  For nI := 1 to Length(Texto) do
  begin
    if Texto[nI] in ['0' .. '9'] then
      TextoLimpo := TextoLimpo + Texto[nI];
  end;
  Result := TextoLimpo;
end;

procedure ScrollMemo(Memo: TMemo; Direction: Integer);
{ Mover cursor do Memo para inicio ou final
  ScrollMemo(Memo1, SB_LINEUP); // Rola para o início
  ScrollMemo(Memo1, SB_LINEDOWN); // Rola para o final }
var
  ScrollMessage: TWMVScroll;
  i: Integer;
begin
  ScrollMessage.Msg := WM_VSCROLL;
  Memo.Lines.BeginUpdate;
  try
    for i := 0 to Memo.Lines.Count do
    begin
      ScrollMessage.ScrollCode := Direction;
      ScrollMessage.pos := 0;
      Memo.Dispatch(ScrollMessage);
    end;
  finally
    Memo.Lines.EndUpdate;
  end;
end;

function Retorna_IdentCurrent(Tabela: string): String;
var
  ID: string;
begin

  LimpaQuery(dmdados.qryIdentCurrent);
  dmdados.qryIdentCurrent.Open('SELECT IDENT_CURRENT(' + quotedstr(Tabela)
    + ') ID');

  ID := dmdados.qryIdentCurrent.FieldByName('ID').AsString;
  LimpaQuery(dmdados.qryIdentCurrent);
  Result := ID;

end;

procedure LimpaQuery(Query: TFDQuery);
begin
  Query.Close;
  Query.SQL.clear;
end;

function IifStr(Teste: Boolean; ValorTrue, ValorFalse: String): String;
  overload;
begin
  If Teste then
    Result := ValorTrue
  else
    Result := ValorFalse;
end;

Function GeraCPJCNPJ(TipoPessoa: string): String;
var
  Texto, Resultado: string;
  i, Tamanho: Integer;
begin
  LimpaQuery(dmdados.qryGeral);
  dmdados.qryGeral.Open('select IDENT_CURRENT (' + quotedstr('PESSOAS') +
    ') + 1 as ID');
  Texto := dmdados.qryGeral.FieldByName('ID').AsString;
  Tamanho := Length(Texto);
  if TipoPessoa = 'F' then
  begin
    for i := 1 to 11 - Tamanho do
    begin
      Resultado := '0' + Resultado;
    end;
  end
  else
  begin
    for i := 1 to 14 - Tamanho do
    begin
      Resultado := '0' + Resultado;
    end;
  end;
  Result := Resultado + Texto;

end;

Procedure MsgInformacao(Msg: String);
Begin
  MessageDlg(Msg, mtInformation, [mbOK], 0);
End;

Procedure MsgErro(Msg: String);
Begin
  MessageDlg(Msg, mtError, [mbOK], 0);
End;

function TrocaCaracterEspecial_EmailUTF8(aTexto: string;
  aLimExt: Boolean): string;
const
  // Lista de caracteres especiais
  xCarEsp: array [1 .. 41] of String = ('á', 'à', 'ã', 'â', 'ä', 'Á', 'À', 'Ã',
    'Â', 'Ä', 'é', 'è', 'É', 'È', 'í', 'ì', 'Í', 'Ì', 'ó', 'ò', 'ö', 'õ', 'ô',
    'Ó', 'Ò', 'Ö', 'Õ', 'Ô', 'ú', 'ù', 'ü', 'Ú', 'Ù', 'Ü', 'ç', 'Ç', 'ñ', 'Ñ',
    '“', '<', '>');
  // Lista de caracteres para troca
  xCarTro: array [1 .. 41] of String = ('&aacute;', '&agrave;', '&atilde;',
    '&acirc;', 'a', '&Aacute;', '&Agrave;', '&Atilde;', '&Acirc;', 'A',
    '&eacute;', 'e', '&Eacute;', 'E', '&iacute;', 'i', '&Iacute;', 'I',
    '&oacute;', 'o', 'o', '&otilde;', '&ocirc;', '&Oacute;', 'O', 'O',
    '&Otilde;', '&Ocirc;', '&uacute;', 'u', 'u', '&Uacute;', 'u', 'u',
    '&ccedil;', '&Ccedil;', '&ntilde;', '&Ntilde;', '&quot;', '&lt;', '&gt;');
  // Lista de Caracteres Extras
  xCarExt: array [1 .. 48] of string = ('<', '>', '!', '@', '#', '$', '%', '¨',
    '&', '*', '(', ')', '_', '+', '=', '{', '}', '[', ']', '?', ';', ':', ',',
    '|', '*', '"', '~', '^', '´', '`', '¨', 'æ', 'Æ', 'ø', '£', 'Ø', 'ƒ', 'ª',
    'º', '¿', '®', '½', '¼', 'ß', 'µ', 'þ', 'ý', 'Ý');
var
  xTexto: string;
  i: Integer;
begin
  xTexto := aTexto;
  for i := 1 to 41 do
    xTexto := StringReplace(xTexto, xCarEsp[i], xCarTro[i], [rfReplaceAll]);

  // De acordo com o parâmetro aLimExt, elimina caracteres extras.
  if (aLimExt) then
    for i := 1 to 42 do
      xTexto := StringReplace(xTexto, xCarExt[i], '', [rfReplaceAll]);

  Result := xTexto;
end;

procedure CarregaCheckListBox(Lst: TCheckListBox; Select: string);
var
  i: Integer;
begin
  LimpaQuery(dmdados.qryGeral);
  dmdados.qryGeral.Open(Select);

  Lst.clear;

  for i := 1 to dmdados.qryGeral.RecordCount do
  begin
    Lst.Items.AddObject(dmdados.qryGeral.FieldByName('Nome').AsString,
      TObject(dmdados.qryGeral.FieldByName('ID').AsInteger));
    dmdados.qryGeral.Next;
  end;

  LimpaQuery(dmdados.qryGeral);
end;

Function RetornaIDCombo(Combo: TComboBox): String;
begin
  if Combo.ItemIndex = -1 then
    Result := '0'
  else
    Result := inttostr(Integer(Combo.Items.Objects[Combo.ItemIndex]));
end;

procedure CarregaCombo(Select: string; Combo: TComboBox;
  IncluirNenhum: Boolean = false; ItemIndex: Integer = -1;
  IncluirTodos: Boolean = false; IncluirTodas: Boolean = false;
  Personalizado: string = ''; LimpaCombo: Boolean = false;
  IdAtual: Integer = 0);
var
  i: Integer;
begin
  LimpaQuery(dmdados.qryGeral3);

  dmdados.qryGeral3.Open(Select);

  if LimpaCombo then
    Combo.clear;

  if IncluirNenhum then
    Combo.Items.AddObject('Nenhuma', TObject(0));

  if IncluirTodos then
    Combo.Items.AddObject('Todos', TObject(0));

  if IncluirTodas then
    Combo.Items.AddObject('Todas', TObject(0));

  if Personalizado <> '' then
    Combo.Items.AddObject(Personalizado, TObject(0));

  for i := 1 to dmdados.qryGeral3.RecordCount do
  begin
    Combo.Items.AddObject(dmdados.qryGeral3.FieldByName('Nome').AsString,
      TObject(dmdados.qryGeral3.FieldByName('ID').AsInteger));
    dmdados.qryGeral3.Next;
  end;

  Combo.ItemIndex := ItemIndex;

  if IdAtual > 0 then
    for i := 0 to Combo.Items.Count - 1 do
    begin
      Combo.ItemIndex := i;
      if RetornaIDCombo(Combo) = inttostr(IdAtual) then
        break
      else
        Combo.ItemIndex := ItemIndex;
    end;

  LimpaQuery(dmdados.qryGeral3);

end;

function NavBarHTML: String;
Begin

  Result :=
  ' <nav class="navbar navbar-expand-lg navbar-light bg-light"> ' +
    ' <div class="container-fluid"> ' +
      ' <div class="collapse navbar-collapse" id="navbarScroll"> ' +
        ' <ul class="navbar-nav me-auto my-2 my-lg-0 navbar-nav-scroll" style="--bs-scroll-height: 100px;"> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link active" aria-current="page" href="ActionCallBackJS:GraphOne('''','''','''','''','''','''','''')">Resumo de Operações</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link disabled" aria-current="page">|</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link active" aria-current="page" href="ActionCallBackJS:GraphTwo('''','''','''','''','''','''','''')">Comparativo por Contatos</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link disabled" aria-current="page">|</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link active" aria-current="page" href="ActionCallBackJS:GraphThree('''','''','''','''','''','''','''')">Produção</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link disabled" aria-current="page">|</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link active" aria-current="page" href="ActionCallBackJS:GraphFour('''','''','''','''','''','''','''')">Conta Corrente</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link disabled" aria-current="page">|</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link active" aria-current="page" href="ActionCallBackJS:GraphFive('''','''','''','''','''','''','''')">Vendas por Origem</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link disabled" aria-current="page">|</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link active" aria-current="page" href="ActionCallBackJS:GraphSix('''','''','''','''','''','''','''')">Comparativo por Semanas</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link disabled" aria-current="page">|</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link active" aria-current="page" href="ActionCallBackJS:GraphSeven('''','''','''','''','''','''','''')">Ranking</a> ' +
          ' </li> ' +
          ' <li class="nav-item"> ' +
            ' <a class="nav-link disabled" aria-current="page">|</a> ' +
          ' </li> ' +
        ' </ul> ' +
       ' <button class="btn btn-outline-secondary" type="button" data-bs-toggle="offcanvas" data-bs-target="#offcanvasScrolling" aria-controls="offcanvasScrolling">Filtros <span class="navbar-toggler-icon"></span></button> ' +
      ' </div> ' +
    ' </div> ' +
  ' </nav> ';

End;

function MontaNavBar(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false; usaGruposRO: Boolean = false): String;
var
  navBar: String;
  montaNav: String;
  dropCampContatos, dropCampGrupos, dropGrupos: String;
Begin

  navBar := NavBarHTML;

  if usaCampanha then
    dropCampContatos := MontaDropCampanhas(campanhaID, 'cbCampanhasContatos', 'hidden');

  if usaGruposRO then
  Begin
    dropCampGrupos := MontaDropCampanhas(campanhaID, 'cbCampanhasGrupos', 'hidden');
    dropGrupos := MontaDropGrupos('', 'cbGrupos', 'hidden');
  End;

  montaNav := navBar +
     //MONTA O MENU LATERAL A ESQUERDA
  ' <div class="offcanvas offcanvas-start" data-bs-scroll="true" data-bs-backdrop="false" tabindex="-1" id="offcanvasScrolling" aria-labelledby="offcanvasScrollingLabel"> ' +
    ' <div class="offcanvas-header"> ' +
     ' <h5 class="offcanvas-title" id="offcanvasScrollingLabel">Filtros</h5> ' +
     ' <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button> ' +
    ' </div> ' +
    ' <div class="offcanvas-body"> ' +

      ' <style> ';

      montaNav := montaNav +
      '   body { ' +
      '     font-family: Roboto; ' +
      '   } ' +
      ' </style> ';

      montaNav := montaNav +
      ' <form class="row g-3"> ' +
      '   <div class="col-auto"> ' +
      '     <input ' +
      '       type="month" ' +
      '       id="beginMonth" ' +
      '       value="" ' +
      '       style="font-size: 13px; width: 200px" ' +
      '     /> ' +
      '   </div> ' +
      '   <div class="form-check"> ' +
      '     <input ' +
      '       class="form-check-input" ' +
      '       type="radio" ' +
      '       name="radioBtn" ' +
      '       id="chkCadastro" ' +
      '       checked="checked" ' +
      '       value="Cadastros" ' +
      '       onclick="chamaCadastros()" ' +
      '     /> ' +
      '     <label class="form-check-label" for="chkCadastro"> ' +
      '       Comparativo por Cadastro ' +
      '     </label> ' +
      '   </div> ' +
      '   <div class="form-check"> ' +
      '     <input ' +
      '       class="form-check-input" ' +
      '       type="radio" ' +
      '       name="radioBtn" ' +
      '       id="chkContato" ' +
      '       value="Contatos" ' +
      '       onclick="chamaContatos()" ' +
      '     /> ' +
      '     <label class="form-check-label" for="chkContato"> ' +
      '       Comparativo por Contato ' +
      '     </label> ' +
      '   </div> ';

      //DROP LISTA DE CAMPANHAS
      if usaCampanha then
        montaNav := montaNav + dropCampContatos;

      montaNav := montaNav +
      ' <div class="form-check"> ' +
      '   <input ' +
      '     class="form-check-input" ' +
      '     type="radio" ' +
      '     name="radioBtn" ' +
      '     id="chkGrupos" ' +
      '     value="Grupos" ' +
      '     onclick="chamaGrupos()" ' +
      '   /> ' +
      '   <label class="form-check-label" for="chkGrupos"> ' +
      '     Comparativo por Grupos de RO ' +
      '   </label> ' +
      ' </div> ';

      if usaGruposRO then
        montaNav := montaNav + dropCampGrupos + dropGrupos;

      montaNav := montaNav +
      //BUTTON DE CONFIRMAR FILTRO
      '   <div class="col-12"> ' +
      '     <button id="btnFilter" type="button" class="btn btn-primary">Filtrar</button>' +
      '   </div> ';

      montaNav := montaNav +
      ' </form> ' +
      ' <script> ';

      montaNav := montaNav +
      ' $("#btnFilter").click(function () { ' +
      '   testClicked(); ' +
      ' }); ' +
      '  ' +
      ' const btnFilter = document.getElementById("btnFilter"); ' +
      ' function testClicked() { ' +
      '   let selectedSize; ' +
      '   for (const radioButton of radioButtons) { ' +
      '     if (radioButton.checked) { ' +
      '       selectedSize = radioButton.value; ' +
      '       break; ' +
      '     } ' +
      '   } ' +
      '   if (selectedSize == "Cadastros") { ' +
      '       var url = "ActionCallBackJS:AtualizaCadastros("+ window.btoa($(''#beginMonth'').val()) +")"; ' +
      '       location.assign(url); ' +
      '   } else if (selectedSize == "Contatos") { ' +
      '       var url = "ActionCallBackJS:AtualizaContatos("+ window.btoa(document.getElementById("beginMonth").value)+"," + window.btoa(campCodContatos)+ "," + window.btoa(campIDContatos)+ ")"; ' +
      '       location.assign(url); ' +
      '   } else if (selectedSize == "Grupos") { ' +
      '       var url = "ActionCallBackJS:AtualizaGrupos("+ window.btoa(document.getElementById("beginMonth").value)+"," + window.btoa(campCodGrupos)+ "," + window.btoa(campIDGrupos)+ ","+ window.btoa(gruposID)+")"; ' +
      '       location.assign(url); ' +
      '   } ' +
      ' } ';

      montaNav := montaNav +
      ' const radioButtons = document.querySelectorAll( ' +
      '   ''input[name="radioBtn"]'' ' +
      ' ); ' +
      '  ' +
      ' var campCodContatos = ""; ' +
      ' var campIDContatos = ""; ' +
      ' var campContatos; ' +
      ' const cbCampanhasContatos = document.getElementById( ' +
      '   "cbCampanhasContatos" ' +
      ' ); ' +
      ' cbCampanhasContatos.addEventListener("change", campCont); ' +
      ' function campCont() { ' +
      '   campContatos = cbCampanhasContatos.value; ' +
      '   campCodContatos = $(this) ' +
      '     .find(":selected") ' +
      '     .data("value").codigo; ' +
      '   campIDContatos = $(this).find(":selected").data("value").campID; ' +
      ' } ' +
      '  ' +
      ' var campCodGrupos = ""; ' +
      ' var campIDGrupos = ""; ' +
      ' var campGrupos; ' +
      ' const cbCampanhasGrupos = ' +
      ' document.getElementById("cbCampanhasGrupos"); ' +
      ' cbCampanhasGrupos.addEventListener("change", campGr); ' +
      ' function campGr() { ' +
      '   campGrupos = cbCampanhasGrupos.value; ' +
      '   campCodGrupos = $(this).find(":selected").data("value").codigo; ' +
      '   campIDGrupos = $(this).find(":selected").data("value").campID; ' +
      ' } ' +
      '  ' +
      ' var gruposID = ""; ' +
      ' var grupos; ' +
      ' const cbGrupos = document.getElementById("cbGrupos"); ' +
      ' cbGrupos.addEventListener("change", gr); ' +
      ' function gr() { ' +
      '   grupos = cbGrupos.value; ' +
      '   gruposID = $(this).find(":selected").data("value").grupoID; ' +
      ' } ' +
      '  ' +
      ' function chamaCadastros() { ' +
      '   cbCampanhasContatos.style.visibility = "hidden"; ' +
      '   cbCampanhasGrupos.style.visibility = "hidden"; ' +
      '   cbGrupos.style.visibility = "hidden"; ' +
      ' } ' +
      '  ' +
      ' function chamaContatos() { ' +
      '   cbCampanhasContatos.style.visibility = "visible"; ' +
      '   cbCampanhasGrupos.style.visibility = "hidden"; ' +
      '   cbGrupos.style.visibility = "hidden"; ' +
      ' } ' +
      '  ' +
      ' function chamaGrupos() { ' +
      '   cbCampanhasContatos.style.visibility = "hidden"; ' +
      '   cbCampanhasGrupos.style.visibility = "visible"; ' +
      '   cbGrupos.style.visibility = "visible"; ' +
      ' } ';

      montaNav := montaNav +
      ' </script> ' +

    ' </div> ' +
  ' </div> ';


  Result := montaNav;

End;

function MontaNavBarHTML(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
var
  navBar: String;
  montaNav: String;
  montaDrop: String;
  montaRO1, montaRO2, montaRO3, montaRO4: String;
Begin

  if usaCampanha then
    montaDrop := MontaDropCampanhas(campanhaID, 'cbCampanhas', 'visible');

  navBar := NavBarHTML;

  montaNav := navBar +

    //MONTA O MENU LATERAL A ESQUERDA
  ' <div class="offcanvas offcanvas-start" data-bs-scroll="true" data-bs-backdrop="false" tabindex="-1" id="offcanvasScrolling" aria-labelledby="offcanvasScrollingLabel"> ' +
    ' <div class="offcanvas-header"> ' +
     ' <h5 class="offcanvas-title" id="offcanvasScrollingLabel">Filtros</h5> ' +
     ' <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button> ' +
    ' </div> ' +
    ' <div class="offcanvas-body"> ' +

      ' <style> ' +
      //DEIXANDO DATA VERMELHA QUANDO É MENOR
      '   #finalDate:invalid { color: red; } ';

      if usaMes then
      begin
        //DEIXANDO MÊS VERMELHO QUANDO É MENOR
        montaNav := montaNav +
        '   #finalMonth:invalid { color: red; } ';
      end;

      montaNav := montaNav +
      '   body { ' +
      '     font-family: Roboto; ' +
      '   } ' +
      ' </style> ' +
      ' <form class="row g-3"> ';

      if usaMes then
      Begin

        //RADIO BUTTON PARA SELECIONAR DATA COMPLETA
        montaNav := montaNav +
        '   <div class="form-check"> ' +
        '     <input class="form-check-input" type="radio" name="radioDates" id="chkDate" value="chkDate1" checked="checked"> ' +
        '     <label class="form-check-label" for="chkDate"> ' +
        '       Habilitar Busca por Data: ' +
        '     </label> ' +
        '   </div> ';

      End;

      //DATA COMPLETAS JÁ COM A VALIDAÇÃO
      montaNav := montaNav +
      '   <div class="col-auto"> ' +
      '     <input type="date" name="beginDate" id="beginDate" value="' + dataI + '" style="font-size: 12px; width: 130px"> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <label for="staticText" class="visually"> até </label> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <input type="date" name="finalDate" id="finalDate" min="" value="' + dataF + '" style="font-size: 12px; width: 130px"> ' +
      '   </div> ';

      if usaMes then
      Begin

        //RADIO BUTTON PARA SELECIONAR MES
        montaNav := montaNav +
        '   <div class="form-check"> ' +
        '     <input class="form-check-input" type="radio" name="radioDates" id="chkMonth" value="chkMonth1"> ' +
        '     <label class="form-check-label" for="chkMonth"> ' +
        '       Habilitar Busca por Mês: ' +
        '     </label> ' +
        '   </div> ' +
        //MES
        '   <div class="col-auto"> ' +
        '     <input type="month" id="beginMonth" value="' + mesI + '" style="font-size: 12px; width: 140px" disabled="disabled"> ' +
        '   </div> ' +
        '   <div class="col-auto"> ' +
        '     <label for="staticText" class="visually"> até </label> ' +
        '   </div> ' +
        '   <div class="col-auto"> ' +
        '     <input type="month" id="finalMonth" value="' + mesF + '" style="font-size: 12px; width: 140px" disabled="disabled"> ' +
        '   </div> ';

      End;

      //DROP LISTA DE CAMPANHAS
      if usaCampanha then
        montaNav := montaNav + montaDrop;

      montaNav := montaNav +
      //BUTTON DE CONFIRMAR FILTRO
      '   <div class="col-12"> ' +
      '     <button id="btnFilter" type="button" class="btn btn-primary">Filtrar</button>' +
      '   </div> ' +

      ' </form> ' +

      ' <script> ' +

        ' $(''#btnFilter'').click(function(){ ' +
            ' testClicked(); ' +
        ' }); ';

        if usaMes then
        Begin


          montaNav := montaNav +
          ' function testClicked() { ' +

          '   if(document.getElementById("chkDate").checked) { ' +
          '       var url = "ActionCallBackJS:' + procNome + '(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ","+ window.btoa($(''#beginDate'').val())+"," + window.btoa($(''#finalDate'').val())+",'''','''','''')"; ' +
          '         location.assign(url); ' +
          '   } ' +
          '   else if(document.getElementById("chkMonth").checked) { ' +
          '       var url = "ActionCallBackJS:' + procNome + '(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ",'''','''',"+ window.btoa($(''#beginMonth'').val())+"," + window.btoa($(''#finalMonth'').val())+",'''')"; ' +
          '         location.assign(url); ' +
          '   } ' +
          ' } ';


        End
        else
        begin

          if usaCampanha then
          Begin

            montaNav := montaNav +
            ' function testClicked() { ' +

            '  var url = "ActionCallBackJS:' + procNome + '(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ","+ window.btoa($(''#beginDate'').val())+"," + window.btoa($(''#finalDate'').val())+",'''','''','''')"; ' +
            '  location.assign(url); ' +
            ' } ';

          End
          else
          Begin

            montaNav := montaNav +
            ' function testClicked() { ' +

            '  var url = "ActionCallBackJS:' + procNome + '('''','''',"+ window.btoa($(''#beginDate'').val())+"," + window.btoa($(''#finalDate'').val())+",'''','''','''')"; ' +
            '  location.assign(url); ' +
            ' } ';

          End;

        end;



      montaNav := montaNav +
//      ' </script>' +
////
//      ' <script type="text/javascript"> ' +

        ' const btnFilter = document.getElementById("btnFilter");' +
        ' btnFilter.disabled = true; ' +
        '  ' +
        ' var checkIn = document.getElementById("beginDate"); ' +
        ' var checkOut = document.getElementById("finalDate"); ' +
        '  ' +
        ' checkIn.addEventListener("change", updatedate); ' +
        ' checkOut.addEventListener("change", updateButtonDate); ' +
        '  ' +
        ' function updateButtonDate() { ' +
        '   if (checkOut.value < checkIn.value) { ' +
        '     btnFilter.disabled = true; ' +
        '   } ' +
        '   if (checkOut.value >= checkIn.value) { ' +
        '     btnFilter.disabled = false; ' +
        '   } ' +
        ' } ' +
        ' function updatedate() { ' +
        '   var firstdate = checkIn.value; ' +
        '  ' +
        '   checkOut.min = firstdate; ' +
        '   checkOut.value = firstdate; ' +
        '  ' +
        '   if (checkOut.value < checkIn.value) { ' +
        '     btnFilter.disabled = true; ' +
        '   } ' +
        '   if (checkOut.value >= checkIn.value) { ' +
        '     btnFilter.disabled = false; ' +
        '   } ' +
        ' } ' +

        '  ';

        if usaMes then
        Begin

          montaNav := montaNav +
          ' const chkDate = document.getElementById("chkDate");' +
          ' const chkMonth = document.getElementById("chkMonth");' +
          //CUIDA PARA NÃO DEIXAR O MES FINAL SER MAIOR QUE O INICIAL
          ' var monthIn = document.getElementById("beginMonth"); ' +
          ' var monthOut = document.getElementById("finalMonth"); ' +
          '  ' +
          ' monthIn.addEventListener("change", updateMonthStart); ' +
          ' monthOut.addEventListener("change", updateMonthFinal); ' +
          '  ' +
          ' function updateMonthStart() { ' +
          '  ' +
          '   monthOut.min = monthIn.value; ' +
          '  ' +
          '   if (!monthOut.value){ ' +
          '     monthOut.value = monthIn.value; ' +
          '   } ' +
          '   if (monthOut.value < monthIn.value) { ' +
          '     btnFilter.disabled = true; ' +
          '   } ' +
          '   if (monthOut.value >= monthIn.value) { ' +
          '     btnFilter.disabled = false; ' +
          '   } ' +
          ' } ' +
          ' function updateMonthFinal() { ' +
          '   if (monthOut.value < monthIn.value) { ' +
          '     btnFilter.disabled = true; ' +
          '   } ' +
          '   if (monthOut.value >= monthIn.value) { ' +
          '     btnFilter.disabled = false; ' +
          '   } ' +
          '   if (!monthIn.value){ ' +
          '     btnFilter.disabled = true; ' +
          '   } ' +
          ' } ' +

          //SE CASO O RADIO BUTTON MUDAR, HABILITA OS CAMPOS PARA ESCOLHER DATA/MES
          ' chkDate.addEventListener("click", bloqueiaMes); ' +
          '  ' +
          ' function bloqueiaMes() { ' +
          '   document.getElementById("beginDate").disabled = false; ' +
          '   document.getElementById("finalDate").disabled = false; ' +
          '   document.getElementById("beginMonth").disabled = true; ' +
          '   document.getElementById("finalMonth").disabled = true; ' +
          ' } ' +
          '  ' +
          ' chkMonth.addEventListener("click", bloqueiaData); ' +
          '  ' +
          ' function bloqueiaData() { ' +
          '   document.getElementById("beginDate").disabled = true; ' +
          '   document.getElementById("finalDate").disabled = true; ' +
          '   document.getElementById("beginMonth").disabled = false; ' +
          '   document.getElementById("finalMonth").disabled = false; ' +
          ' } ' +
          '  ' +

          //CHECA O ÚLTIMO USADO E JÁ TRÁS HABILITADO
          ' function chkLastUsed(){  ' +
          ' ' +
          '   if (!checkIn.value && !checkOut.value && monthIn.value !== '''' && monthOut.value !== ''''){  ' +
          '     chkMonth.checked = true;' +
          '     btnFilter.disabled = false; ' +
          '     bloqueiaData(); ' +
          '   } ' +
          '   if (!monthIn.value && !monthOut.value && checkIn.value !== '''' && checkOut.value !== ''''){ ' +
          '     chkDate.checked = true; ' +
          '     btnFilter.disabled = false; ' +
          '     bloqueiaMes(); ' +
          '   } ' +
          ' }' +
          ' ' +
          ' window.onload = chkLastUsed;' +
          ' ' +
          ' ' +
          ' ';

        End
        else
        Begin

          montaNav := montaNav +
          //CHECA O ÚLTIMO USADO E JÁ TRÁS HABILITADO
          ' function chkLastUsed(){  ' +
          ' ' +
          '   if (!checkIn.value && !checkOut.value){  ' +
          '     btnFilter.disabled = true; ' +
          '   } ' +
          '   else{' +
          '     btnFilter.disabled = false; ' +
          '   }' +
          ' ' +
          ' }' +
          ' ' +
          ' window.onload = chkLastUsed;' +
          ' ' +
          ' ' +
          ' ';

        End;

        if usaCampanha then
        Begin

          montaNav := montaNav +
          '   var caCodigo = ""; ' +
          '   var campanhasID = ""; ' +
          '   var campanha; ' +
          '   const cbCampanhas = document.getElementById("cbCampanhas"); ' +
          '   cbCampanhas.addEventListener("change", campCodigo); ' +
          '  ' +
          '   function campCodigo() { ' +
          '     campanha = cbCampanhas.value; ' +
          '     caCodigo = $(this).find(":selected").data("value").codigo; ' +
          '     campanhasID = $(this).find(":selected").data("value").campID; ' +

          '     var url = "ActionCallBackJS:' + procNome + '(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ "," + window.btoa($(''#beginDate'').val())+ "," + window.btoa($(''#finalDate'').val())+ ",'''','''','''')"; ' +
          '			location.assign(url); ' +
          '   } ' +
          '  ';


        End;

      montaNav := montaNav +
      ' </script> ' +

    ' </div> ' +
  ' </div> ';


  Result := montaNav;

end;

function MontaNavBar_Marketing(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
var
  navBar: String;
  montaNav: String;
  montaDrop: String;
  montaRO1, montaRO2, montaRO3, montaRO4: String;
Begin

  if usaCampanha then
    montaDrop := MontaDropCampanhas(campanhaID, 'cbCampanhas', 'visible');

  navBar := NavBarHTML;

  montaNav := navBar +

    //MONTA O MENU LATERAL A ESQUERDA
  ' <div class="offcanvas offcanvas-start" data-bs-scroll="true" data-bs-backdrop="false" tabindex="-1" id="offcanvasScrolling" aria-labelledby="offcanvasScrollingLabel"> ' +
    ' <div class="offcanvas-header"> ' +
     ' <h5 class="offcanvas-title" id="offcanvasScrollingLabel">Filtros</h5> ' +
     ' <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button> ' +
    ' </div> ' +
    ' <div class="offcanvas-body"> ' +

      ' <style> ' +
      //DEIXANDO DATA VERMELHA QUANDO É MENOR
      '   #finalDate:invalid { color: red; } ';

      montaNav := montaNav +
      '   body { ' +
      '     font-family: Roboto; ' +
      '   } ' +
      ' </style> ' +
      ' <form class="row g-3"> ';

      montaNav := montaNav +
      ' <div class="form-check"> ' +
      '   <input ' +
      '     class="form-check-input" ' +
      '     type="radio" ' +
      '     name="radioBtnContatos" ' +
      '     id="chkUltimoRO" ' +
      '     checked="checked" ' +
      '     value="UltimoRO" ' +
      '   /> ' +
      '     <label class="form-check-label" for="chkUltimoRO"> ' +
      '       Usar Último R.O. ' +
      '     </label> ' +
      ' </div> ' +
      ' <div class="form-check"> ' +
      '   <input ' +
      '     class="form-check-input" ' +
      '     type="radio" ' +
      '     name="radioBtnContatos" ' +
      '     id="chkContatos" ' +
      '     value="Contatos" ' +
      '   /> ' +
      '     <label class="form-check-label" for="chkContatos"> ' +
      '       Usar Contatos ' +
      '     </label> ' +
      ' </div> ';


      //DATA COMPLETAS JÁ COM A VALIDAÇÃO
      montaNav := montaNav +
      '   <div class="col-auto"> ' +
      '     <input type="date" name="beginDate" id="beginDate" value="' + dataI + '" style="font-size: 12px; width: 130px"> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <label for="staticText" class="visually"> até </label> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <input type="date" name="finalDate" id="finalDate" min="" value="' + dataF + '" style="font-size: 12px; width: 130px"> ' +
      '   </div> ';

      montaNav := montaNav +
      '   <div class="form-check"> ' +
      '     <input ' +
      '       class="form-check-input" ' +
      '       type="radio" ' +
      '       name="radioBtn" ' +
      '       id="chkOrigem" ' +
      '       checked="checked" ' +
      '       value="Origem" ' +
      '     /> ' +
      '     <label class="form-check-label" for="chkOrigem"> ' +
      '       Comparativo por Origens ' +
      '     </label> ' +
      '   </div> ' +
      '   <div class="form-check"> ' +
      '     <input ' +
      '       class="form-check-input" ' +
      '       type="radio" ' +
      '       name="radioBtn" ' +
      '       id="chkRegiao" ' +
      '       value="Regiao" ' +
      '     /> ' +
      '     <label class="form-check-label" for="chkRegiao"> ' +
      '       Comparativo por Regiões ' +
      '     </label> ' +
      '   </div> ';

      //DROP LISTA DE CAMPANHAS
      if usaCampanha then
        montaNav := montaNav + montaDrop;

      montaNav := montaNav +
      //BUTTON DE CONFIRMAR FILTRO
      '   <div class="col-12"> ' +
      '     <button id="btnFilter" type="button" class="btn btn-primary">Filtrar</button>' +
      '   </div> ' +

      ' </form> ' +

      ' <script> ' +

        ' $(''#btnFilter'').click(function(){ ' +
            ' testClicked(); ' +
        ' }); ';

      montaNav := montaNav +
      ' function testClicked() { ' +
      '   let selectedSize; ' +
      '   for (const radioButton of radioButtons) { ' +
      '     if (radioButton.checked) { ' +
      '       selectedSize = radioButton.value; ' +
      '       break; ' +
      '     } ' +
      '   } ' +
      '  ' +
      '   let selectedChkContatos; ' +
      '   for (const radioButton of radioButtonsContatos) { ' +
      '     if (radioButton.checked) { ' +
      '       selectedChkContatos = radioButton.value; ' +
      '       break; ' +
      '     } ' +
      '   } ' +
      '  ' +
      '   let procAtt; ' +
      '  ' +
      '   if (selectedChkContatos == "Contatos"){ ' +
      '     procAtt = "I"; ' +
      '   } else if (selectedChkContatos == "UltimoRO"){ ' +
      '     procAtt = ""; ' +
      '   } ' +
      '  ' +
      '   if (selectedSize == "Origem") { ' +
      '     var url = ' +
      '       "ActionCallBackJS:GraphFive" + procAtt + "(" + ' +
      '       window.btoa(caCodigo) + ' +
      '       "," + ' +
      '       window.btoa(campanhasID) + ' +
      '       "," + ' +
      '       window.btoa($("#beginDate").val()) + ' +
      '       "," + ' +
      '       window.btoa($("#finalDate").val()) + ' +
      '       ",'''','''','''')"; ' +
      '     location.assign(url); ' +
      '   } else if (selectedSize == "Regiao") { ' +
      '     var url = ' +
      '       "ActionCallBackJS:AttGraphFive" + procAtt + "(" + ' +
      '       window.btoa(caCodigo) + ' +
      '       "," + ' +
      '       window.btoa(campanhasID) + ' +
      '       "," + ' +
      '       window.btoa($("#beginDate").val()) + ' +
      '       "," + ' +
      '       window.btoa($("#finalDate").val()) + ' +
      '       ",'''','''','''')"; ' +
      '     location.assign(url); ' +
      '   } ' +
      '  ' +
      ' } ';

      montaNav := montaNav +
      ' const radioButtons = document.querySelectorAll( ' +
      '   ''input[name="radioBtn"]'' ' +
      ' ); ';

      montaNav := montaNav +
      ' const radioButtonsContatos = document.querySelectorAll( ' +
      '   ''input[name="radioBtnContatos"]'' ' +
      ' ); ';

      montaNav := montaNav +
        ' const btnFilter = document.getElementById("btnFilter");' +
        ' btnFilter.disabled = true; ' +
        '  ' +
        ' var checkIn = document.getElementById("beginDate"); ' +
        ' var checkOut = document.getElementById("finalDate"); ' +
        '  ' +
        ' checkIn.addEventListener("change", updatedate); ' +
        ' checkOut.addEventListener("change", updateButtonDate); ' +
        '  ' +
        ' function updateButtonDate() { ' +
        '   if (checkOut.value < checkIn.value) { ' +
        '     btnFilter.disabled = true; ' +
        '   } ' +
        '   if (checkOut.value >= checkIn.value) { ' +
        '     btnFilter.disabled = false; ' +
        '   } ' +
        ' } ' +
        ' function updatedate() { ' +
        '   var firstdate = checkIn.value; ' +
        '  ' +
        '   checkOut.min = firstdate; ' +
        '   checkOut.value = firstdate; ' +
        '  ' +
        '   if (checkOut.value < checkIn.value) { ' +
        '     btnFilter.disabled = true; ' +
        '   } ' +
        '   if (checkOut.value >= checkIn.value) { ' +
        '     btnFilter.disabled = false; ' +
        '   } ' +
        ' } ' +

        '  ';

        montaNav := montaNav +
        '   var caCodigo = ""; ' +
        '   var campanhasID = ""; ' +
        '   var campanha; ' +
        '   const cbCampanhas = document.getElementById("cbCampanhas"); ' +
        '   cbCampanhas.addEventListener("change", campCodigo); ' +
        '  ' +
        '   function campCodigo() { ' +
        '     campanha = cbCampanhas.value; ' +
        '     caCodigo = $(this).find(":selected").data("value").codigo; ' +
        '     campanhasID = $(this).find(":selected").data("value").campID; ' +
        '   } ' +
        '  ';

      montaNav := montaNav +
      ' </script> ' +

    ' </div> ' +
  ' </div> ';


  Result := montaNav;

end;

function MontaNavBar_Ranking(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campanhaID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
var
  navBar: String;
  montaNav: String;
  montaDrop, montaEquipe: String;
  montaRO1, montaRO2, montaRO3, montaRO4: String;
Begin

  montaDrop := MontaDropEquipes('', 'cbEquipes', 'visible');

  navBar := NavBarHTML;

  montaNav := navBar +

    //MONTA O MENU LATERAL A ESQUERDA
  ' <div class="offcanvas offcanvas-start" data-bs-scroll="true" data-bs-backdrop="false" tabindex="-1" id="offcanvasScrolling" aria-labelledby="offcanvasScrollingLabel"> ' +
    ' <div class="offcanvas-header"> ' +
     ' <h5 class="offcanvas-title" id="offcanvasScrollingLabel">Filtros</h5> ' +
     ' <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button> ' +
    ' </div> ' +
    ' <div class="offcanvas-body"> ' +

      ' <style> ' +
      //DEIXANDO DATA VERMELHA QUANDO É MENOR
      '   #finalDate:invalid { color: red; } ';

      montaNav := montaNav +
      '   body { ' +
      '     font-family: Roboto; ' +
      '   } ' +
      ' </style> ' +
      ' <form class="row g-3"> ';

      //DATA COMPLETAS JÁ COM A VALIDAÇÃO
      montaNav := montaNav +
      '   <div class="col-auto"> ' +
      '     <input type="date" name="beginDate" id="beginDate" value="' + dataI + '" style="font-size: 12px; width: 130px"> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <label for="staticText" class="visually"> até </label> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <input type="date" name="finalDate" id="finalDate" min="" value="' + dataF + '" style="font-size: 12px; width: 130px"> ' +
      '   </div> ';

        montaNav := montaNav + montaDrop;

      montaNav := montaNav +
      //BUTTON DE CONFIRMAR FILTRO
      '   <div class="col-12"> ' +
      '     <button id="btnFilter" type="button" class="btn btn-primary">Filtrar</button>' +
      '   </div> ' +

      ' </form> ' +

      ' <script> ' +

        ' $(''#btnFilter'').click(function(){ ' +
            ' testClicked(); ' +
        ' }); ';

      montaNav := montaNav +
      ' function testClicked() { ' +
      '     var url = ' +
      '       "ActionCallBackJS:GraphSeven( ' +
      '       ''''  ' +
      '       , ' +
      '       '''' ' +
      '       ," + ' +
      '       window.btoa($("#beginDate").val()) + ' +
      '       "," + ' +
      '       window.btoa($("#finalDate").val()) + ' +
      '       ",'''',''''," + window.btoa(equipesID) + ")"; ' +
      '     location.assign(url); ' +
      '  ' +
      ' } ';

      montaNav := montaNav +
        ' const btnFilter = document.getElementById("btnFilter");' +
        ' btnFilter.disabled = true; ' +
        '  ' +
        ' var checkIn = document.getElementById("beginDate"); ' +
        ' var checkOut = document.getElementById("finalDate"); ' +
        '  ' +
        ' checkIn.addEventListener("change", updatedate); ' +
        ' checkOut.addEventListener("change", updateButtonDate); ' +
        '  ' +
        ' function updateButtonDate() { ' +
        '   if (checkOut.value < checkIn.value) { ' +
        '     btnFilter.disabled = true; ' +
        '   } ' +
        '   if (checkOut.value >= checkIn.value) { ' +
        '     btnFilter.disabled = false; ' +
        '   } ' +
        ' } ' +
        ' function updatedate() { ' +
        '   var firstdate = checkIn.value; ' +
        '  ' +
        '   checkOut.min = firstdate; ' +
        '   checkOut.value = firstdate; ' +
        '  ' +
        '   if (checkOut.value < checkIn.value) { ' +
        '     btnFilter.disabled = true; ' +
        '   } ' +
        '   if (checkOut.value >= checkIn.value) { ' +
        '     btnFilter.disabled = false; ' +
        '   } ' +
        ' } ' +

        '  ';

        montaNav := montaNav +
        '   var equipesID = ""; ' +
        '   var equipes; ' +
        '   const cbEquipes = document.getElementById("cbEquipes"); ' +
        '   cbEquipes.addEventListener("change", pegaEquipe); ' +
        '  ' +
        '   function pegaEquipe() { ' +
        '     equipes = cbEquipes.value; ' +
        '     equipesID = $(this).find(":selected").data("value").grupoID; ' +
        '   } ' +
        '  ';

      montaNav := montaNav +
      ' </script> ' +

    ' </div> ' +
  ' </div> ';


  Result := montaNav;

end;

function MontaNavBarRO(procNome: String; dataI: String = ''; dataF: String = '';
  usaMes: Boolean = false; mesI: String = ''; mesF: String = ''; usaCampanha: Boolean = false;
  campID: String = ''; usaOperador: Boolean = false; usaRO: Boolean = false): String;
var
  navBar: String;
  montaNav: String;
  montaDrop: String;
  montaRO1, montaRO2, montaRO3, montaRO4: String;
Begin

  if usaCampanha then
    montaDrop := MontaDropCampanhas(campID, 'cbCampanhas', 'visible');

  if usaRO then
  Begin
    montaRO1 := MontaDropROs;
    montaRO2 := MontaDropROs2;
    montaRO3 := MontaDropROs3;
    montaRO4 := MontaDropROs4;
  End;

  navBar := NavBarHTML;

  montaNav := navBar +
    //MONTA O MENU LATERAL A ESQUERDA
  ' <div class="offcanvas offcanvas-start" data-bs-scroll="true" data-bs-backdrop="false" tabindex="-1" id="offcanvasScrolling" aria-labelledby="offcanvasScrollingLabel"> ' +
    ' <div class="offcanvas-header"> ' +
     ' <h5 class="offcanvas-title" id="offcanvasScrollingLabel">Filtros</h5> ' +
     ' <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button> ' +
    ' </div> ' +
    ' <div class="offcanvas-body"> ' +

      ' <style> ';
      //DEIXANDO DATA VERMELHA QUANDO É MENOR

      if usaMes then
      begin
        //DEIXANDO MÊS VERMELHO QUANDO É MENOR
        montaNav := montaNav +
        '   #finalMonth:invalid { color: red; } ';
      end;

      montaNav := montaNav +
      '   body { ' +
      '     font-family: Roboto; ' +
      '   } ' +
      ' </style> ' +
      ' <form class="row g-3"> ' +

      '   <div class="col-auto"> ' +
      '     <input type="month" id="beginMonth" value="' + mesI + '" style="font-size: 12px; width: 140px"> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <label for="staticText" class="visually"> até </label> ' +
      '   </div> ' +
      '   <div class="col-auto"> ' +
      '     <input type="month" id="finalMonth" value="' + mesF + '" style="font-size: 12px; width: 140px"> ' +
      '   </div> ';


      //DROP LISTA DE CAMPANHAS
      if usaCampanha then
        montaNav := montaNav + montaDrop;


      if usaRO then
      Begin

        montaNav := montaNav +
          montaRO1 +
          montaRO2 +
          montaRO3 +
          montaRO4;

      End;

      montaNav := montaNav +
      ' </form> ' +

      ' <script> ';

      montaNav := montaNav +
      //CUIDA PARA NÃO DEIXAR O MES FINAL SER MAIOR QUE O INICIAL
      ' var monthIn = document.getElementById("beginMonth"); ' +
      ' var monthOut = document.getElementById("finalMonth"); ' +
      '  ' +
      ' monthIn.addEventListener("change", updateMonthStart); ' +
      '  ' +
      ' function updateMonthStart() { ' +
      '  ' +
      '   monthOut.min = monthIn.value; ' +
      '  ' +
      '   if (!monthOut.value){ ' +
      '     monthOut.value = monthIn.value; ' +
      '   } ' +
      ' } ';



      if usaCampanha then
      Begin

        montaNav := montaNav +
        '   var caCodigo = ""; ' +
        '   var campanhasID = ""; ' +
        '   var campanha; ' +
        '   const cbCampanhas = document.getElementById("cbCampanhas"); ' +
        '   cbCampanhas.addEventListener("change", campCodigo); ' +
        '  ' +
        '   function campCodigo() { ' +
        '     campanha = cbCampanhas.value; ' +
        '     caCodigo = $(this).find(":selected").data("value").codigo; ' +
        '     campanhasID = $(this).find(":selected").data("value").campID; ' +
        '     var url = "ActionCallBackJS:GraphTwo(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ",'''','''',"+ window.btoa($(''#beginMonth'').val())+"," + window.btoa($(''#finalMonth'').val())+",'''')"; ' +
        '			location.assign(url); ' +
        '   } ' +
        '  ';
      End;

      if usaRO then
      Begin

        montaNav := montaNav +
          ' var roID1 = ""; ' +
          ' var roID2 = ""; ' +
          ' var roID3 = ""; ' +
          ' var roID4 = ""; ' +
          ' const cbRO1 = document.getElementById("cbROs1"); ' +
          ' cbRO1.addEventListener("change", ro); ' +
          ' function ro() { ' +
          '  roID1 = $(this).find(":selected").data("value").roID;' +
          '     var url1 = "ActionCallBackJS:AttGraphTwo(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ",'''','''',"+ window.btoa($(''#beginMonth'').val())+"," + window.btoa($(''#finalMonth'').val())+",''''," + window.btoa(roID1)+ ")"; ' +
          '			location.assign(url1); ' +
          ' } ' +

          ' const cbRO2 = document.getElementById("cbROs2"); ' +
          ' cbRO2.addEventListener("change", ro2); ' +
          ' function ro2() { ' +
          '  roID2 = $(this).find(":selected").data("value").roID;' +
          '     var url2 = "ActionCallBackJS:AttGraphTwoII(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ",'''','''',"+ window.btoa($(''#beginMonth'').val())+"," + window.btoa($(''#finalMonth'').val())+",''''," + window.btoa(roID2)+ ")"; ' +
          '			location.assign(url2); ' +
          ' } ' +

          ' const cbRO3 = document.getElementById("cbROs3"); ' +
          ' cbRO3.addEventListener("change", ro3); ' +
          ' function ro3() { ' +
          '  roID3 = $(this).find(":selected").data("value").roID;' +
          '     var url3 = "ActionCallBackJS:AttGraphTwoIII(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ",'''','''',"+ window.btoa($(''#beginMonth'').val())+"," + window.btoa($(''#finalMonth'').val())+",''''," + window.btoa(roID3)+ ")"; ' +
          '			location.assign(url3); ' +
          ' } ' +

          ' const cbRO4 = document.getElementById("cbROs4"); ' +
          ' cbRO4.addEventListener("change", ro4); ' +
          ' function ro4() { ' +
          '  roID4 = $(this).find(":selected").data("value").roID;' +
          '     var url4 = "ActionCallBackJS:AttGraphTwoIV(" + window.btoa(caCodigo)+ "," + window.btoa(campanhasID)+ ",'''','''',"+ window.btoa($(''#beginMonth'').val())+"," + window.btoa($(''#finalMonth'').val())+",''''," + window.btoa(roID4)+ ")"; ' +
          '			location.assign(url4); ' +
          ' } ' +

          ' ' ;


      End;


      montaNav := montaNav +
      ' </script> ' +

    ' </div> ' +
  ' </div> ';


  Result := montaNav;

end;

function MontaDropGrupos(recebeID: String = ''; idHTML: String = ''; visible: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    ' SELECT GruposResumoOperacoes_ID ID, ReGrupoNome NOME ' +
    ' FROM GruposResumoOperacoes WITH(NOLOCK) ' +
    ' WHERE ISNULL(ReAtivo,1) = 1 ORDER BY ReGrupoNome');

  montaDrop :=
    ' <select id="' + idHTML + '" class="form-select form-select-sm" required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px; visibility: ' + visible + ';"> ';

  if recebeID <> '' then
  Begin
    montaDrop := montaDrop +
    '   <option data-value=''{"grupoID":""}'' disabled class="hidden">-- Selecione um Grupo de R.O. --</option> ';
  End
  else
  Begin

    montaDrop := montaDrop +
    '   <option data-value=''{"grupoID":""}'' disabled selected class="hidden">-- Selecione um Grupo de R.O.  --</option> ';

  End;

  for I := 1 to dmDados.qryGeral.RecordCount do
  Begin

    if recebeID <> '' then
    Begin

      if recebeID = dmDados.qryGeral.FieldByName('ID').AsString then
      Begin

        montaDrop := montaDrop +
        ' <option data-value=''{"grupoID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}'' selected> ';
        montaDrop := montaDrop + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option> ';

      End
      else
      Begin

        montaDrop := montaDrop +
        ' <option data-value=''{"grupoID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}''> ';
        montaDrop := montaDrop + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option> ';

      End;


    End
    else
    Begin

      montaDrop := montaDrop +
      ' <option data-value=''{"grupoID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}''> ';
      montaDrop := montaDrop + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option> ';

    End;

    dmDados.qryGeral.Next;

  End;

  montaDrop := montaDrop + ' </select> ';

  Result := montaDrop;

End;

function MontaDropCampanhas(recebeID: String = ''; idHTML: String = ''; visible: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    'SELECT DISTINCT Campanhas_ID, CaCodigo, CaCodigo + '' - '' + CaNome Campanha FROM Campanhas WITH(NOLOCK)  '+
    'Inner Join UsuariosCampanhas With(NOLOCK) ON usCampanhasID = Campanhas_ID And UsUsuariosID='+ dmdados.GB_UsuarioID +
    ' WHERE CaEncerrada = 0 ' +
    ' ORDER BY CaCodigo ');

  montaDrop :=
    ' <select id="' + idHTML + '" class="form-select form-select-sm" required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px; visibility: ' + visible + ';"> ';

  if recebeID <> '' then
  Begin
    montaDrop := montaDrop +
    '   <option data-value=''{"codigo":"","campID":""}'' disabled class="hidden">-- Selecione uma Campanha --</option> ';
  End
  else
  Begin

    montaDrop := montaDrop +
    '   <option data-value=''{"codigo":"","campID":""}'' disabled selected class="hidden">-- Selecione uma Campanha --</option> ';

  End;

  for I := 1 to dmDados.qryGeral.RecordCount do
  Begin

    if recebeID <> '' then
    Begin

      if recebeID = dmDados.qryGeral.FieldByName('Campanhas_ID').AsString then
      Begin

        montaDrop := montaDrop +
        ' <option data-value=''{"codigo":"' + dmDados.qryGeral.FieldByName('CaCodigo').AsString + '","campID":"' + dmDados.qryGeral.FieldByName('Campanhas_ID').AsString + '"}'' selected> ';
        montaDrop := montaDrop + dmDados.qryGeral.FieldByName('Campanha').AsString + ' </option> ';

      End
      else
      Begin

        montaDrop := montaDrop +
        ' <option data-value=''{"codigo":"' + dmDados.qryGeral.FieldByName('CaCodigo').AsString + '","campID":"' + dmDados.qryGeral.FieldByName('Campanhas_ID').AsString + '"}''> ';
        montaDrop := montaDrop + dmDados.qryGeral.FieldByName('Campanha').AsString + ' </option> ';

      End;


    End
    else
    Begin

      montaDrop := montaDrop +
      ' <option data-value=''{"codigo":"' + dmDados.qryGeral.FieldByName('CaCodigo').AsString + '","campID":"' + dmDados.qryGeral.FieldByName('Campanhas_ID').AsString + '"}''> ';
      montaDrop := montaDrop + dmDados.qryGeral.FieldByName('Campanha').AsString + ' </option> ';

    End;

    dmDados.qryGeral.Next;

  End;

  montaDrop := montaDrop + ' </select> ';

  Result := montaDrop;

End;

function MontaDropEquipes(recebeID: String = ''; idHTML: String = ''; visible: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    ' SELECT Equipes_ID ID, EqNomeEquipe NOME ' +
    ' FROM Equipes WITH(NOLOCK) ' +
    ' WHERE ISNULL(EqAtiva,1) = 1 ORDER BY NOME');

  montaDrop :=
    ' <select id="' + idHTML + '" class="form-select form-select-sm" required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px; visibility: ' + visible + ';"> ';

  if recebeID <> '' then
  Begin
    montaDrop := montaDrop +
    '   <option data-value=''{"grupoID":""}'' disabled class="hidden">-- Selecione uma Equipe  --</option> ';
  End
  else
  Begin

    montaDrop := montaDrop +
    '   <option data-value=''{"grupoID":""}'' disabled selected class="hidden">-- Selecione uma Equipe  --</option> ';

  End;

  for I := 1 to dmDados.qryGeral.RecordCount do
  Begin

    if recebeID <> '' then
    Begin

      if recebeID = dmDados.qryGeral.FieldByName('ID').AsString then
      Begin

        montaDrop := montaDrop +
        ' <option data-value=''{"grupoID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}'' selected> ';
        montaDrop := montaDrop + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option> ';

      End
      else
      Begin

        montaDrop := montaDrop +
        ' <option data-value=''{"grupoID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}''> ';
        montaDrop := montaDrop + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option> ';

      End;


    End
    else
    Begin

      montaDrop := montaDrop +
      ' <option data-value=''{"grupoID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}''> ';
      montaDrop := montaDrop + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option> ';

    End;

    dmDados.qryGeral.Next;

  End;

  montaDrop := montaDrop + ' </select> ';

  Result := montaDrop;

End;

function MontaDropOperadores(recebeID: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    ' SELECT DISTINCT Usuarios_ID, UsNome ' +
    ' FROM Usuarios WITH(NOLOCK) ' +
    ' WHERE ISNULL(UsTipoDeUsuario,'''') <> ''S'' AND ISNULL(UsAtivo,0) = 1 ' +
    ' ORDER BY UsNome ');

  montaDrop :=
    ' <select id="cbOperadores" class="form-select form-select-sm" disabled required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px;"> ';

  montaDrop := montaDrop +
  '   <option data-value=''{"opID":""}'' disabled selected class="hidden">-- Selecione um(a) Operador(a) --</option> ';

  montaDrop := montaDrop + ' </select> ';


  montaDrop := montaDrop +
        ' <script> ';

   for I := 1 to dmDados.qryGeral.RecordCount do
   Begin
      montaDrop := montaDrop +
        ' $("#cbOperadores").append(''<option data-value={"opID":"' + dmDados.qryGeral.FieldByName('Usuarios_ID').AsString + '"}> ' + dmDados.qryGeral.FieldByName('UsNome').AsString + ' </option>''); ';

      dmDados.qryGeral.Next;
   End;

  montaDrop := montaDrop +
      ' </script> ';

  Result := montaDrop;

End;

function MontaDropROs(recebeID: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    ' SELECT ResumoDeOperacoes_ID ID, ReNome NOME ' +
    ' FROM ResumoDeOperacoes WITH(NOLOCK) ' +
    ' WHERE ISNULL(ReResumoAtivo,0) = 1 ' +
    ' ORDER BY ReNome ');

  montaDrop :=
    ' <select id="cbROs1" class="form-select form-select-sm" required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px;"> ';

  montaDrop := montaDrop +
  '   <option data-value=''{"roID":""}'' disabled selected class="hidden">-- Selecione um R.O. --</option> ';

  montaDrop := montaDrop + ' </select> ';


  montaDrop := montaDrop +
        ' <script> ';

   for I := 1 to dmDados.qryGeral.RecordCount do
   Begin
      montaDrop := montaDrop +
        ' $("#cbROs1").append(''<option data-value={"roID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}> ' + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option>''); ';

      dmDados.qryGeral.Next;
   End;

  montaDrop := montaDrop +
      ' </script> ';

  Result := montaDrop;

End;

function MontaDropROs2(recebeID: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    ' SELECT ResumoDeOperacoes_ID ID, ReNome NOME ' +
    ' FROM ResumoDeOperacoes WITH(NOLOCK) ' +
    ' WHERE ISNULL(ReResumoAtivo,0) = 1 ' +
    ' ORDER BY ReNome ');

  montaDrop :=
    ' <select id="cbROs2" class="form-select form-select-sm" required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px;"> ';

  montaDrop := montaDrop +
  '   <option data-value=''{"roID":""}'' disabled selected class="hidden">-- Selecione um R.O. --</option> ';

  montaDrop := montaDrop + ' </select> ';


  montaDrop := montaDrop +
        ' <script> ';

   for I := 1 to dmDados.qryGeral.RecordCount do
   Begin
      montaDrop := montaDrop +
        ' $("#cbROs2").append(''<option data-value={"roID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}> ' + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option>''); ';

      dmDados.qryGeral.Next;
   End;

  montaDrop := montaDrop +
      ' </script> ';

  Result := montaDrop;

End;

function MontaDropROs3(recebeID: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    ' SELECT ResumoDeOperacoes_ID ID, ReNome NOME ' +
    ' FROM ResumoDeOperacoes WITH(NOLOCK) ' +
    ' WHERE ISNULL(ReResumoAtivo,0) = 1 ' +
    ' ORDER BY ReNome ');

  montaDrop :=
    ' <select id="cbROs3" class="form-select form-select-sm" required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px;"> ';

  montaDrop := montaDrop +
  '   <option data-value=''{"roID":""}'' disabled selected class="hidden">-- Selecione um R.O. --</option> ';

  montaDrop := montaDrop + ' </select> ';


  montaDrop := montaDrop +
        ' <script> ';

   for I := 1 to dmDados.qryGeral.RecordCount do
   Begin
      montaDrop := montaDrop +
        ' $("#cbROs3").append(''<option data-value={"roID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}> ' + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option>''); ';

      dmDados.qryGeral.Next;
   End;

  montaDrop := montaDrop +
      ' </script> ';

  Result := montaDrop;

End;

function MontaDropROs4(recebeID: String = ''): String;
var
  montaDrop: String;
  I: Integer;
Begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.Open(
    ' SELECT ResumoDeOperacoes_ID ID, ReNome NOME ' +
    ' FROM ResumoDeOperacoes WITH(NOLOCK) ' +
    ' WHERE ISNULL(ReResumoAtivo,0) = 1 ' +
    ' ORDER BY ReNome ');

  montaDrop :=
    ' <select id="cbROs4" class="form-select form-select-sm" required aria-label=".form-select-sm example" style="font-size: 12px; margin-left: 30px; width: 250px; height: 30px;"> ';

  montaDrop := montaDrop +
  '   <option data-value=''{"roID":""}'' disabled selected class="hidden">-- Selecione um R.O. --</option> ';

  montaDrop := montaDrop + ' </select> ';


  montaDrop := montaDrop +
        ' <script> ';

   for I := 1 to dmDados.qryGeral.RecordCount do
   Begin
      montaDrop := montaDrop +
        ' $("#cbROs4").append(''<option data-value={"roID":"' + dmDados.qryGeral.FieldByName('ID').AsString + '"}> ' + dmDados.qryGeral.FieldByName('NOME').AsString + ' </option>''); ';

      dmDados.qryGeral.Next;
   End;

  montaDrop := montaDrop +
      ' </script> ';

  Result := montaDrop;

End;

procedure PreencheTableDash_Semanas;
begin

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.ExecSQL(
  ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_Semanas'') ' +
  ' 	DROP TABLE TempDash_Semanas  ' +
  '  ' +
  ' CREATE TABLE TempDash_Semanas  ' +
  ' 	(id INT IDENTITY(1,1) NOT NULL, DiaSemana varchar(100), Mes0 varchar(10), Mes1 varchar(10), Mes2 VARCHAR(10))  ' +
  '  ' +
  ' DECLARE @I INT = 1 ' +
  ' DECLARE @strSQL VARCHAR(MAX)  ' +
  '  ' +
  ' WHILE @I <= 6 ' +
  ' BEGIN ' +
  '  ' +
  ' 	SET @strSQL = ''INSERT INTO TempDash_Semanas(DiaSemana) '' ' +
  ' 	SET @strSQL = @strSQL + ''VALUES(''''Domingo'''') '' ' +
  '  ' +
  ' 	SET @strSQL = @strSQL + ''INSERT INTO TempDash_Semanas(DiaSemana) '' ' +
  ' 	SET @strSQL = @strSQL + ''VALUES(''''Segunda-Feira'''') '' ' +
  '  ' +
  ' 	SET @strSQL = @strSQL + ''INSERT INTO TempDash_Semanas(DiaSemana) '' ' +
  ' 	SET @strSQL = @strSQL + ''VALUES(''''Terça-Feira'''') '' ' +
  '  ' +
  ' 	SET @strSQL = @strSQL + ''INSERT INTO TempDash_Semanas(DiaSemana) '' ' +
  ' 	SET @strSQL = @strSQL + ''VALUES(''''Quarta-Feira'''') '' ' +
  '  ' +
  ' 	SET @strSQL = @strSQL + ''INSERT INTO TempDash_Semanas(DiaSemana) '' ' +
  ' 	SET @strSQL = @strSQL + ''VALUES(''''Quinta-Feira'''') '' ' +
  '  ' +
  ' 	SET @strSQL = @strSQL + ''INSERT INTO TempDash_Semanas(DiaSemana) '' ' +
  ' 	SET @strSQL = @strSQL + ''VALUES(''''Sexta-Feira'''') '' ' +
  '  ' +
  ' 	SET @strSQL = @strSQL + ''INSERT INTO TempDash_Semanas(DiaSemana) '' ' +
  ' 	SET @strSQL = @strSQL + ''VALUES(''''Sábado'''') '' ' +
  '  ' +
  ' 	EXEC (@strSQL) ' +
  ' 		 ' +
  ' 	SET @I = @I + 1 ' +
  ' 	 ' +
  ' END ');

end;

procedure PreencheTableDash_Ranking;
var
  montaSQL: String;
  I: Integer;
begin

  LimpaQuery(dmDados.qryGeral2);
  dmDados.qryGeral2.Open('SELECT DISTINCT GruposResumoOperacoes_ID ID, ReGrupoNome NOME FROM GruposResumoOperacoes WITH(NOLOCK) ORDER BY NOME');

  montaSQL := ' IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ''TempDash_Ranking'') ' +
  ' 	DROP TABLE TempDash_Ranking  ' +
  '  ' +
  ' CREATE TABLE TempDash_Ranking  ' +
  ' (id int IDENTITY(1,1) NOT NULL, nomeOP varchar(100), recebidos varchar(10), trabalhados varchar(10) ';

  for I := 0 to dmDados.qryGeral2.RecordCount - 1 do
    montaSQL := montaSQL + ', grRO' + dmDados.qryGeral2.FieldByName('ID').AsString + ' varchar(10) ';

  montaSQL := montaSQL + ' ) ';

  LimpaQuery(dmDados.qryGeral);
  dmDados.qryGeral.ExecSQL(montaSQL);

end;
end.
