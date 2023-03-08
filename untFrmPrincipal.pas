unit untFrmPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, AdvPanel, System.StrUtils,
  System.Types;

type
  TFrmPrincipal = class(TForm)
    mmoEnviar: TMemo;
    mmoReceber: TMemo;
    AdvPanel1: TAdvPanel;
    Label1: TLabel;
    btnEnviar: TButton;
    lblEnviar: TLabel;
    lblReceber: TLabel;
    btnCopiar: TButton;
    OpenDialog1: TOpenDialog;
    btnAbrirArquivo: TButton;
    btnLimparReceber: TButton;
    lblQtdLinhas: TLabel;
    btnLimparEnviar: TButton;
    procedure btnEnviarClick(Sender: TObject);
    procedure btnCopiarClick(Sender: TObject);
    procedure btnAbrirArquivoClick(Sender: TObject);
    procedure btnLimparReceberClick(Sender: TObject);
    procedure mmoReceberChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnLimparEnviarClick(Sender: TObject);
    procedure mmoEnviarChange(Sender: TObject);
  private
    procedure VerificarReceber;
    function Sanitize(txt: String = ''): String;
    procedure LinhasCountUpdate;
    procedure VerificarEnviar;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmPrincipal: TFrmPrincipal;

implementation

{$R *.dfm}

procedure TFrmPrincipal.btnAbrirArquivoClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
  begin
    OpenDialog1.Title := 'SELECIONE O ARQUIVO';
    OpenDialog1.DefaultExt := '.txt';

    if SplitString(OpenDialog1.FileName, '.')[1] = 'txt' then
    begin
      mmoEnviar.Lines.LoadFromFile(OpenDialog1.FileName, TEncoding.UTF8);
    end
      else
      begin
        MessageDlg('Suporte apenas para arquivos .txt', mtInformation, [mbOK], 0);
      end;
  end
    else
    begin
      MessageDlg('Não foi possível carregar o arquivo', mtWarning, [mbOK], 0);
    end;
end;

procedure TFrmPrincipal.btnCopiarClick(Sender: TObject);
begin
  mmoReceber.SelectAll;
  mmoReceber.CopyToClipboard;
  mmoReceber.Lines.Clear;
  MessageDlg('Copiado para a área de transferência', mtInformation, [mbOK], 0);
end;

procedure TFrmPrincipal.btnEnviarClick(Sender: TObject);
var
  linha, tmpString: String;
  I, J: Integer;
  linhaQtd, tamanhoLogradouro, id: Integer;

  nome,         cpfCnpj,
  nomeFantasia, RG,
  telefone,     celular,
  logradouro,   email,
  complemento,  bairro,
  cep,          cidade,
  numero,       tipoPessoa: String;

  linhas:           TStringDynArray;
  logradouroNumero: TStringDynArray;
  enderecoArray:    TStringDynArray;
begin
  linhas := SplitString(mmoEnviar.Text, #13);
  Try
    try
      linhaQtd := 1;
      id := 0;
      for I := 0 to (Length(linhas) -1) do
      begin
        linha := Trim(linhas[I]);
        if (Length(linhas) > 0) and ((linha = '') or (linhaQtd > 6)) then
        begin
          linhaQtd := 1;
          Continue;
        end;

        case linhaQtd of
          1:
          begin
            id   := StrToIntDef(SplitString(linha, ' ')[0], 0);
            nome := linha.Substring(intToStr(id).Length);
          end;

          2:
          begin
            cpfCnpj := StringReplace(linha, 'CPF/CNPJ:', '', [rfReplaceAll, rfIgnoreCase]);
          end;

          3:
          begin
            nomeFantasia := linha;
          end;

          4:
          begin
            tmpString := StringReplace(linha, 'RG/I.E:', '', [rfReplaceAll, rfIgnoreCase]);
            tmpString := StringReplace(tmpString, 'Fone:', '', [rfReplaceAll, rfIgnoreCase]);
            RG        := SplitString(tmpString, ' ')[0];
            telefone  := SplitString(tmpString, ' ')[1];
          end;

          5:
          begin
            tmpString := StringReplace(linha, 'Celular:', '', [rfReplaceAll, rfIgnoreCase]);
            tmpString := StringReplace(tmpString, 'E-Mail:', '', [rfReplaceAll, rfIgnoreCase]);
            celular   := SplitString(tmpString, ' ')[0];
            email     := SplitString(tmpString, ' ')[1];
          end;

          6:
          begin
            enderecoArray := SplitString(StringReplace(linha, 'Endereço:', '', [rfReplaceAll, rfIgnoreCase]), ',');

            if Length(enderecoArray) <> 5 then
              Continue;

            logradouroNumero := SplitString(Trim(enderecoArray[0]), ' ');
            numero := IfThen(StrToIntDef(logradouroNumero[Length(logradouroNumero)-1], 0) > 0, logradouroNumero[Length(logradouroNumero)-1], 'S/N');

            if numero = 'S/N' then
            begin
              tamanhoLogradouro := Length(logradouroNumero) -1;
            end
              else
              begin
                tamanhoLogradouro := Length(logradouroNumero) - 2;
              end;

            logradouro := '';
            for J := 0 to tamanhoLogradouro do
            begin
              logradouro := logradouro + ' ' + logradouroNumero[J];
            end;

            complemento := enderecoArray[1];
            bairro      := enderecoArray[2];
            cep         := enderecoArray[3];
            cidade      := enderecoArray[4];
          end;
        end;

        if linhaQtd = 6 then
        begin
          tipoPessoa := IfThen(cpfCnpj.Length > 11, 'J', 'F');

          nome := Sanitize(nome);
          complemento := Sanitize(complemento);
          cpfCnpj := Sanitize(cpfCnpj);
          logradouro := Sanitize(logradouro);
          numero := Sanitize(numero);
          bairro := Sanitize(bairro);
          cep := Sanitize(cep);
          cidade := Sanitize(cidade);

          mmoReceber.Lines.Add('INSERT INTO clients (id, name, cpf/cnpj, person_type, fantasy_nickname, public_place, street_number, complement, district, zipcode, cidade) VALUES ('+
          IfThen(id > 0, IntToStr(id), 'NULL') + ', ' + nome + ', ' + cpfCnpj + ', ' + QuotedStr(tipoPessoa) + ', ' + logradouro + ', ' + numero + ', ' + complemento + ', ' + bairro + ', ' + cep + ', ' + cidade +');');
        end;

        linhaQtd := linhaQtd + 1;
      end;
    except
      on E: Exception do
      begin
        if MessageDlg('Erro no sistema', mtError, [mbClose], 0) = mrClose then
        begin
          Close;
          Abort;
        end;

        Close;
        Abort;
      end;
    end;
  Finally

  End;
end;

procedure TFrmPrincipal.btnLimparReceberClick(Sender: TObject);
begin
  mmoReceber.Lines.Clear;
  VerificarReceber();
  LinhasCountUpdate();
end;

procedure TFrmPrincipal.btnLimparEnviarClick(Sender: TObject);
begin
  mmoEnviar.Lines.Clear;
  VerificarEnviar();
end;

procedure TFrmPrincipal.FormCreate(Sender: TObject);
begin
  LinhasCountUpdate();
  VerificarEnviar();
  VerificarReceber();
end;

procedure TFrmPrincipal.mmoEnviarChange(Sender: TObject);
begin
  VerificarEnviar();
end;

procedure TFrmPrincipal.mmoReceberChange(Sender: TObject);
begin
  VerificarReceber();
  LinhasCountUpdate();
end;

procedure TFrmPrincipal.VerificarReceber();
begin
  if not Trim(mmoReceber.Text).IsEmpty then
  begin
    btnCopiar.Enabled        := True;
    btnLimparReceber.Enabled := True;
  end
    else
    begin
      btnCopiar.Enabled        := False;
      btnLimparReceber.Enabled := False;
    end;
end;

procedure TFrmPrincipal.VerificarEnviar;
begin
  if not Trim(mmoEnviar.Text).IsEmpty then
  begin
    btnEnviar.Enabled       := True;
    btnLimparEnviar.Enabled := True;
  end
    else
    begin
      btnEnviar.Enabled       := False;
      btnLimparEnviar.Enabled := False;
    end;
end;

procedure TFrmPrincipal.LinhasCountUpdate;
begin
   lblQtdLinhas.Caption := IntToStr(Length(SplitString(mmoReceber.Text, #13))) + IfThen(Length(SplitString(mmoReceber.Text, #13)) = 1, ' Linha' , ' Linhas');
end;

function TFrmPrincipal.Sanitize(txt: String = ''): String;
begin
  txt := Trim(txt);
  Result := IfThen(txt.IsEmpty, 'NULL', QuotedStr(txt));
end;

end.
