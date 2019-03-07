unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms,
  Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.OleCtrls, SHDocVw, dxGDIPlusClasses, SMCVersInfo;

type
  Tmain = class(TForm)
    lstdosprawdzenia: TListBox;
    mmowynik: TMemo;
    pnl1: TPanel;
    pnl2: TPanel;
    stat1: TStatusBar;
    btnwczytaj: TButton;
    btnanalizuj: TButton;
    pnl3: TPanel;
    pnl4: TPanel;
    img1: TImage;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    WebBrowser1: TWebBrowser;
    lbl1: TLabel;
    lbl2: TLabel;
    SMVersionInfo1: TSMVersionInfo;
    procedure btnwczytajClick(Sender: TObject);
    procedure btnanalizujClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure WebBrowser1DocumentComplete(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  main: Tmain;
  log, log_nip: TextFile;
  wczytanie_kompletne: Boolean;

implementation

{$R *.dfm}

function czyszczenieNUMER(S: string): string;
var
  i: Integer;
begin
  i := 1;
  while i <= Length(S) do
    if S[i] in ['0'..'9'] then
      Inc(i)
    else
      Delete(S, i, 1);
  Result := S;
end;

function zapisLog(komunikat: string): Boolean;
begin
  if not DirectoryExists(ExtractFilePath(Application.ExeName) + 'log') then
    MkDir(ExtractFilePath(Application.ExeName) + 'log');

  AssignFile(log, ExtractFilePath(Application.ExeName) + 'log\log.txt');
  if not FileExists(ExtractFilePath(Application.ExeName) + 'log\log.txt') then // Sprawdzenie, czy plik istnieje
  begin
    ReWrite(log); // Je¿eli nie - stwórz
    Writeln(log, DateTimeToStr(Now) + #09 + komunikat);
  end
  else
  begin
    Append(log); // je¿eli tak - otwórz do odczytu
    Writeln(log, DateTimeToStr(Now) + #09 + komunikat);
  end;
  CloseFile(log);
  Result := True;
end;

function zapisVAT(komunikat: string): Boolean;
begin
  if not DirectoryExists(ExtractFilePath(Application.ExeName) + 'log') then
    MkDir(ExtractFilePath(Application.ExeName) + 'log');

  AssignFile(log, ExtractFilePath(Application.ExeName) + 'log\vat.txt');
  if not FileExists(ExtractFilePath(Application.ExeName) + 'log\vat.txt') then // Sprawdzenie, czy plik istnieje
  begin
    ReWrite(log); // Je¿eli nie - stwórz
    Writeln(log, DateTimeToStr(Now) + #09 + komunikat);
  end
  else
  begin
    Append(log); // je¿eli tak - otwórz do odczytu
    Writeln(log, DateTimeToStr(Now) + #09 + komunikat);
  end;
  CloseFile(log);
  Result := True;
end;

function zapisNIP(komunikat: string): Boolean;
begin
  if not DirectoryExists(ExtractFilePath(Application.ExeName) + 'log') then
    MkDir(ExtractFilePath(Application.ExeName) + 'log');

  AssignFile(log_nip, ExtractFilePath(Application.ExeName) + 'log\log_NIP_' + DateToStr(Now) + '.txt');
  if not FileExists(ExtractFilePath(Application.ExeName) + 'log\log_NIP_' + DateToStr(Now) + '.txt') then // Sprawdzenie, czy plik istnieje
  begin
    ReWrite(log_nip); // Je¿eli nie - stwórz
    Writeln(log_nip, komunikat);
  end
  else
  begin
    Append(log_nip); // je¿eli tak - otwórz do odczytu
    Writeln(log_nip, komunikat);
  end;
  CloseFile(log_nip);
  Result := True;
end;

procedure Tmain.btnanalizujClick(Sender: TObject);
var
  suma, suma_kontrolna: Integer;
  licznik_nip, i: Integer;
  licznik_czynny, licznik_zwolniony, licznk_brak_vat: Integer;
  nip_cyfry, podatnik: string;
begin
  licznik_nip := 1;

  if lstdosprawdzenia.Items.Count > 0 then
  begin
    zapisLog('uruchomiono sprawdzanie');
    for i := 0 to lstdosprawdzenia.Items.Count - 1 do
    begin
      nip_cyfry := czyszczenieNUMER(lstdosprawdzenia.Items[i]); // pozostawienie tylko cyfr

      if Length(nip_cyfry) = 10 then  // czy nip ma 10 znaków
      begin
        suma := (StrToInt(nip_cyfry[1]) * 6) + (StrToInt(nip_cyfry[2]) * 5) + (StrToInt(nip_cyfry[3]) * 7) + (StrToInt(nip_cyfry
          [4]) * 2) + (StrToInt(nip_cyfry[5]) * 3) + (StrToInt(nip_cyfry[6]) * 4) + (StrToInt(nip_cyfry[7]) * 5) + (StrToInt
          (nip_cyfry[8]) * 6) + (StrToInt(nip_cyfry[9]) * 7);
        suma_kontrolna := suma mod 11;  // wyliczenie sumy kontrolnej nip-u

        if (StrToInt(nip_cyfry[10]) = suma_kontrolna) then
        begin
          wczytanie_kompletne := False;
          WebBrowser1.OleObject.Document.GetElementById('b-7').Value := nip_cyfry;
          WebBrowser1.OleObject.Document.GetElementByID('b-8').Click;

          repeat
            Application.ProcessMessages;
          until wczytanie_kompletne;

          podatnik := WebBrowser1.OleObject.Document.GetElementById('b-3').innertext;

          if Pos('podatnik VAT czynny', podatnik) > 0 then
          begin
            mmowynik.Lines.Add('NIP: ' + lstdosprawdzenia.Items[i] + ' odpowiedŸ: Czynny podatnik VAT');
            Inc(licznik_czynny);
            zapisNIP(Format('%.4d', [licznik_nip]) + '. Podatnik o numerze VAT: ' + lstdosprawdzenia.Items[i] +
              ' jest zarejestrowany jako p³atnik VAT.');
          end;

          if Pos('podatnik VAT zwolniony', podatnik) > 0 then
          begin
            mmowynik.Lines.Add('NIP: ' + lstdosprawdzenia.Items[i] + ' odpowiedŸ: Czynny podatnik VAT - zwolniony');
            Inc(licznik_zwolniony);
            zapisNIP(Format('%.4d', [licznik_nip]) + '. Podatnik o numerze VAT: ' + lstdosprawdzenia.Items[i] +
              ' jest zarejestrowany jako p³atnik VAT ZWOLNIONY.');
          end;

          if Pos('nie jest zarejestrowany jako podatnik VAT', podatnik) > 0 then
          begin
            mmowynik.Lines.Add('NIP: ' + lstdosprawdzenia.Items[i] + ' odpowiedŸ: UWAGA: nie zarejestrowany jako podatnik VAT');
            licznk_brak_vat := licznk_brak_vat + 1;

            zapisNIP(Format('%.4d', [licznik_nip]) + '. Podatnik o numerze VAT: ' + lstdosprawdzenia.Items[i] +
              ' NIE jest zarejestrowany jako p³atnik VAT.');
          end;

          wczytanie_kompletne := False;
          WebBrowser1.OleObject.Document.GetElementByID('b-9').Click;

          repeat
            Application.ProcessMessages;
          until wczytanie_kompletne;

          inc(licznik_nip);
        end
        else
        begin
          zapisNIP(Format('%.4d', [licznik_nip]) + '. NIP: ' + lstdosprawdzenia.Items[i] + ' ma nieprawid³ow¹ sumê kontroln¹.');
          inc(licznik_nip);
        end;
      end
      else
      begin
        zapisNIP(Format('%.4d', [licznik_nip]) + '. NIP: ' + lstdosprawdzenia.Items[i] + ' ma nieprawid³ow¹ d³ugoœæ!');
        Inc(licznik_nip);
      end;
    end;

    zapisLog('Przeanalizowano: ' + IntToStr(lstdosprawdzenia.Items.Count) + ' numerów NIP');
    zapisLog('Czynnych podatników VAT by³o: ' + IntToStr(licznik_czynny));
    zapisLog('Czynnych zwolnionych podatników VAT by³o: ' + IntToStr(licznik_zwolniony));
    zapisLog('nie zarejestrowany jako podatnik VAT by³o: ' + IntToStr(licznk_brak_vat));
  end
  else
  begin
    mmowynik.Lines.Add('brak rekordów do analizy!');
    zapisLog('brak rekordów do analizy!');
  end;
end;

procedure Tmain.btnwczytajClick(Sender: TObject);
begin
  OpenDialog1.InitialDir := ExtractFilePath(Application.ExeName);
  OpenDialog1.Execute();
  if FileExists(OpenDialog1.FileName) then
  begin
    lstdosprawdzenia.Items.LoadFromFile(OpenDialog1.FileName);
    zapisLog('Wczytano plik: ' + OpenDialog1.FileName);
  end;
end;

procedure Tmain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  case Application.MessageBox('Czy zakoñczyæ program?', 'Uwaga!', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) of
    IDYES:
      begin
        zapisLog('Wy³¹czono program');
        Application.Terminate;
      end;
    IDNO:
      begin
        CanClose := False;
      end;
  end;
end;

procedure Tmain.FormCreate(Sender: TObject);
begin
  zapisLog('Uruchomiono program');
  stat1.Panels[0].Text := 'wersja: ' + SMVersionInfo1.FileVersion;
end;

procedure Tmain.FormShow(Sender: TObject);
begin
  mmowynik.Text := '';
  WebBrowser1.Navigate('https://ppuslugi.mf.gov.pl/?link=VAT&;');
end;

procedure Tmain.WebBrowser1DocumentComplete(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
begin
  wczytanie_kompletne := True;
end;

end.

