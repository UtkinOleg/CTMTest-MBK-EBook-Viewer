unit exebook;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, OleCtrls, SHDocVw, ExtCtrls, ComCtrls, ImgList, BookSav, MSHTML_TLB, StrUtils,
  ToolWin, ActiveX, StdCtrls, TB97Ctls, RXSplit, ClipBrd,
  IAsemiPanel, Registry, ShellApi, ButtonComps, Pass, OPenf;

const
   MainTitle = 'Электронный учебник';

type
  TMainForm = class(TForm)
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    ImageList1: TImageList;
    sb: TStatusBar;
    wb2: TWebBrowser;
    ImageList2: TImageList;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton3: TToolButton;
    N5: TMenuItem;
    N6: TMenuItem;
    ToolButton5: TToolButton;
    ToolButtonFind: TToolButton;
    Panel1: TPanel;
    ItemsTree: TTreeView;
    PanelSeek: TPanel;
    Splitter1: TSplitter;
    ToolButton2: TToolButton;
    ToolButton6: TToolButton;
    N7: TMenuItem;
    N8: TMenuItem;
    Label3: TLabel;
    Edit1: TEdit;
    Button1: TButton;
    Panel2: TIAsemiPanel;
    CloseBtn: TToolbarButton97;
    ImageList3: TImageList;
    Panel3: TIAsemiPanel;
    CloseBtn2: TToolbarButton97;
    N9: TMenuItem;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ReopenBookMenu: TPopupMenu;
    PopupMenu2: TPopupMenu;
    N10: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    Panel6: TPanel;
    Image5: TImage;
    ImageButton17: TImageButton;
    ImageButton18: TImageButton;
    ImageButton19: TImageButton;
    ImageButton20: TImageButton;
    ImageButton21: TImageButton;
    ImageButton22: TImageButton;
    ImageButton23: TImageButton;
    Image6: TImage;
    ImageButton24: TImageButton;
    ToolButton4: TToolButton;
    ComboBox1: TComboBox;
    N20: TMenuItem;
    N21: TMenuItem;
    procedure N2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure twClose(Sender: TObject);
    procedure ItemsTreeClick(Sender: TObject);
    procedure wb2DocumentComplete(Sender: TObject; const pDisp: IDispatch;
      var URL: OleVariant);
    procedure wb2BeforeNavigate2(Sender: TObject; const pDisp: IDispatch;
      var URL, Flags, TargetFrameName, PostData, Headers: OleVariant;
      var Cancel: WordBool);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ItemsTreeCollapsing(Sender: TObject; Node: TTreeNode;
      var AllowCollapse: Boolean);
    procedure ItemsTreeExpanding(Sender: TObject; Node: TTreeNode;
      var AllowExpansion: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure wb2CommandStateChange(Sender: TObject; Command: Integer;
      Enable: WordBool);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure CloseBtnClick(Sender: TObject);
    procedure ItemsTreeContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure ItemsTreeAdvancedCustomDrawItem(Sender: TCustomTreeView;
      Node: TTreeNode; State: TCustomDrawState; Stage: TCustomDrawStage;
      var PaintImages, DefaultDraw: Boolean);
    procedure ToolButtonFindClick(Sender: TObject);
    procedure Splitter1Moved(Sender: TObject);
    procedure Splitter2Moved(Sender: TObject);
    procedure CloseBtn2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure ToolButton8Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure ToolBar1CustomDraw(Sender: TToolBar; const ARect: TRect;
      var DefaultDraw: Boolean);
    procedure ToolBar1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormResize(Sender: TObject);
    procedure ToolBar1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ItemsTreeCustomDraw(Sender: TCustomTreeView;
      const ARect: TRect; var DefaultDraw: Boolean);
    procedure ImageButton5Click(Sender: TObject);
    procedure ImageButton4Click(Sender: TObject);
    procedure ImageButton2Click(Sender: TObject);
    procedure ImageButton3Click(Sender: TObject);
    procedure ImageButton1Click(Sender: TObject);
    procedure ImageButton7Click(Sender: TObject);
    procedure ImageButton6Click(Sender: TObject);
    procedure Image2Click(Sender: TObject);
    procedure Image2MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure N19Click(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure ImageButton16Click(Sender: TObject);
    procedure ImageButton15Click(Sender: TObject);
    procedure ImageButton24Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ItemsTreeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N20Click(Sender: TObject);
  private
    { Private declarations }
    NextWindow: HWND;
    procedure WMChangeCBChain(var Msg: TWMChangeCBChain); message
    WM_CHANGECBCHAIN;
    procedure WMDrawClipboard(var Msg: TWMDrawClipboard); message
    WM_DRAWCLIPBOARD;
  public
    { Public declarations }
    TmpList : TStringList;
    ResPageNavigate : boolean;
    FirstGen, URLGo : boolean;
    ContentIndex, TestTreeIndex : integer;
    HomeIndex : integer;
    TableColor, FIO, ORG : string;
    TestList : TStringList;
    FileList : TStringList;
    dt : TDateTime;
    BitMapLogo : TBitMap;
    function  OpenBFile(F : TStream):boolean;
    procedure AddChildNodes(PQ: THTMLDoc; PNode: TTreeNode; PTree: TTreeView; AddLeaf, AddTests : boolean);
    procedure LoadHTMLDoc(H: THTMLDoc; var F : TStream; ver : integer);
    procedure DeployPictures(HTMLDoc : THTMLDoc);
    procedure Navigate(navmode: boolean);
    procedure Nav(HTMLDoc: THTMLDoc);
    procedure Nav2(HTMLDoc: THTMLDoc);
    function  CreateBookMarkMenu(Reg: TRegistry):boolean;
    procedure BookPopupHandler(Sender: TObject);
    procedure GenerateBookMarkHTML(Reg: TRegistry);
    procedure ShowBookMark;
    procedure DeleteBookMark(s: string);
    procedure DeleteResult(s: string);
    procedure SeekMainPage;
    procedure TreeSaveToFile;
    procedure TreeLoadFromFile;
    procedure ViewTree;
    procedure GenerateTestPage(HTMLDoc : THTMLDoc);
    procedure GenerateFilePage(HTMLDoc : THTMLDoc);
    procedure AddConfig(fname : string);
    procedure SaveResults;
  end;

var
  MainForm: TMainForm;
  NewDoc : Boolean;
  cBuffer: array[0..MAX_PATH] of Char;
  Disp: IDispatch;
  Editor: IHTMLDocument2;
  HookID: THandle;

implementation

uses NewBookMarkUnit;

{$R *.DFM}
{$R manifest.res}

function GetTmpDir: string;
 var
   Buffer: array[0..MAX_PATH] of Char;
 begin
   GetTempPath(SizeOf(Buffer) - 1, Buffer);
   Result := StrPas(Buffer);
 end;


procedure TMainForm.WMChangeCBChain(var Msg: TWMChangeCBChain);
 begin
inherited;
   Msg.Result := 0;
   if Msg.Remove = NextWindow then
     NextWindow := Msg.Next
   else
     SendMessage(NextWindow, WM_CHANGECBCHAIN, Msg.Remove, Msg.Next);
 end;

procedure TMainForm.WMDrawClipboard(var Msg: TWMDrawClipboard);
 begin
  inherited;
   try
     // Очистим буфер обмена
     Clipboard.Clear;
   finally
     SendMessage(NextWindow, WM_DRAWCLIPBOARD, 0, 0);
   end;
 end;

function MouseProc(nCode: Integer; wParam, lParam: Longint): Longint; stdcall;
var
  szClassName: array[0..255] of Char;
const
  ie_name = 'Internet Explorer_Server';
begin
  case nCode < 0 of
    True:
      Result := CallNextHookEx(HookID, nCode, wParam, lParam)
  else
    case wParam of
      WM_RBUTTONDOWN,
        WM_RBUTTONUP:
        begin
          GetClassName(PMOUSEHOOKSTRUCT(lParam)^.HWND, szClassName,
            SizeOf(szClassName));
          if lstrcmp(@szClassName[0], @ie_name[1]) = 0 then
            Result := HC_SKIP
          else
            Result := CallNextHookEx(HookID, nCode, wParam, lParam);
        end
    else
      Result := CallNextHookEx(HookID, nCode, wParam, lParam);
    end;
  end;
end;

procedure TMainForm.LoadHTMLDoc(H: THTMLDoc; var F : TStream; ver : integer);
var
  Item1,Item2,Item3, Richlen, RA, I: Integer;
  k, QT, QP, Cnt, J, L, len: Integer;
  B: Byte;
  buffer : PChar;
  M: TStream;
  l1, l2, s, s1, s2, s3 : string;
  zb, RBuffer: PByte;
  M1, F1, MemIn, MemOut : TMemoryStream;
begin
 with H do
 begin
   try
    F.Read(len,4);
    buffer := AllocMem(len+1);
    F.Read(buffer^,len);
    buffer[len]:=Chr(0);
    Nam := StrPas(buffer);
    FreeMem(buffer);

    F.Read(ID,4); // id
    F.Read(QParent,4); // родитель
    F.Read(HTMLType,4);
    F.Read(b,1);
    MainPage := b=1;

    F.Read(len,4);
    RBuffer := AllocMem(len);
    F.Read(RBuffer^,len);
    HTMLText.Write(RBuffer^,len);
    FreeMem(RBuffer);

    F.Read(len,4);
    RBuffer := AllocMem(len);
    F.Read(RBuffer^,len);
    Pictures.Write(RBuffer^,len);
    FreeMem(RBuffer);

    if ver>=2 then
    begin
     F.Read(len,4);
     RBuffer := AllocMem(len);
     F.Read(RBuffer^,len);
     AnyFiles.Write(RBuffer^,len);
     FreeMem(RBuffer);
    end;

    if (ver>=4) and (HTMLType=5) then
    begin
     F.Read(len,4);
     buffer := AllocMem(len+1);
     F.Read(buffer^,len);
     buffer[len]:=Chr(0);
     Ext := StrPas(buffer);
     FreeMem(buffer);
    end;

    F.Read(len,4);
    Cnt := len;

    Children.Clear;
    for I := 0 to Cnt - 1 do
    begin
     Children.Add(THTMLDoc.Create);
     LoadHTMLDoc(THTMLDoc(Children.Items[Children.Count-1]),F,ver);
    end;
   except
   end;
 end;
end;

procedure TMainForm.N2Click(Sender: TObject);
begin
 Close
end;

procedure TMainForm.ShowBookMark;
var
 Reg:TRegistry;
begin
  ResPageNavigate := false;

  Reg:=TRegistry.Create;
  try
   Reg.RootKey := HKEY_CURRENT_USER;
   if Reg.OpenKey(BookRegStr+Caption,True) then
    if CreateBookMarkMenu(Reg) then
     GenerateBookMarkHTML(Reg);
  finally
   Reg.CloseKey;
  end;
end;

procedure TMainForm.FormActivate(Sender: TObject);
var
 rs : TResourceStream;
 Reg:TRegistry;
begin

 ShowBookMark;
 WindowState := wsMaximized;


 Reg:=TRegistry.Create;
 try
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey('\Software\CTMTest\EBook\'+Caption,False) then
  begin
{   if Reg.ValueExists('FullScreen') then
   begin
    N9.Checked := Reg.ReadBool('FullScreen');
    if N9.Checked then
     BorderStyle := bsNone
    else
     BorderStyle := bsSizeable;
   end; }
   if Reg.ValueExists('ContentWidth') then
   begin
    Panel1.Width := Reg.ReadInteger('ContentWidth');
    CloseBtn.Left := Panel2.Width - 23;
    ComboBox1.Width := Panel2.Width - 28;
    if Panel2.Width > 125 then
     Panel2.Caption := 'Содержание'
    else
     Panel2.Caption := '';
   end;
  end;
  Reg.CloseKey;
 finally
  Reg.Free;
 end;
//  ItemsTree.Width := Trunc(ClientWidth/4);
  Repaint;
end;

procedure TMainForm.N4Click(Sender: TObject);
begin
 PanelSeek.Visible := false;
 ToolButtonFind.Down := false;
 N4.Checked := not N4.Checked;
 ToolButton2.Down := N4.Checked;
 Splitter1.Visible := N4.Checked;
 Panel1.Visible := N4.Checked;
end;

procedure TMainForm.twClose(Sender: TObject);
begin
 N4.Checked := false;
end;

procedure TMainForm.AddChildNodes(PQ: THTMLDoc; PNode: TTreeNode; PTree: TTreeView; AddLeaf, AddTests : boolean);
var
  J, I: Integer;
  Node: TTreeNode;
begin
   case PQ.HTMLType of
        0 :
                   begin
                      Node := PTree.Items.AddChildObject(PNode, PQ.Nam, PQ);
                      Node.ImageIndex := 0;
                      Node.SelectedIndex := 0;
                   end;
        1:         if AddLeaf then
                   begin
                      Node := PTree.Items.AddChildObject(PNode, PQ.Nam, PQ);
                      Node.ImageIndex := 2;
                      Node.SelectedIndex := 2;
                   end;

        2:         if AddTests then
                    TestList.AddObject(PQ.Nam,PQ);
        5:         if AddTests then
                    FileList.AddObject(PQ.Nam,PQ);
   end;
  for I := 0 to PQ.Children.Count - 1 do
    AddChildNodes(THTMLDoc(PQ.Children.Items[I]), Node, PTree, AddLeaf, AddTests);
end;

function TMainForm.OpenBFile;
var
 F1 : TFileStream;
 ver, j,len, i : integer;
 buffer, buffer2 : PChar;
 M, MO : TMemoryStream;
 s : string;
 OldBook, QT, Cnt : integer;
 Node: TTreeNode;
 ContRepeat, ex : boolean;
 zb : PByte;
 HTMLDoc : THTMLDoc;
 b : byte;

begin

try
 DocsBook.CreateBook;
 with DocsBook.ActiveBook do
  begin
  Screen.Cursor := crHourGlass;
  F.Position := 0;
  F.Read(ver,4); // версия 1, 2.0
  if ver > BookVer then
   begin
    DocsBook.Delete(DocsBook.CurBook);
    Dec(DocsBook.CurBook);
    Screen.Cursor := crDefault;
    Result := False;
    Exit;
   end;
  F.Seek(18,soFromCurrent);

  F.Read(len,4);
  buffer := AllocMem(len+1);
  buffer[len] := Chr(0);
  F.Read(buffer^,len);
  buffer[len]:=Chr(0);
  Nam := StrPas(buffer);
  Caption := Nam;
  Application.Title := Nam;
  FreeMem(buffer);

  F.Read(len,4);
  buffer := AllocMem(len+1);
  buffer[len] := Chr(0);
  F.Read(buffer^,len);
  buffer[len]:=Chr(0);
  FIO := StrPas(buffer);
  FreeMem(buffer);

  F.Read(len,4);
  buffer := AllocMem(len+1);
  buffer[len] := Chr(0);
  F.Read(buffer^,len);
  buffer[len]:=Chr(0);
  ORG := StrPas(buffer);
  FreeMem(buffer);

  F.Read(len,4);
  F.Seek(len, soFromCurrent);

  if ver>=2 then
  begin
   F.Read(len,4);
   buffer := AllocMem(len+1);
   buffer[len] := Chr(0);
   F.Read(buffer^,len);
   buffer[len]:=Chr(0);
   Email := StrPas(buffer);
   FreeMem(buffer);

   F.Read(len,4);
   buffer := AllocMem(len+1);
   buffer[len] := Chr(0);
   F.Read(buffer^,len);
   buffer[len]:=Chr(0);
   Place := StrPas(buffer);
   FreeMem(buffer);

   F.Read(len,4);
   buffer := AllocMem(len+1);
   buffer[len] := Chr(0);
   F.Read(buffer^,len);
   buffer[len]:=Chr(0);
   Phone := StrPas(buffer);
   FreeMem(buffer);

   F.Read(len,4);
   buffer := AllocMem(len+1);
   buffer[len] := Chr(0);
   F.Read(buffer^,len);
   buffer[len]:=Chr(0);
   Version := StrPas(buffer);
   FreeMem(buffer);

   F.Read(len,4);
   buffer := AllocMem(len+1);
   buffer[len] := Chr(0);
   F.Read(buffer^,len);
   buffer[len]:=Chr(0);
   Comment := StrPas(buffer);
   FreeMem(buffer);
  end;

  if ver>=3 then
  begin
   F.Read(b,1);
   Locked := b;
   F.Read(b,1);
   ExtMode := b;
  end
  else
  begin
   Locked := 0;
   ExtMode := 0;
  end;

  F.Read(dt,8);

  F.Read(Cnt,4);

  for I := 0 to Cnt - 1 do begin
   Add(THTMLDoc.Create);
   Application.ProcessMessages;
   LoadHTMLDoc(THTMLDoc(Items[Count - 1]),F,ver);
  end;

{  M := TMemoryStream.Create;
  F.Read(len,4);
  GetMem(zb,len);
  F.Read(zb^,len);
  M.Write(zb^,len);
  FreeMem(zb);
  M.Free;}

  ComboBox1.ItemIndex := 0;
  FirstGen := true;

  ViewTree;

  FirstGen := false;

  Screen.Cursor := crDefault;

  Result := True;
 end;
 except
  DocsBook.Delete(DocsBook.CurBook);
  Dec(DocsBook.CurBook);
  Screen.Cursor := crDefault;
  Result := False;
 end;
end;

procedure TMainForm.Nav(HTMLDoc: THTMLDoc);
begin
  try
      ResPageNavigate := false;

      sb.Panels[0].Text := HTMLDoc.Nam;
      DeployPictures(HTMLDoc);
      HTMLDoc.HTMLText.Seek(0, 0);
      Application.ProcessMessages;
      try
       HTMLDoc.HTMLText.SaveToFile(GetTmpDir+'tmp.html');
      except
      end;
      wb2.Navigate(GetTmpDir+'tmp.html');

      with wb2 do
      if Document <> nil then
        with Application as IOleobject do
         DoVerb(OLEIVERB_UIACTIVATE, nil, wb2, 0, Handle, GetClientRect);
  except
  end;
end;

procedure TMainForm.Nav2(HTMLDoc: THTMLDoc);
begin
    try
      ResPageNavigate := false;

      sb.Panels[0].Text := HTMLDoc.Nam;
      DeployPictures(HTMLDoc);
      HTMLDoc.HTMLText.Seek(0, 0);
      Application.ProcessMessages;
      try
       HTMLDoc.HTMLText.SaveToFile(GetTmpDir+'tmp.html');
      except
      end;
      wb2.Navigate(GetTmpDir+'tmp.html');
    except
    end;  
end;

procedure TMainForm.GenerateTestPage(HTMLDoc : THTMLDoc);
var
 sout : TStringList;
 J, I: Integer;

procedure AddTests(PQ, SelHTML, ParentHTML: THTMLDoc; var sl : TStringList);
var
  J, I: Integer;
begin
  if PQ.HTMLType = 2 then
  begin
   if ParentHTML = SelHTML then
    sout.Add('<li><P><font face="Tahoma,Arial" size="2"><a href="mtest://'+PQ.Nam+'">'+PQ.Nam+'</a></FONT></P></li>')
   else
   if ItemsTree.Selected.AbsoluteIndex = 0 then
    sout.Add('<li><P><font face="Tahoma,Arial" size="2"><a href="mtest://'+PQ.Nam+'">'+PQ.Nam+'</a></FONT></P></li>');
  end;
  if PQ.HTMLType <> 2 then
   for I := 0 to PQ.Children.Count - 1 do
    AddTests(THTMLDoc(PQ.Children.Items[I]), SelHTML, PQ, sl);
end;

begin
 sout := TStringList.Create;
 sout.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">');
 sout.Add('<HTML>');
 sout.Add('<HEAD>');
 sout.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 sout.Add('<meta http-equiv="Content-Language" content="ru">');
 sout.Add('<TITLE>Тестовые задания</TITLE>');
 sout.Add('</HEAD>');

 sout.Add('<BODY BGCOLOR="#FFFFFF" TEXT="#000000" leftmargin="0" marginheight="0" marginwidth="0" topmargin="0">');

 sout.Add('<table width="100%" border="0" cellspacing="0" cellpadding="10">');
 sout.Add('<td>');

 if HTMLDoc = nil then
  sout.Add('<P><font face="Tahoma,Arial" size="2"><B>Все тестовые задания</B></FONT></P>')
 else
  sout.Add('<P><font face="Tahoma,Arial" size="2"><B>Тестовые задания раздела (темы) "'+HTMLDoc.Nam+'":</B></FONT></P>');

 sout.Add('<OL>');
 with DocsBook.ActiveBook do
   begin
      for I := 0 to Count - 1 do begin
        if HTML[I].HTMLType = 2 then
         if ItemsTree.Selected.AbsoluteIndex = 0 then
          sout.Add('<li><P><font face="Tahoma,Arial" size="2"><a href="mtest://'+HTML[I].Nam+'">'+HTML[I].Nam+'</a></FONT></P></li>');
        if HTML[I].HTMLType <> 2 then
         for J := 0 to HTML[I].Children.Count-1 do
          AddTests(THTMLDoc(HTML[I].Children.Items[J]), HTMLDoc, HTML[I], sout);
      end;
   end;

 sout.Add('</OL>');
 sout.Add('</td>');
 sout.Add('</table>');

 sout.Add('</BODY>');
 sout.Add('</HTML>');

 sout.SaveToFile(GetTmpDir+'seek.html');
 sout.Free;

 if FileExists(GetTmpDir+'seek.html') then
  wb2.Navigate(GetTmpDir+'seek.html');
end;

procedure TMainForm.GenerateFilePage(HTMLDoc : THTMLDoc);
var
 sout : TStringList;
 J, I: Integer;
 s : string;

procedure AddFiles(PQ, SelHTML, ParentHTML: THTMLDoc; var sl : TStringList);
var
 J, I: Integer;
 s : string;
begin
  if PQ.HTMLType = 5 then
  begin

   if trunc(PQ.AnyFiles.Size/1048575)>0 then
    s := '('+IntToStr(trunc(PQ.AnyFiles.Size/1048575))+' Мбайт)'
   else
   if trunc(PQ.AnyFiles.Size/1024)>0 then
    s := '('+IntToStr(trunc(PQ.AnyFiles.Size/1024))+' Кбайт)'
   else
    s := '('+IntToStr(PQ.AnyFiles.Size)+' байт)';

   if ParentHTML = SelHTML then
    sout.Add('<li><P><font face="Tahoma,Arial" size="2"><a href="mtest://'+PQ.Nam+'">'+PQ.Nam+'</a> '+s+'</FONT></P></li>')
   else
   if ItemsTree.Selected.AbsoluteIndex = 0 then
    sout.Add('<li><P><font face="Tahoma,Arial" size="2"><a href="mtest://'+PQ.Nam+'">'+PQ.Nam+'</a> '+s+'</FONT></P></li>');
  end;
  if PQ.HTMLType <> 5 then
   for I := 0 to PQ.Children.Count - 1 do
    AddFiles(THTMLDoc(PQ.Children.Items[I]), SelHTML, PQ, sl);
end;

begin
 sout := TStringList.Create;
 sout.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">');
 sout.Add('<HTML>');
 sout.Add('<HEAD>');
 sout.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 sout.Add('<meta http-equiv="Content-Language" content="ru">');
 sout.Add('<TITLE>Тестовые задания</TITLE>');
 sout.Add('</HEAD>');

 sout.Add('<BODY BGCOLOR="#FFFFFF" TEXT="#000000" leftmargin="0" marginheight="0" marginwidth="0" topmargin="0">');

 sout.Add('<table width="100%" border="0" cellspacing="0" cellpadding="10">');
 sout.Add('<td>');

 if HTMLDoc = nil then
  sout.Add('<P><font face="Tahoma,Arial" size="2"><B>Все файлы</B></FONT></P>')
 else
  sout.Add('<P><font face="Tahoma,Arial" size="2"><B>Файлы раздела (темы) "'+HTMLDoc.Nam+'":</B></FONT></P>');

 sout.Add('<OL>');
 with DocsBook.ActiveBook do
   begin
      for I := 0 to Count - 1 do begin
        if HTML[I].HTMLType = 5 then
         if ItemsTree.Selected.AbsoluteIndex = 0 then
         begin
          if trunc(HTML[I].AnyFiles.Size/1048575)>0 then
           s := '('+IntToStr(trunc(HTML[I].AnyFiles.Size/1048575))+' Мбайт)'
          else
          if trunc(HTML[I].AnyFiles.Size/1024)>0 then
           s := '('+IntToStr(trunc(HTML[I].AnyFiles.Size/1024))+' Кбайт)'
          else
           s := '('+IntToStr(HTML[I].AnyFiles.Size)+' байт)';
          sout.Add('<li><P><font face="Tahoma,Arial" size="2"><a href="mtest://'+HTML[I].Nam+'">'+HTML[I].Nam+'</a> '+s+'</FONT></P></li>');
         end;
        if HTML[I].HTMLType <> 5 then
         for J := 0 to HTML[I].Children.Count-1 do
          AddFiles(THTMLDoc(HTML[I].Children.Items[J]), HTMLDoc, HTML[I], sout);
      end;
   end;

 sout.Add('</OL>');
 sout.Add('</td>');
 sout.Add('</table>');

 sout.Add('</BODY>');
 sout.Add('</HTML>');

 sout.SaveToFile(GetTmpDir+'seek.html');
 sout.Free;

 if FileExists(GetTmpDir+'seek.html') then
  wb2.Navigate(GetTmpDir+'seek.html');
end;

procedure TMainForm.Navigate(navmode : boolean);
var
 HTMLDoc : THTMLDoc;
begin
 ResPageNavigate := false;

 if not URLGo then
 begin
  sb.Panels[0].Text := '';
  HTMLDoc := THTMLDoc(ItemsTree.Selected.Data);
  N14.Enabled := false;
  N10.Enabled := false;
  N16.Enabled := false;

  if ComboBox1.ItemIndex = 0 then
  begin
   TreeSaveToFile;
   ContentIndex := ItemsTree.Selected.AbsoluteIndex;
  end
  else
  if ComboBox1.ItemIndex = 1 then
   TestTreeIndex := ItemsTree.Selected.AbsoluteIndex
  else
  if ComboBox1.ItemIndex = 2 then
   TestTreeIndex := ItemsTree.Selected.AbsoluteIndex;

  if ComboBox1.ItemIndex = 1 then
   GenerateTestPage(HTMLDoc)
  else
  if ComboBox1.ItemIndex = 2 then
   GenerateFilePage(HTMLDoc)
  else
  if HTMLDoc <> nil then
   if HTMLDoc.HTMLType = 1 then
    if HTMLDoc.HTMLText.Size > 0 then
    begin
     N14.Enabled := true;
     N10.Enabled := true;
     if ItemsTree.Selected.ImageIndex = 4 then N16.Enabled := true;
      if navmode then
       Nav(HTMLDoc)
      else
       Nav2(HTMLDoc);
    end;
 end;
end;


procedure TMainForm.ItemsTreeClick(Sender: TObject);
begin
 Navigate(true);
end;

procedure TMainForm.DeployPictures(HTMLDoc : THTMLDoc);
var
  len : integer;
  M : TMemoryStream;
  buffer : PChar;
  rbuffer : PByte;
  fname : string;
begin
 if HTMLDoc.Pictures.Size > 0 then
 begin
  HTMLDoc.Pictures.Position := 0;
  while HTMLDoc.Pictures.Position < HTMLDoc.Pictures.Size do
  try
      HTMLDoc.Pictures.Read(len,4);
      buffer := AllocMem(len+1);
      HTMLDoc.Pictures.Read(buffer^,len);
      fname := StrPas(buffer);
      FreeMem(buffer);
      HTMLDoc.Pictures.Read(len,4);
      GetMem(rbuffer,len);
      HTMLDoc.Pictures.Read(rbuffer^,len);
      M := TMemoryStream.Create;
      M.Write(rbuffer^,len);
      Application.ProcessMessages;
      if fileexists(GetTmpDir+fname) then
       deletefile(GetTmpDir+fname);
      try
       M.SaveToFile(GetTmpDir+fname);
      except
      end;
      TmpList.Add(GetTmpDir+fname);
      M.Free;
      FreeMem(rbuffer);
  except
  end;
 end;
 if HTMLDoc.AnyFiles.Size > 0 then
 begin
  HTMLDoc.AnyFiles.Position := 0;
  while HTMLDoc.AnyFiles.Position < HTMLDoc.AnyFiles.Size do
  try
      HTMLDoc.AnyFiles.Read(len,4);
      buffer := AllocMem(len+1);
      HTMLDoc.AnyFiles.Read(buffer^,len);
      fname := StrPas(buffer);
      FreeMem(buffer);
      HTMLDoc.AnyFiles.Read(len,4);
      GetMem(rbuffer,len);
      HTMLDoc.AnyFiles.Read(rbuffer^,len);
      M := TMemoryStream.Create;
      M.Write(rbuffer^,len);
      Application.ProcessMessages;
      if fileexists(GetTmpDir+fname) then
       deletefile(GetTmpDir+fname);
      try
       M.SaveToFile(GetTmpDir+fname);
      except
      end;
      TmpList.Add(GetTmpDir+fname);
      M.Free;
      FreeMem(rbuffer);
  except
  end;
 end;
end;

procedure TMainForm.wb2DocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  Disp := pDisp;
end;

procedure TMainForm.DeleteBookMark(s: string);
var
 Reg : TRegistry;
begin
  Reg:=TRegistry.Create;
  try
   Reg.RootKey := HKEY_CURRENT_USER;
   if Reg.OpenKey(BookRegStr + Caption,False) then
   begin
    if Reg.ValueExists('BookMark'+s) then
    begin
     Reg.DeleteValue('BookMark'+s);
     Reg.DeleteValue('BookMarkComm'+s);
     CreateBookMarkMenu(Reg);
     GenerateBookMarkHTML(Reg);
    end;
    Reg.CloseKey;
   end;
  finally
   Reg.Free;
  end;
end;

procedure TMainForm.DeleteResult(s: string);
var
 Reg : TRegistry;
 rescnt, reslen, index ,j, len : integer;
 buffer : PByte;
 sbuffer : PChar;
 resmem,  m : TMemoryStream;
begin
  index := StrToInt(s);
  Reg:=TRegistry.Create;
  try
   Reg.RootKey := HKEY_CURRENT_USER;
   if Reg.OpenKey(BookRegStr + Caption,False) then
   begin
    if Reg.ValueExists('ResultCnt') then
    begin
             rescnt := Reg.ReadInteger('ResultCnt');
             reslen := Reg.ReadInteger('ResultLength');
             GetMem(buffer,reslen);
             Reg.ReadBinaryData('ResultBuffer',buffer^,reslen);
             resmem := TMemoryStream.Create;
             resmem.Write(buffer^,reslen);
             FreeMem(buffer);
             resmem.Position := 0;
             M := TMemoryStream.Create;

             for j := 1 to rescnt do
             begin
              resmem.Read(len,4);
              if j<>index then
               m.Write(len,4);
              resmem.Read(len,4);
              if j<>index then
               m.Write(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              if j<>index then
               m.Write(sbuffer^,len);
              freemem(sbuffer);

              resmem.Read(len,4);
              if j<>index then
               m.Write(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              if j<>index then
               m.Write(sbuffer^,len);
              freemem(sbuffer);

              buffer := AllocMem(48);
              resmem.Read(buffer^,48);
              if j<>index then
               m.Write(buffer^,48);
              freemem(buffer);

              resmem.Read(len,4);
              if j<>index then
               m.Write(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              if j<>index then
               m.Write(sbuffer^,len);
              freemem(sbuffer);
             end;

            m.Position := 0;
            GetMem(buffer,m.Size);
            m.Read(buffer^,m.Size);
            Reg.WriteBinaryData('ResultBuffer',buffer^,m.Size);
            Reg.WriteInteger('ResultLength',m.Size);
            Reg.WriteInteger('ResultCnt',rescnt-1);
            FreeMem(buffer);

            resmem.Free;
            m.Free;

    end;
    Reg.CloseKey;
   end;
  finally
   Reg.Free;
  end;
  N19Click(nil);
end;

procedure AddResource(const FileName:string;Data:Pointer;Size:Integer;r:string);
var
  HLib:THandle;
  Discard:Boolean;
begin
  HLib:=BeginUpdateResource(PChar(FileName),False);
  Win32Check(HLib<>0);
  Discard:=True;
  try
    Win32Check(UpdateResource(HLib,RT_RCDATA,PChar(r),0,Data,Size));
    Discard:=False;
  finally
    if not EndUpdateResource(HLib,Discard)and not Discard then
      ShowMessage('Ошибка при создании исполняемого файла.');
  end;
end;

procedure TMainForm.SaveResults;
var
 sout : TStringList;
 s, c, s2 : string;
 reslen, len, j, rescnt : integer;
 resmem : TMemoryStream;
 Reg : TRegistry;
 sbuffer : PChar;
 psent : PByte;
 dt1, dt2 : TDateTime;
 maxball, userball : real;
 allq : integer;
 userq, userq2 : word;
begin

  s := 'Мои результаты';
  if OpenSaveFileDialog(Handle, 'html', 'HTML файл|*.html', GetCurrentDir,
    'Сохранение результатов тестирования...', s, False) then
  begin
   if FileExists(s) then
   begin
     if MessageBOX(Handle, PChar('База данных результатов тестирования с именем ' + ExtractFileName(s) +
     ' существует. Перезаписать?'), PCHAR(MainTitle), MB_YesNo or MB_ICONQUESTION) =IDNo then
      Exit;
   end;

 sout := TStringList.Create;
 sout.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">');
 sout.Add('<HTML>');
 sout.Add('<HEAD>');
 sout.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 sout.Add('<meta http-equiv="Content-Language" content="ru">');
 sout.Add('<TITLE>Результаты выполенения тестов</TITLE>');
 sout.Add('</HEAD>');

 sout.Add('<BODY BGCOLOR="#FFFFFF" TEXT="#000000" leftmargin="0" marginheight="0" marginwidth="0" topmargin="0">');

 sout.Add('<table width="100%" border="0" cellspacing="0" cellpadding="10">');
 sout.Add('<td>');

 sout.Add('<center>');
 sout.Add('<P><font face="Tahoma,Arial" size="2">Результаты выполнения тестов</FONT></P>');

 if DocsBook.ActiveBook.ExtMode >=1 then
 begin
  c := 'bgcolor="#'+tablecolor+'"'
 end
 else
 if DocsBook.ActiveBook.ExtMode = 0 then
 begin
  c := 'bgcolor="#FFFFFF"';
 end;

        Reg := TRegistry.Create;
        try
         Reg.RootKey := HKEY_CURRENT_USER;
         if Reg.OpenKey(BookRegStr + Caption, False) then
          begin
           if Reg.ValueExists('ResultCnt') then
            if Reg.ReadInteger('ResultCnt') > 0 then
            begin
             rescnt := Reg.ReadInteger('ResultCnt');
             reslen := Reg.ReadInteger('ResultLength');
             GetMem(psent,reslen);
             Reg.ReadBinaryData('ResultBuffer',psent^,reslen);
             resmem := TMemoryStream.Create;
             resmem.Write(psent^,reslen);
             FreeMem(psent);
             resmem.Position := 0;

             sout.Add('<table '+c+' width="100%" border="1" cellspacing="0" cellpadding="0">');
             sout.Add('<tr><td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">№</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Фамилия Имя Отчество</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Наименование теста</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Дата</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Всего вопросов</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Отвечено правильно</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Всего баллов</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Набрано баллов</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Процент</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Оценка</FONT></B></td>');
             sout.Add('</tr>');

             for j := 1 to rescnt do
             begin
              sout.Add('<tr><td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+IntToStr(j)+'</FONT></P></td>');

              resmem.Read(len,4);
              resmem.Read(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              s2:= StrPas(sbuffer);
              sout.Add('<td><P><font face="Tahoma,Arial" size="2">'+s2+'</FONT></P></td>');
              freemem(sbuffer);

              resmem.Read(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              s2:= StrPas(sbuffer);
              sout.Add('<td><P><font face="Tahoma,Arial" size="2">'+s2+'</FONT></P></td>');
              freemem(sbuffer);

              resmem.Read(dt1,8);
              resmem.Read(dt2,8);
              resmem.Read(userball,8);
              resmem.Read(maxball,8);
              resmem.Read(len,4);
              resmem.Read(allq,4);
              resmem.Read(userq,2);
              resmem.Read(len,4);
              resmem.Read(userq2,2);

              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+DateToStr(dt1)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+IntToStr(allq)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+IntToStr(len)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+FormatFloat('0.00',maxball)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+FormatFloat('0.00',userball)+'</FONT></P></td>');

              if maxball<>0 then
               sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+FormatFloat('0.00',userball/maxball*100)+'</FONT></P></td>')
              else
               sout.Add('<td><P><font face="Tahoma,Arial" size="2"></FONT></P></td>');

              resmem.Read(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              s2:= StrPas(sbuffer);
              sout.Add('<td><P><font face="Tahoma,Arial" size="2">'+s2+'</FONT></P></td>');
              freemem(sbuffer);
              sout.Add('</tr>');
             end;

            resmem.Free;
           end;
          Reg.CloseKey;
         end;
        finally
         Reg.Free;
        end;

        sout.Add('</table>');


    sout.Add('</td>');
    sout.Add('</table>');

    sout.Add('<hr align="center" width="100%" size="1" noshade>');
    sout.Add('<p align="center"><font size="-2" face="Tahoma, Arial, Times New Roman">Copyright');
    sout.Add('  &copy; 2005-2011 <a href="http://siberia-soft.ru" target="_blank">Siberia-Soft</a></font></p>');
    sout.Add('</BODY>');
    sout.Add('</HTML>');

    sout.SaveToFile(s);
    sout.Free;
   end;
end;

procedure TMainForm.AddConfig(fname : string);
var
 s, s2 : string;
 HTMLDoc : THTMLDoc;
 reslen, rescnt, j, len, i, jj : integer;
 ex : boolean;
 F : TFileStream;
 psent, buffer : PByte;
 sbuffer : PChar;
 resmem, M, M2 : TMemoryStream;
 Reg : TRegistry;
 strs : TStringList;
begin
        M := TMemoryStream.Create;
        len := Length(Caption);
        M.Write(len,4);
        M.Write(PChar(Caption)^,len);
        AddResource(fname,M.Memory,M.Size,'BOOKNAME');
        M.Free;

        // Конфигурация - только фио
        sbuffer := AllocMem(29);
        AddResource(fname,PChar(sbuffer),29,'BOOKCFG');
        FreeMem(sbuffer);

        // Сформируем users25
        M := TMemoryStream.Create;
        M2 := TMemoryStream.Create;
        Len := 2;
        M2.Write(len,4);
        Reg := TRegistry.Create;
        try
         Reg.RootKey := HKEY_CURRENT_USER;
         if Reg.OpenKey(BookRegStr + Caption, False) then
          begin
           if Reg.ValueExists('ResultCnt') then
            begin
             rescnt := Reg.ReadInteger('ResultCnt');
             reslen := Reg.ReadInteger('ResultLength');
             GetMem(psent,reslen);
             Reg.ReadBinaryData('ResultBuffer',psent^,reslen);
             resmem := TMemoryStream.Create;
             resmem.Write(psent^,reslen);
             FreeMem(psent);
             resmem.Position := 0;
             Strs := TStringList.Create;

             for j := 1 to rescnt do
             begin
              resmem.Read(len,4);
              resmem.Read(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              s2:= StrPas(sbuffer);
              freemem(sbuffer);
              resmem.Read(len,4);
              resmem.Seek(len,soFromCurrent);
              resmem.Seek(48,soFromCurrent);
              resmem.Read(len,4);
              resmem.Seek(len,soFromCurrent);
              if not strs.Find(s2,jj) then
               strs.Add(s2);
             end;

             for j := 0 to Strs.Count-1 do
             begin
              M.Clear;
              len := length(strs[j])+1;
              sbuffer := AllocMem(len);
              CryptStr(strs[j],sbuffer);
              M.Write(len,4);
              M.Write(sbuffer^,len);
              freemem(sbuffer);

              Len := 0;
              M.Write(len,4);
              M.Write(len,4);
              M.Write(len,4);

              s := '';
              len := length(s)+1;
              sbuffer := AllocMem(len);
              CryptStr(s,sbuffer);
              M.Write(len,4);
              M.Write(sbuffer^,len);
              freemem(sbuffer);
              s := '';
              len := length(s)+1;
              sbuffer := AllocMem(len);
              CryptStr(s,sbuffer);
              M.Write(len,4);
              M.Write(sbuffer^,len);
              freemem(sbuffer);

              M.Write(len,4);

              M.Position := 0;
              Len := M.Size;
              M2.Write(len,4);
              buffer := AllocMem(len);
              M.Read(buffer^,len);
              M2.Write(buffer^,len);
              FreeMem(buffer);
             end;

            strs.Free;
            resmem.Free;
           end;
          Reg.CloseKey;
         end;
        finally
         Reg.Free;
        end;
        if M2.Size>0 then
         AddResource(fname,M2.Memory,M2.Size,'BOOKUSER');
        M.Free;
        M2.Free;
end;

procedure TMainForm.wb2BeforeNavigate2(Sender: TObject;
  const pDisp: IDispatch; var URL, Flags, TargetFrameName, PostData,
  Headers: OleVariant; var Cancel: WordBool);
var
 s,s2 : string;
 HTMLDoc : THTMLDoc;
 i : integer;
 ex : boolean;
begin
 ResPageNavigate := false;
 s := URL;
 if copy(URL,1,5) = 'mtest' then
 begin
  Cancel := true;
  ex := false;
  s := copy(s,9,length(s)-9);
  s2 := '';
  if Pos('#',s) <> 0 then
  begin
   s2 := copy(URL, Pos('#',URL), Length(URL) - Pos('#',URL) + 1);
   s := copy(s, 1, Pos('#',s) - 2);
  end;
  s := UrlDecode(s);
  for i:=0 to ItemsTree.Items.Count-1 do
  begin
   HTMLDoc := THTMLDoc(ItemsTree.Items[i].Data);
   if HTMLDoc <> nil then
    if HTMLDoc.HTMLType = 1 then
     if CompStr(HTMLDoc.Nam,s) = 0 then
      begin
       if HTMLDoc.HTMLText.Size > 0 then
       begin
        URLGo := True;
        ItemsTree.Selected := ItemsTree.Items[i];
        URLGo := False;
        sb.Panels[0].Text := HTMLDoc.Nam;
        HTMLDoc.HTMLText.Seek(0, 0);
        HTMLDoc.HTMLText.SaveToFile(GetTmpDir+HTMLDoc.Nam+'.htm');
        TmpList.Add(GetTmpDir+HTMLDoc.Nam+'.htm');
        DeployPictures(HTMLDoc);
       end;
       ex := true;
       if FileExists(GetTmpDir+HTMLDoc.Nam+'.htm') then
        if Length(s2)>0 then
         wb2.Navigate(GetTmpDir+HTMLDoc.Nam+'.htm'+s2)
        else
         wb2.Navigate(GetTmpDir+HTMLDoc.Nam+'.htm');
       break;
      end;
  end;

  if not ex then
  begin

  for i:=0 to TestList.Count-1 do
  begin
   HTMLDoc := THTMLDoc(TestList.Objects[i]);
   if HTMLDoc.HTMLType = 2 then
    if HTMLDoc <> nil then
     if CompStr(HTMLDoc.Nam,s) = 0 then
      begin
       if HTMLDoc.AnyFiles.Size > 0 then
       try
        AddResource(GetTmpDir+'prtest.exe',HTMLDoc.AnyFiles.Memory,HTMLDoc.AnyFiles.Size,'TEST1');
        AddConfig(GetTmpDir+'prtest.exe');
        ShellExecute(0, Nil, PChar(GetTmpDir+'prtest.exe'), nil, PChar(GetTmpDir), SW_NORMAL);
        break;
       except
       end;
      end;
  end;

  for i:=0 to FileList.Count-1 do
   begin
     HTMLDoc := THTMLDoc(FileList.Objects[i]);
     if HTMLDoc.HTMLType = 5 then
      if CompStr(HTMLDoc.Nam,s) = 0 then
      begin
       if HTMLDoc.AnyFiles.Size > 0 then
       try
        HTMLDoc.AnyFiles.SaveToFile(GetTmpDir + 'tmp' + HTMLDoc.Ext);
        TmpList.Add(GetTmpDir + 'tmp' + HTMLDoc.Ext);
        if ShellExecute(0, Nil, PChar(GetTmpDir + 'tmp' + HTMLDoc.Ext), nil, PChar(GetTmpDir), SW_NORMAL) <= 32 then
         MessageBOX(Handle, PChar('Не удалось открыть прикрепленный файл. Возможно, в операционной системе данному типу файлов не сопоставлено ни одного приложения.'), PChar(MainTitle) ,
         MB_ICONERROR);
        break;
       except
       end;
      end;
   end;

 end;

 end
 else if copy(URL,1,5) = 'rbook' then
 begin
  Cancel := true;
  s := copy(s,9,length(s)-9);
  s := UrlDecode(s);
  DeleteBookMark(s);
 end
 else if copy(URL,1,5) = 'rrest' then
 begin
  if MessageBOX(Application.Handle,PChar('Вы действительно хотите удалить результат тестирования?'),PCHAR(MainTitle), mb_YesNo or MB_ICONQUESTION)=IDYes then
  begin
   Cancel := true;
   s := copy(s,9,length(s)-9);
   s := UrlDecode(s);
   DeleteResult(s);
  end
  else
   Cancel := true;
 end
 else if copy(URL,1,5) = 'rsave' then
 begin
  SaveResults;
  Cancel := true;
 end;
end;

procedure TMainForm.FormShow(Sender: TObject);
var
 rs : TResourceStream;
begin
 URLGo := False;
 PanelSeek.Visible := false;
 ToolButton2.Down := true;
 ComboBox1.Visible := TestList.Count > 0;
 ComboBox1.Visible := FileList.Count > 0;
 if TestList.Count = 0 then
 begin
  N19.Enabled  := false;
  ToolButton4.Enabled  := false;
  ImageButton24.Enabled  := false;
 end;
 if ComboBox1.Visible then
  ComboBox1.ItemIndex := 0;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
 rs : TResourceStream;
 F1 : TFileStream;
 resourcebreak : boolean;
 sl : TStringList;
begin
 if DebuggerPresent then Application.Terminate;

 try
   rs := TResourceStream.Create(hinstance, 'BOOK', RT_RCDATA);
   if rs.Size = 0 then
    Application.Terminate;
   rs.Free;
 except
 end;

 BitMapLogo := TBitMap.Create;
 BitMapLogo.LoadFromResourceName(HInstance,'LOGOM');
 bitmapLogo.Transparent := True;
 bitmapLogo.TransParentColor := bitmapLogo.canvas.pixels[1, 1];

 TmpList := TStringList.Create;
 TestList := TStringList.Create;
 FileList := TStringList.Create;

 GetTempPath(SizeOf(cBuffer) - 1, cBuffer);
 tmplist.Add(GetTmpDir+'tree.cfg');

 try
  rs := TResourceStream.Create(hinstance, 'PERS', RT_RCDATA);
  rs.SaveToFile(GetTmpDir+'pers.dll');
  tmplist.Add(GetTmpDir+'pers.dll');
  rs.Free;
 finally
 end;

 try
  rs := TResourceStream.Create(hinstance, 'LIBBZ2', RT_RCDATA);
  rs.SaveToFile(GetTmpDir+'libbz2.dll');
  tmplist.Add(GetTmpDir+'libbz2.dll');
  rs.Free;
 finally
 end;

 try
  rs := TResourceStream.Create(hinstance, 'PRTEST', RT_RCDATA);
  rs.SaveToFile(GetTmpDir+'prtest.exe');
  tmplist.Add(GetTmpDir+'prtest.exe');
  rs.Free;
 finally
 end;

 try
  F1 := TFileStream.Create(GetTmpDir+'prtest.dt3',fmCreate);
  F1.Free;
  tmplist.Add(GetTmpDir+'prtest.dt3');
 except
 end;

 try
   rs := TResourceStream.Create(hinstance, 'BOOK', RT_RCDATA);
   if rs.Size > 0 then
    if not OpenBFile(rs) then
     Exit;
 finally
   rs.Free;
 end;


 if DocsBook.ActiveBook <> nil then
 begin
 if DocsBook.ActiveBook.Locked = 1 then
  begin
   HookID := SetWindowsHookEx(WH_MOUSE, MouseProc, 0, GetCurrentThreadId());
   NextWindow := SetClipboardViewer(Handle);
  end;

 resourcebreak := false;
  
 if DocsBook.ActiveBook.ExtMode >= 1 then
 begin
  try
   rs := TResourceStream.Create(hinstance, 'BTN1', RT_RCDATA);
   if rs.Size = 0 then
    resourcebreak := true;
   rs.Free;

   if not resourcebreak then
   begin
   rs := TResourceStream.Create(hinstance, 'BTN1', RT_RCDATA);
   rs.Position := 0;
   ImageButton17.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN1DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton17.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN1OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton17.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton17.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN1DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton17.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BTN2', RT_RCDATA);
   rs.Position := 0;
   ImageButton18.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN2DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton18.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN2OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton18.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton18.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN2DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton18.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BTN3', RT_RCDATA);
   rs.Position := 0;
   ImageButton19.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN3DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton19.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN3OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton19.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton19.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN3DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton19.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BTN4', RT_RCDATA);
   rs.Position := 0;
   ImageButton20.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN4DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton20.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN4OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton20.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton20.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN4DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton20.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BTN5', RT_RCDATA);
   rs.Position := 0;
   ImageButton21.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN5DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton21.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN5OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton21.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton21.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN5DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton21.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BTN6', RT_RCDATA);
   rs.Position := 0;
   ImageButton22.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN6DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton22.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN6OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton22.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton22.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN6DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton22.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BTN7', RT_RCDATA);
   rs.Position := 0;
   ImageButton23.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN7DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton23.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN7OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton23.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton23.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN7DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton23.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BTN8', RT_RCDATA);
   rs.Position := 0;
   ImageButton24.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN8DIS', RT_RCDATA);
   rs.Position := 0;
   ImageButton24.BitmapDisabled.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN8OVER', RT_RCDATA);
   rs.Position := 0;
   ImageButton24.BitmapOver.LoadFromStream(rs);
   rs.Position := 0;
   ImageButton24.BitmapUp.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'BTN8DOWN', RT_RCDATA);
   rs.Position := 0;
   ImageButton24.BitmapDown.LoadFromStream(rs);
   rs.Free;

   rs := TResourceStream.Create(hinstance, 'BGT', RT_RCDATA);
   rs.Position := 0;
   Image5.Picture.Bitmap.LoadFromStream(rs);
   rs.Free;
   rs := TResourceStream.Create(hinstance, 'LOGOT', RT_RCDATA);
   rs.Position := 0;
   Image6.Picture.Bitmap.LoadFromStream(rs);
   rs.Free;

   sl := TStringList.Create;
   rs := TResourceStream.Create(hinstance, 'COLORS', RT_RCDATA);
   rs.Position := 0;
   sl.LoadFromStream(rs);

   Panel2.Color1 := StrToInt(sl[0]);
   Panel3.Color1 := StrToInt(sl[0]);
   ItemsTree.Color := StrToInt(sl[1]);
   PanelSeek.Color := StrToInt(sl[1]);
   Splitter1.Color := StrToInt(sl[2]);
   ComboBox1.Color := StrToInt(sl[3]);
   TableColor := sl[4];
   sl.Free;

   ImageButton17.Repaint;
   ImageButton18.Repaint;
   ImageButton19.Repaint;
   ImageButton20.Repaint;
   ImageButton21.Repaint;
   ImageButton22.Repaint;
   ImageButton23.Repaint;
   ImageButton24.Repaint;
   Image5.Repaint;
   end;
  except
   resourcebreak := true;
  end;

  if not resourcebreak then
  begin
   Panel6.Visible := true;
   CoolBar1.Visible := false;
  end
  else
  begin
   Panel6.Visible := false;
   CoolBar1.Visible := true;
   Panel2.Color1 := clBtnFace;
   Panel3.Color1 := clBtnFace;
   Splitter1.Color := clBtnFace;
   ItemsTree.Color := clWindow;
   PanelSeek.Color := clWindow;
   ComboBox1.Color := clWindow;
  end;
 end
 else
 if DocsBook.ActiveBook.ExtMode = 0 then
 begin
  Panel6.Visible := false;
  CoolBar1.Visible := true;
  Panel2.Color1 := clBtnFace;
  Panel3.Color1 := clBtnFace;
  Splitter1.Color := clBtnFace;
  ItemsTree.Color := clWindow;
  PanelSeek.Color := clWindow;
  ComboBox1.Color := clWindow;
 end
 else
 begin
   Panel6.Visible := false;
   CoolBar1.Visible := true;
   Panel2.Color1 := clBtnFace;
   Panel3.Color1 := clBtnFace;
   Splitter1.Color := clBtnFace;
   ItemsTree.Color := clWindow;
   PanelSeek.Color := clWindow;
   ComboBox1.Color := clWindow;
 end;
 end;

 if resourcebreak then
  DocsBook.ActiveBook.ExtMode := 0;

 tmplist.Add(GetTmpDir+'spr.cf2');
 tmplist.Add(GetTmpDir+'users21.dat');

end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
var
 i : integer;
begin

 if DocsBook.ActiveBook<> nil then
  DocsBook.ActiveBook.DeleteAllDocs;

 for i:=0 to TmpList.Count-1 do
  DeleteFile(TmpList.Strings[i]);

 if fileexists(GetTmpDir+'link.html') then
  DeleteFile(GetTmpDir+'link.html');
 if fileexists(GetTmpDir+'seek.html') then
  DeleteFile(GetTmpDir+'seek.html');
 if fileexists(GetTmpDir+'tmp.html') then
  DeleteFile(GetTmpDir+'tmp.html');

 TmpList.Free;
 TestList.Free;
 FileList.Free;
 BitMapLogo.Free;
end;

procedure TMainForm.ItemsTreeCollapsing(Sender: TObject; Node: TTreeNode;
  var AllowCollapse: Boolean);
begin
 if Node.ImageIndex = 1 then
  begin
   Node.ImageIndex := 0;
   Node.SelectedIndex := 0;
  end;
end;

procedure TMainForm.ItemsTreeExpanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
begin
 if Node.ImageIndex = 0 then
  begin
   Node.ImageIndex := 1;
   Node.SelectedIndex := 1;
  end;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
if DocsBook.ActiveBook <> nil then
 if DocsBook.ActiveBook.Locked = 1 then
 begin
  if HookID <> 0 then
    UnHookWindowsHookEx(HookID);
  ChangeClipboardChain(Handle, NextWindow);
 end; 
end;

procedure TMainForm.wb2CommandStateChange(Sender: TObject;
  Command: Integer; Enable: WordBool);
begin
case Command of
     CSC_NAVIGATEBACK: begin
      ToolButton1.Enabled := Enable;
      ImageButton17.Enabled := Enable;
      end;
     CSC_NAVIGATEFORWARD: begin
      ToolButton3.Enabled := Enable;
      ImageButton19.Enabled := Enable;
      end;
   end;
end;

procedure TMainForm.ToolButton1Click(Sender: TObject);
begin
 wb2.GoBack;
end;

procedure TMainForm.ToolButton3Click(Sender: TObject);
begin
 wb2.GoForward;
end;

procedure TMainForm.ToolButton5Click(Sender: TObject);
var
 HTMLDoc : THTMLDoc;
begin
   ItemsTree.Selected := ItemsTree.Items[HomeIndex];
   HTMLDoc := THTMLDoc(ItemsTree.Items[HomeIndex].Data);
   if HTMLDoc<> nil then
    if HTMLDoc.HTMLText.Size > 0 then
     Nav(HTMLDoc);
end;

procedure TMainForm.N6Click(Sender: TObject);
var
 sout : TStringList;
 rs : TResourceStream;
begin
  try
    rs := TResourceStream.Create(hinstance, 'LOGOMOVIE', RT_RCDATA);
    rs.SaveToFile(GetTmpDir+'logomovie.gif');
   rs.Free;
  finally
  end;

  TmpList.Add(GetTmpDir+'logomovie.gif');

 with DocsBook.ActiveBook do
 begin

 sout := TStringList.Create;
 sout.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">');
 sout.Add('<HTML>');
 sout.Add('<HEAD>');
 sout.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 sout.Add('<meta http-equiv="Content-Language" content="ru">');
 sout.Add('<TITLE>О Программе</TITLE>');
 sout.Add('</HEAD>');

 sout.Add('<BODY BGCOLOR="#FFFFFF" TEXT="#000000" leftmargin="0" marginheight="0" marginwidth="0" topmargin="0">');

 sout.Add('<table  align="center" border="0" cellspacing="0" cellpadding="0">');
 sout.Add('<tr><td><img SRC="logomovie.gif" border="0"  width="53" height="45"></td></tr>');
 sout.Add('</table><hr align="center" width="100%" size="1" noshade>');

 sout.Add('<table border="0" cellspacing="0" cellpadding="10">');
 sout.Add('<td>');

 sout.Add('<center>');
 sout.Add('<P><A NAME="support"></A><B><U><font face="Tahoma,Arial" size="2">О программе "Электронный учебник"</FONT></U></B></P>');

 sout.Add('<P><font face="Tahoma,Arial" size="2">Автономный модуль "Электронный учебник" является составной частью программного комплекса "Инструментальна среда для создания программно-педагогических тестов и адаптивного тестирования".'+
         ' Дополнительную информацию о программном комплексе можно получить на сайте компании <a href="http://siberia-soft.ru" target="_blank">Siberia-Soft</a> и '+
         'по электронной почте <a href="mailto:siberia-soft@yandex.ru">siberia-soft@yandex.ru</a>.</FONT></P>');

 if Length(Caption)>0 then
  sout.Add('<B><P><font face="Tahoma,Arial" size="2">Наименование и версия учебника:<br>'+Caption+' '+Version+';</FONT></P>');
 if Length(FIO)>0 then
  sout.Add('<P><font face="Tahoma,Arial" size="2">Составители учебника: <a href="mailto:'+Email+'" target="_blank">'+FIO+' ('+ORG+' '+Place+')</a> Телефон: '+Phone+';</FONT></P>');
 if Length(Comment)>0 then
  sout.Add('<P><font face="Tahoma,Arial" size="2">Краткое описание учебника:<br>'+Comment+';</FONT></P>');
 if Length(DateToStr(dt))>0 then
  sout.Add('<P><font face="Tahoma,Arial" size="2">Дата последнего обновления: '+DateToStr(dt)+'.</FONT></P></B></CENTER>');


 sout.Add('</td>');
 sout.Add('</table>');
 sout.Add('<hr align="center" width="100%" size="1" noshade>');
 sout.Add('<p align="center"><font size="-2" face="Tahoma, Arial, Times New Roman">Copyright');
 sout.Add('  &copy; 2005-2011 <a href="http://siberia-soft.ru" target="_blank">Siberia-Soft</a></font></p>');
 sout.Add('</BODY>');
 sout.Add('</HTML>');

 sout.SaveToFile(GetTmpDir+'seek.html');
 sout.Free;

 if FileExists(GetTmpDir+'seek.html') then
  wb2.Navigate(GetTmpDir+'seek.html');

 end;
end;

procedure TMainForm.CloseBtnClick(Sender: TObject);
begin
 Panel1.Visible := false;
 Splitter1.Visible := false;
 N4.Checked := false;
 ToolButton2.Down := N4.Checked;
end;

procedure TMainForm.ItemsTreeContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
var
   tmpNode: TTreeNode;
begin
   tmpNode := (Sender as TTreeView).GetNodeAt(MousePos.X, MousePos.Y);
   if tmpNode <> nil then
   begin
     TTreeView(Sender).Selected := tmpNode;
     ItemsTreeClick(Sender);
   end;
end;

procedure TMainForm.ItemsTreeAdvancedCustomDrawItem(
  Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState;
  Stage: TCustomDrawStage; var PaintImages, DefaultDraw: Boolean);
var
  ARect: TRect;
  S: string;
  HTMLDoc : THTMLDoc;
begin
  HTMLDoc := THTMLDoc(Node.Data);
  case Stage of
    cdPostPaint:
      begin
        ARect := Node.DisplayRect(True);
        ARect.Right := ItemsTree.ClientWidth;
        with ItemsTree.Canvas do
        begin
          if cdsSelected in State then
          begin
            if DocsBook.ActiveBook.ExtMode >= 1 then
             Brush.Color := clWindow
            else
            if DocsBook.ActiveBook.ExtMode = 0 then
             Brush.Color := clBtnFace
          end
          else
          begin
            if DocsBook.ActiveBook.ExtMode >= 1 then
             Brush.Color := Panel3.Color1
            else
            if DocsBook.ActiveBook.ExtMode = 0 then
             Brush.Color := clWindow
          end;
          FillRect(ARect);
          if cdsSelected in State then
          begin
            Brush.Color := clBlack;
            Rectangle(ARect);
          end;
          if HTMLDoc <> nil then
          begin
//           if HTMLDoc.HTMLType = 0 then
//           begin
             Font.Color := clBlack
//           end;
          end
          else
          begin
            Font.Color := clBlack;
            Font.Style := [fsBold];
          end;
          TextOut(ARect.Left+1, ARect.Top+1, Node.Text);
        end;
      end;
  end; { Case }
end;

procedure TMainForm.ToolButtonFindClick(Sender: TObject);
begin
 N20Click(Sender);
end;

procedure TMainForm.Splitter1Moved(Sender: TObject);
var
 Reg : TRegistry;
begin
 CloseBtn.Left := Panel2.Width - 23;
 ComboBox1.Width := Panel2.Width - 28;
 if Panel2.Width > 125 then
  Panel2.Caption := 'Содержание'
 else
  Panel2.Caption := '';
 Reg:=TRegistry.Create;
 try
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey('\Software\CTMTest\EBook\'+Caption,True) then
   Reg.WriteInteger('ContentWidth',Panel1.Width);
  Reg.CloseKey;
 finally
  Reg.Free;
 end;
end;

procedure TMainForm.Splitter2Moved(Sender: TObject);
begin
 CloseBtn2.Left := Panel3.Width - 25;
 Edit1.Width := Panel3.Width - 7;
end;

procedure TMainForm.CloseBtn2Click(Sender: TObject);
begin
 PanelSeek.Visible := false;
 ToolButtonFind.Down := false;
 if N4.Checked then
 begin
  Splitter1.Visible := N4.Checked;
  Panel1.Visible := N4.Checked;
 end;
end;

procedure TMainForm.Button1Click(Sender: TObject);
var
 i,j : integer;
 HTMLDoc, HTMLDoc2 : THTMLDoc;
 s : TStringStream;
 s1, s2 : string;
 sout, s3 : TStringList;
 found : boolean;
 Node : TTreeNode;
begin
 ResPageNavigate := false;

 Screen.Cursor := crHourGlass;
 sout := TStringList.Create;
 sout.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">');
 sout.Add('<html>');
 sout.Add('<head>');
 sout.Add('<title>Результаты поиска</title>');
 sout.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 sout.Add('</head>');
 sout.Add('<body>');
 found := false;
 sout.Add('<font face="Tahoma,Arial" size="2">Результаты поиска ('+ItemsTree.Items[0].Text+'):</font><font face="Tahoma,Arial" size="2"><br><ol>');
 try
 for i:=1 to ItemsTree.Items.Count-1 do
  begin
   HTMLDoc := THTMLDoc(ItemsTree.Items[i].Data);
   if HTMLDoc <> nil then
    if HTMLDoc.HTMLType = 1 then
    begin
      HTMLDoc.HTMLText.Seek(0, 0);
      Application.ProcessMessages;
      try
       s := TStringStream.Create('');
       s.CopyFrom(HTMLDoc.HTMLText,HTMLDoc.HTMLText.Size);
       if Pos(Edit1.Text,s.DataString)<>0 then
       begin
        s1 := copy(s.DataString,Pos(Edit1.Text,s.DataString),200);
        if pos('<',s1)<>0 then s1 := copy(s1,1,pos('<',s1)-1);
        if pos('>',s1)<>0 then s1 := copy(s1,1,pos('>',s1)-1);
        if pos('"',s1)<>0 then s1 := copy(s1,1,pos('"',s1)-1);

        s2 := '';
        s3 := TStringList.Create;
        Node := ItemsTree.Items[i];
        for j := ItemsTree.Items[i].Level downto 1 do
         begin
          Node := Node.Parent;
          HTMLDoc2 := THTMLDoc(Node.Data);
          if HTMLDoc2 <> nil then
           s3.Add(HtmlDoc2.Nam);
         end;

        for j := s3.Count-1 downto 0 do
           s2 := s2 + s3.Strings[j] + '&nbsp;&#8212;&nbsp;';

        s3.Free;
        s1 := '<p><li><b>' + s2 + HTMLDoc.Nam+':</b>&nbsp;&nbsp;<a href="mtest://'+
        HTMLDoc.Nam+'/">'+s1+'</a></p>';
        sout.Add(s1);
        found := true;
       end;
       s.Free;
      except
      end;
    end;
  end;
  if not found then sout.Add('Фраза или слово не найдены.') else sout.Add('</ol>');
 finally
  sout.Add('</font></body>');
  sout.Add('</html>');
  sout.SaveToFile(GetTmpDir+'seek.html');
  if FileExists(GetTmpDir+'seek.html') then
   wb2.Navigate(GetTmpDir+'seek.html');
  sout.Free;
 end;
 Screen.Cursor := crDefault;
end;

procedure TMainForm.GenerateBookMarkHTML(Reg: TRegistry);
var
 i,j : integer;
 s, s1 : string;
 sout : TStringList;
 found : boolean;
 Node : TTreeNode;
 HTMLDoc, HTMLDoc2 : THTMLDoc;
begin
 Screen.Cursor := crHourGlass;
 sout := TStringList.Create;
 sout.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">');
 sout.Add('<html>');
 sout.Add('<head>');
 sout.Add('<title>Закладки учебника</title>');
 sout.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 sout.Add('</head>');
 sout.Add('<body>');
 sout.Add('<font face="Tahoma,Arial" size="2">Закладки учебника "'+ItemsTree.Items[0].Text+'":</font><font face="Tahoma,Arial" size="2"><br><ol>');
 try
   for j:=1 to ItemsTree.Items.Count-1 do
   begin
            HTMLDoc := ItemsTree.Items[j].Data;
            if HTMLDoc <> nil then
             if HTMLDoc.HTMLType = 1 then
               begin
                ItemsTree.Items[j].ImageIndex := 2;
                ItemsTree.Items[j].SelectedIndex := 2;
               end;
   end;
   for i:=1 to 10 do
    if Reg.ValueExists('BookMark'+IntToStr(i)) then
    begin
        s := Reg.ReadString('BookMark'+IntToStr(i));
        for j:=1 to ItemsTree.Items.Count-1 do
         if CompStr(ItemsTree.Items[j].Text,s)=0 then
           begin
            HTMLDoc := ItemsTree.Items[j].Data;
            if HTMLDoc <> nil then
             if HTMLDoc.HTMLType = 1 then
               begin
                ItemsTree.Items[j].ImageIndex := 4;
                ItemsTree.Items[j].SelectedIndex := 4;
               end;
           end;
        s1 := '<p><li><a href="mtest://'+
        s+'/">'+s+'</a><br><font face="Tahoma,Arial" size="1">'+Reg.ReadString('BookMarkComm'+IntToStr(i))+'&nbsp;&nbsp;&nbsp;<a href="rbook://'+IntToStr(i)+'">удалить закладку</a></font></p>';
        sout.Add(s1);
    end;
 finally
  sout.Add('</font></body>');
  sout.Add('</html>');
  sout.SaveToFile(GetTmpDir+'seek.html');
  if FileExists(GetTmpDir+'seek.html') then
   wb2.Navigate(GetTmpDir+'seek.html');
  sout.Free;
 end;
 Screen.Cursor := crDefault;
end;

procedure TMainForm.ToolButton2Click(Sender: TObject);
begin
 N4Click(Sender);
end;

procedure TMainForm.N7Click(Sender: TObject);
begin
  ToolButton5Click(Sender);
end;

procedure TMainForm.N9Click(Sender: TObject);
var
 Reg:TRegistry;
begin
 N9.Checked := not N9.Checked;
 if N9.Checked then BorderStyle := bsNone else BorderStyle := bsSizeable;
 Reg:=TRegistry.Create;
 try
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey('\Software\CTMTest\EBook\'+Caption,True) then
   Reg.WriteBool('FullScreen',N9.Checked);
  Reg.CloseKey;
 finally
  Reg.Free;
 end;
end;

procedure TMainForm.BookPopupHandler(Sender: TObject);
var
  s : string;
  i : integer;
begin
  with Sender as TMenuItem do begin
   s := Caption;
   s := ReplaceStr(s, '&', '');
  end;
   try
     for i:=0 to ItemsTree.Items.Count-1 do
      if CompStr(s,ItemsTree.Items[i].Text)=0 then
       begin
        ItemsTree.Selected := ItemsTree.Items[i];
        ItemsTreeClick(Sender);
        break;
       end;
   except
    MessageBOX(Handle, PChar('Не удалось открыть закладку "'+s+'".'), PChar(MainTitle) ,
    MB_ICONERROR);
   end;
end;

function TMainForm.CreateBookMarkMenu(Reg: TRegistry):boolean;
var
 i : integer;
 NewItem : TMenuItem;
 NewItem2 : TMenuItem;
 s,f : string;
 ex : boolean;
begin
   ReopenBookMenu.Items.Clear;
   N15.Clear;
   ex := false;
   for i:=1 to 10 do
    if Reg.ValueExists('BookMark'+IntToStr(i)) then
    begin
       s := Reg.ReadString('BookMark'+IntToStr(i));
       NewItem := TMenuItem.Create(N15);
       NewItem.Caption := s;
       NewItem.ImageIndex := 5;
       NewItem.OnClick := bookPopupHandler;
       N15.Add(NewItem);
       NewItem2 := TMenuItem.Create(ReopenBookMenu);
       NewItem2.Caption := NewItem.Caption;
       NewItem2.Hint := NewItem.Hint;
       NewItem2.ImageIndex := 5;
       NewItem2.OnClick := bookPopupHandler;
       ReopenBookMenu.Items.Add(NewItem2);
       ex:=true;
    end;
   Result := ex;
end;

procedure TMainForm.N10Click(Sender: TObject);
var
 Reg : TRegistry;
 i : integer;
 ex : boolean;
begin
 NewBookMarkForm.Caption := 'Добавить закладку на страницу ' + ItemsTree.Selected.Text;
 if NewBookMarkForm.ShowModal = mrOk then
 begin
 Reg:=TRegistry.Create;
 try
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey(BookRegStr+Caption,True) then
  begin
   ex := false;
   for i:=1 to 11 do
    if not Reg.ValueExists('BookMark'+IntToStr(i)) then
      break
    else
      if CompStr(Reg.ReadString('BookMark'+IntToStr(i)),ItemsTree.Selected.Text)=0 then
        ex := true;
   if not ex then
   begin
   if i<11 then
   begin
    Reg.WriteString('BookMark'+IntToStr(i),ItemsTree.Selected.Text);
    Reg.WriteString('BookMarkComm'+IntToStr(i),NewBookMarkForm.Memo1.Text);
   end
   else
    begin
    // сдвинем строки вниз
    for i:=9 downto 1 do
    begin
     Reg.WriteString('BookMark'+IntToStr(i+1),Reg.ReadString('BookMark'+IntToStr(i)));
     Reg.WriteString('BookMarkComm'+IntToStr(i+1),Reg.ReadString('BookMarkComm'+IntToStr(i)));
    end;
    // запишем ее первой
    Reg.WriteString('BookMark1',ItemsTree.Selected.Text);
    Reg.WriteString('BookMarkComm1',NewBookMarkForm.Memo1.Text);
    end;
   end;
   ItemsTree.Selected.ImageIndex := 4;
   ItemsTree.Selected.SelectedIndex := 4;
   CreateBookMarkMenu(Reg);
   Reg.CloseKey;
  end;
 finally
  Reg.Free;
 end;
 end;
end;

procedure TMainForm.N14Click(Sender: TObject);
begin
 N10Click(Sender);
end;

procedure TMainForm.ToolButton8Click(Sender: TObject);
begin
 ToolButton8.CheckMenuDropdown;
end;

procedure TMainForm.N11Click(Sender: TObject);
begin
 ShowBookMark;
end;

procedure TMainForm.N16Click(Sender: TObject);
var
 s : string;
 i : integer;
 Reg : TRegistry;
begin
 Reg:=TRegistry.Create;
 try
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey('\Software\CTMTest\EBook\'+Caption,False) then
  begin
   for i:=1 to 10 do
    if Reg.ValueExists('BookMark'+IntToStr(i)) then
    begin
       s := Reg.ReadString('BookMark'+IntToStr(i));
       if CompStr(ItemsTree.Selected.Text,s)=0 then
        begin
         DeleteBookMark(IntToStr(i));
         break;
        end;
    end;
  end;
 finally
  Reg.Free;
 end;
end;

procedure TMainForm.N17Click(Sender: TObject);
var
 rs : TResourceStream;
 bmp : TBitmap;
begin
  ResPageNavigate := false;

  try
    rs := TResourceStream.Create(hinstance, 'LOGOMOVIE', RT_RCDATA);
    rs.SaveToFile(GetTmpDir+'logomovie.gif');
  finally
   rs.Free;
  end;
  TmpList.Add(GetTmpDir+'logomovie.gif');

  try
    bmp := TBitmap.Create;
    if DocsBook.ActiveBook.ExtMode = 0 then
     bmp.LoadFromResourceName(hinstance, 'BACK')
    else
    if DocsBook.ActiveBook.ExtMode >= 1 then
     bmp.Assign(ImageButton17.Bitmap);
    bmp.SaveToFile(GetTmpDir+'back.bmp');
   bmp.Free;
  finally
  end;
  TmpList.Add(GetTmpDir+'back.bmp');

  try
    bmp := TBitmap.Create;
    if DocsBook.ActiveBook.ExtMode = 0 then
     bmp.LoadFromResourceName(hinstance, 'CONTENT')
    else
    if DocsBook.ActiveBook.ExtMode >= 1 then
     bmp.Assign(ImageButton20.Bitmap);
    bmp.SaveToFile(GetTmpDir+'content.bmp');
  finally
   bmp.Free;
  end;
  TmpList.Add(GetTmpDir+'content.bmp');

  try
    bmp := TBitmap.Create;
    if DocsBook.ActiveBook.ExtMode = 0 then
     bmp.LoadFromResourceName(hinstance, 'HOME')
    else
    if DocsBook.ActiveBook.ExtMode >= 1 then
     bmp.Assign(ImageButton18.Bitmap);
    bmp.SaveToFile(GetTmpDir+'home.bmp');
  finally
   bmp.Free;
  end;
  TmpList.Add(GetTmpDir+'home.bmp');

  try
    bmp := TBitmap.Create;
    if DocsBook.ActiveBook.ExtMode = 0 then
     bmp.LoadFromResourceName(hinstance, 'SEEK')
    else
    if DocsBook.ActiveBook.ExtMode >= 1 then
     bmp.Assign(ImageButton21.Bitmap);
    bmp.SaveToFile(GetTmpDir+'seek.bmp');
  finally
   bmp.Free;
  end;
  TmpList.Add(GetTmpDir+'seek.bmp');

  try
    bmp := TBitmap.Create;
    if DocsBook.ActiveBook.ExtMode = 0 then
     bmp.LoadFromResourceName(hinstance, 'BOOKMARK')
    else
    if DocsBook.ActiveBook.ExtMode >= 1 then
     bmp.Assign(ImageButton22.Bitmap);
    bmp.SaveToFile(GetTmpDir+'bookmark.bmp');
  finally
   bmp.Free;
  end;
  TmpList.Add(GetTmpDir+'bookmark.bmp');

  try
    bmp := TBitmap.Create;
    if DocsBook.ActiveBook.ExtMode = 0 then
     bmp.LoadFromResourceName(hinstance, 'TESTS')
    else
    if DocsBook.ActiveBook.ExtMode >= 1 then
     bmp.Assign(ImageButton24.Bitmap);
    bmp.SaveToFile(GetTmpDir+'tests.bmp');
  finally
   bmp.Free;
  end;
  TmpList.Add(GetTmpDir+'tests.bmp');

  try
    rs := TResourceStream.Create(hinstance, 'HELP', RT_RCDATA);
    rs.SaveToFile(GetTmpDir+'seek.html');
  finally
   rs.Free;
  end;
  if FileExists(GetTmpDir+'seek.html') then
   wb2.Navigate(GetTmpDir+'seek.html');
end;

procedure TMainForm.ToolBar1CustomDraw(Sender: TToolBar;
  const ARect: TRect; var DefaultDraw: Boolean);
begin
// ToolBar1.Canvas.BrushCopy(Bounds(ARect.Right-53, 0, 53, 45), BitmapLogo, BitmapLogo.Canvas.ClipRect, clWhite);
 ToolBar1.Canvas.CopyRect(Bounds(ARect.Right-53, 0, 53, 45), BitmapLogo.Canvas, BitmapLogo.Canvas.ClipRect);
end;

procedure TMainForm.ToolBar1MouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
 if (X>=ToolBar1.Canvas.ClipRect.Right-53) and (X<=ToolBar1.Canvas.ClipRect.Right) then
  Screen.Cursor := crHandPoint
 else
  Screen.Cursor := crDefault;
end;

procedure TMainForm.FormResize(Sender: TObject);
begin
 Repaint;
end;

procedure TMainForm.ToolBar1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
 if (X>=ToolBar1.Canvas.ClipRect.Right-53) and (X<=ToolBar1.Canvas.ClipRect.Right) then
  ShellExecute(0, Nil, PChar('http://www.siberia-soft.ru'), nil, nil, SW_NORMAL);
end;

procedure TMainForm.ItemsTreeCustomDraw(Sender: TCustomTreeView;
  const ARect: TRect; var DefaultDraw: Boolean);
begin
{ with ItemsTree.Canvas do
  begin
      Brush.Color := $00EBEDF3;
      FillRect(ARect);
  end; }
end;

procedure TMainForm.ImageButton5Click(Sender: TObject);
begin
 ToolButtonFindClick(Sender);
end;

procedure TMainForm.ImageButton4Click(Sender: TObject);
begin
 N4Click(Sender);
end;

procedure TMainForm.ImageButton2Click(Sender: TObject);
begin
 ToolButton5Click(Sender);
end;

procedure TMainForm.ImageButton3Click(Sender: TObject);
begin
 ToolButton3Click(Sender);
end;

procedure TMainForm.ImageButton1Click(Sender: TObject);
begin
 ToolButton1Click(Sender);
end;

procedure TMainForm.ImageButton7Click(Sender: TObject);
begin
 N17Click(Sender);
end;

procedure TMainForm.ImageButton6Click(Sender: TObject);
begin
 ShowBookMark;
end;

procedure TMainForm.Image2Click(Sender: TObject);
begin
 ShellExecute(0, Nil, PChar('http://www.siberia-soft.ru'), nil, nil, SW_NORMAL);
end;

procedure TMainForm.Image2MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
//  Screen.Cursor := crHandPoint;
end;

procedure TMainForm.Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
//  Screen.Cursor := crDefault;
end;

procedure TMainForm.N19Click(Sender: TObject);
var
 sout : TStringList;
 c, s2 : string;
 reslen, len, j, rescnt : integer;
 resmem : TMemoryStream;
 Reg : TRegistry;
 sbuffer : PChar;
 psent : PByte;
 dt1, dt2 : TDateTime;
 maxball, userball : real;
 allq : integer;
 userq, userq2 : word;
begin
 ResPageNavigate := true;

 sout := TStringList.Create;
 sout.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">');
 sout.Add('<HTML>');
 sout.Add('<HEAD>');
 sout.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 sout.Add('<meta http-equiv="Content-Language" content="ru">');
 sout.Add('<TITLE>Результаты выполенения тестов</TITLE>');
 sout.Add('</HEAD>');

 sout.Add('<BODY BGCOLOR="#FFFFFF" TEXT="#000000" leftmargin="0" marginheight="0" marginwidth="0" topmargin="0">');

 sout.Add('<table width="100%" border="0" cellspacing="0" cellpadding="10">');
 sout.Add('<td>');

 sout.Add('<center>');
 sout.Add('<P><font face="Tahoma,Arial" size="2">Результаты выполнения тестов</FONT></P>');

 if DocsBook.ActiveBook.ExtMode >= 1 then
  c := 'bgcolor="#'+tablecolor+'"'
 else
 if DocsBook.ActiveBook.ExtMode = 0 then
  c := 'bgcolor="#FFFFFF"';

        Reg := TRegistry.Create;
        try
         Reg.RootKey := HKEY_CURRENT_USER;
         if Reg.OpenKey(BookRegStr + Caption, False) then
          begin
           if Reg.ValueExists('ResultCnt') then
            if Reg.ReadInteger('ResultCnt') > 0 then
            begin
             sout.Add('<table '+c+' width="100%" border="1" cellspacing="0" cellpadding="0">');
             sout.Add('<tr><td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">№</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Фамилия Имя Отчество</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Наименование теста</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Дата</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Всего вопросов</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Отвечено правильно</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Всего баллов</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Набрано баллов</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Процент</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"><B><font face="Tahoma,Arial" size="2">Оценка</FONT></B></td>');
             sout.Add('<td align="center" valign="middle"></td>');
             sout.Add('</tr>');
             rescnt := Reg.ReadInteger('ResultCnt');
             reslen := Reg.ReadInteger('ResultLength');
             GetMem(psent,reslen);
             Reg.ReadBinaryData('ResultBuffer',psent^,reslen);
             resmem := TMemoryStream.Create;
             resmem.Write(psent^,reslen);
             FreeMem(psent);
             resmem.Position := 0;

             for j := 1 to rescnt do
             begin
              sout.Add('<tr><td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+IntToStr(j)+'</FONT></P></td>');
              resmem.Read(len,4);
              resmem.Read(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              s2:= StrPas(sbuffer);
              sout.Add('<td><P><font face="Tahoma,Arial" size="2">'+s2+'</FONT></P></td>');
              freemem(sbuffer);

              resmem.Read(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              s2:= StrPas(sbuffer);
              sout.Add('<td><P><font face="Tahoma,Arial" size="2"><a href="mtest://'+s2+'">'+s2+'</a></FONT></P></td>');
              freemem(sbuffer);

              resmem.Read(dt1,8);
              resmem.Read(dt2,8);
              resmem.Read(userball,8);
              resmem.Read(maxball,8);
              resmem.Read(len,4);
              resmem.Read(allq,4);
              resmem.Read(userq,2);
              resmem.Read(len,4);
              resmem.Read(userq2,2);

              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+DateToStr(dt1)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+IntToStr(allq)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+IntToStr(len)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+FormatFloat('0.00',maxball)+'</FONT></P></td>');
              sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+FormatFloat('0.00',userball)+'</FONT></P></td>');
              if maxball<>0 then
               sout.Add('<td align="center" valign="middle"><P><font face="Tahoma,Arial" size="2">'+FormatFloat('0.00',userball/maxball*100)+'</FONT></P></td>')
              else
               sout.Add('<td><P><font face="Tahoma,Arial" size="2"></FONT></P></td>');

              resmem.Read(len,4);
              sbuffer := AllocMem(len+1);
              resmem.Read(sbuffer^,len);
              s2:= StrPas(sbuffer);
              sout.Add('<td><P><font face="Tahoma,Arial" size="2">'+s2+'</FONT></P></td>');
              freemem(sbuffer);

              sout.Add('<td><P><font face="Tahoma,Arial" size="2"><a href="rrest://'+IntToStr(j)+'">Удалить</a></FONT></P></td>');

              sout.Add('</tr>');
             end;

            resmem.Free;
           end;
          Reg.CloseKey;
         end;
        finally
         Reg.Free;
        end;
        sout.Add('</table>');


 sout.Add('</td>');
 sout.Add('</table>');

 sout.Add('<center><P><font face="Tahoma,Arial" size="2"><a href="rsave://save">Сохранить результаты в файл</a></FONT></P></center>');

 sout.Add('<hr align="center" width="100%" size="1" noshade>');
 sout.Add('<p align="center"><font size="-2" face="Tahoma, Arial, Times New Roman">Copyright');
 sout.Add('  &copy; 2005-2011 <a href="http://siberia-soft.ru" target="_blank">Siberia-Soft</a></font></p>');
 sout.Add('</BODY>');
 sout.Add('</HTML>');

 sout.SaveToFile(GetTmpDir+'seek.html');
 sout.Free;

 if FileExists(GetTmpDir+'seek.html') then
  wb2.Navigate(GetTmpDir+'seek.html');

end;

procedure TMainForm.ToolButton4Click(Sender: TObject);
begin
 N19Click(Sender);
end;

procedure TMainForm.ImageButton16Click(Sender: TObject);
begin
  N19Click(Sender);
end;

procedure TMainForm.ImageButton15Click(Sender: TObject);
begin
 N19Click(Sender);
end;

procedure TMainForm.ImageButton24Click(Sender: TObject);
begin
 N19Click(Sender);
end;

procedure TMainForm.SeekMainPage;
var
 j,i : integer;
 Node : TTreeNode;
 HTMLDoc : THTMLDoc;
 ParamPage : string;
begin
   HomeIndex := 0;

   // Посмотрим параметры строки запуска
   if Length(paramstr(1))<>0 then
   begin
    ParamPage := paramstr(1);
    
    if Length(paramstr(2))<>0 then
     ParamPage := ParamPage + ' ' + paramstr(2);
    if Length(paramstr(3))<>0 then
     ParamPage := ParamPage + ' ' + paramstr(3);
    if Length(paramstr(4))<>0 then
     ParamPage := ParamPage + ' ' + paramstr(4);
    if Length(paramstr(5))<>0 then
     ParamPage := ParamPage + ' ' + paramstr(5);

    for I := 1 to ItemsTree.Items.Count - 1 do begin
    HTMLDoc := THTMLDoc(ItemsTree.Items[I].Data);
    if HTMLDoc.Nam = ParamPage then
    begin
     HomeIndex := i;
     ItemsTree.Selected := ItemsTree.Items[I];
     ContentIndex := ItemsTree.Selected.AbsoluteIndex;
     if HTMLDoc <> nil then
      if HTMLDoc.HTMLText.Size > 0 then
       Nav(HTMLDoc);
     break;
    end;
    end;
   end;

   // Поищем главную страницу
   if HomeIndex = 0 then
   begin
   HomeIndex := 0;
   for I := 1 to ItemsTree.Items.Count - 1 do begin
    HTMLDoc := THTMLDoc(ItemsTree.Items[I].Data);
    if HTMLDoc.MainPage then
    begin
     HomeIndex := i;
     ItemsTree.Selected := ItemsTree.Items[I];
     ContentIndex := ItemsTree.Selected.AbsoluteIndex;
     if HTMLDoc <> nil then
      if HTMLDoc.HTMLText.Size > 0 then
       Nav(HTMLDoc);
     break;
    end;
   end;
   end;

   // Поищем первую страницу
   if HomeIndex = 0 then
   begin
    for I := 1 to ItemsTree.Items.Count - 1 do begin
    HTMLDoc := THTMLDoc(ItemsTree.Items[I].Data);
    if HTMLDoc.HTMLType = 1 then
    begin
     HomeIndex := i;
     ItemsTree.Selected := ItemsTree.Items[I];
     ContentIndex := ItemsTree.Selected.AbsoluteIndex;
     if HTMLDoc <> nil then
      if HTMLDoc.HTMLText.Size > 0 then
       Nav(HTMLDoc);
     break;
    end;
    end;
   end;

 TreeSaveToFile;
end;


procedure TMainForm.TreeSaveToFile;
var
 z,i : integer;
 b : byte;
 F : TFileStream;
begin
 try
  F := TFilestream.Create(GetTmpDir+'tree.cfg',fmCreate);
  for i :=0 to ItemsTree.Items.Count-1 do
  begin
   z := ItemsTree.Items[i].AbsoluteIndex;
   F.Write(z,4);
   if ItemsTree.Items[i].Expanded then
    b := 1
   else
    b := 0;
   F.Write(b,1);
  end;
  F.Free;
 except
 end;
end;

procedure TMainForm.TreeLoadFromFile;
var
 z, i : integer;
 F : TMemoryStream;
 b : byte;
begin
 try
  F := TMemoryStream.Create;
  F.LoadFromFile(GetTmpDir+'tree.cfg');
  F.Position := 0;
  while F.Position < F.Size-5 do
  begin
   F.Read(z,4);
   F.Read(b,1);
   if b=1 then
    ItemsTree.Items[z].Expand(False)
   else
    ItemsTree.Items[z].Collapse(False);
  end;
  F.Free;
 except
 end;
end;

procedure TMainForm.ViewTree;
var
 j,i : integer;
 Node : TTreeNode;
 HTMLDoc : THTMLDoc;

begin
 ResPageNavigate := false;

 ItemsTree.Items.Clear;
 Node := ItemsTree.Items.Add(nil,DocsBook.ActiveBook.Nam);
 ItemsTree.Selected := Node;
 if ComboBox1.ItemIndex = 0 then
 // Содержание
 begin
   with DocsBook.ActiveBook do
   begin
      for I := 0 to Count - 1 do
       begin
        HTML[I].IDLevel0 := i+1;
        case HTML[I].HTMLType of
        0 :
                   begin
                      Node := ItemsTree.Items.AddChildObject(ItemsTree.Items[0], HTML[I].Nam, HTML[I]);
                      Node.ImageIndex := 0;
                      Node.SelectedIndex := 0;
                   end;
        1:
                   begin
                      Node := ItemsTree.Items.AddChildObject(ItemsTree.Items[0], HTML[I].Nam, HTML[I]);
                      Node.ImageIndex := 2;
                      Node.SelectedIndex := 2;
                   end;

        2:         if FirstGen then begin
                      TestList.AddObject(HTML[I].Nam, HTML[I]);
                   end;
        5:         if FirstGen then begin
                      FileList.AddObject(HTML[I].Nam, HTML[I]);
                   end;
        end;
        for J := 0 to HTML[I].Children.Count-1 do
          AddChildNodes(THTMLDoc(HTML[I].Children.Items[J]), Node, ItemsTree, True, FirstGen);
      end;
      ItemsTree.FullCollapse;
   end;

   if not FirstGen then
   begin
    TreeLoadFromFile;
    if ContentIndex  = 0 then
     SeekMainPage
    else
    begin
     ItemsTree.Selected := ItemsTree.Items[ContentIndex];
     ItemsTreeClick(nil);
    end;
   end
   else
    SeekMainPage;
 end
 else
 if ComboBox1.ItemIndex = 1 then
 // Тестовые задания
 begin
   with DocsBook.ActiveBook do
   begin
      for I := 0 to Count - 1 do begin
        HTML[I].IDLevel0 := i+1;
        case HTML[I].HTMLType of
        0 :
                   begin
                      Node := ItemsTree.Items.AddChildObject(ItemsTree.Items[0], HTML[I].Nam, HTML[I]);
                      Node.ImageIndex := 0;
                      Node.SelectedIndex := 0;
                   end;
        end;
        for J := 0 to HTML[I].Children.Count-1 do
          AddChildNodes(THTMLDoc(HTML[I].Children.Items[J]), Node, ItemsTree, False, False);
      end;
      ItemsTree.FullExpand;
      if not FirstGen then
      begin
       ItemsTree.Selected := ItemsTree.Items[TestTreeIndex];
       ItemsTreeClick(nil);
      end
      else
       ItemsTreeClick(nil);
   end;
 end
 else
 if ComboBox1.ItemIndex = 2 then
 // Файлы
 begin
   with DocsBook.ActiveBook do
   begin
      for I := 0 to Count - 1 do begin
        HTML[I].IDLevel0 := i+1;
        case HTML[I].HTMLType of
        0 :
                   begin
                      Node := ItemsTree.Items.AddChildObject(ItemsTree.Items[0], HTML[I].Nam, HTML[I]);
                      Node.ImageIndex := 0;
                      Node.SelectedIndex := 0;
                   end;
        end;
        for J := 0 to HTML[I].Children.Count-1 do
          AddChildNodes(THTMLDoc(HTML[I].Children.Items[J]), Node, ItemsTree, False, False);
      end;
      ItemsTree.FullExpand;
      if not FirstGen then
      begin
       ItemsTree.Selected := ItemsTree.Items[TestTreeIndex];
       ItemsTreeClick(nil);
      end
      else
       ItemsTreeClick(nil);
   end;
 end;

end;

procedure TMainForm.ComboBox1Change(Sender: TObject);
begin
 ViewTree;
end;

procedure TMainForm.ItemsTreeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 Navigate(false);
end;

procedure TMainForm.N20Click(Sender: TObject);
begin
 N20.Checked := not N20.Checked;
 ToolButtonFind.Down := N20.Checked;
 if ToolButtonFind.Down = true then
 begin
  ToolButtonFind.Down := true;
  Panel1.Visible := false;
  Splitter1.Visible := false;
  PanelSeek.Visible := true;
  Edit1.SetFocus;
 end
 else
  CloseBtn2Click(Sender);
end;

end.
