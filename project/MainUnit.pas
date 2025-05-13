unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Grids, Menus, ShlObj, ActiveX, StrUtils, ExtCtrls;

type
  // Типы параметров задач
  TParamType = (ptString, ptInteger, ptBoolean, ptStringList);

  // Описание параметра задачи
  TTaskParamInfo = record
    Name: array[0..255] of Char;
    Description: array[0..1023] of Char;
    ParamType: TParamType;
    Required: Boolean;
  end;

  // Значение параметра
  TTaskParamValue = record
    ParamType: TParamType;
    case Integer of
      0: (StringValue: array[0..1023] of Char);
      1: (IntValue: Integer);
      2: (BoolValue: Boolean);
      3: (StringListPtr: Pointer); // Указатель на TStringList
  end;

  // Информация о задаче
  TTaskInfo = record
    Name: array[0..255] of Char;
    Description: array[0..1023] of Char;
    ParamCount: Integer;
  end;

  // Прогресс выполнения задачи
  TTaskProgress = record
    TaskId: Integer;
    Progress: Integer; // 0-100
    Status: Integer;   // 0 - ожидание, 1 - выполняется, 2 - завершена, 3 - ошибка
    ErrorMessage: array[0..1023] of Char;
  end;

  // Результат выполнения задачи
  TTaskResult = record
    TaskId: Integer;
    Success: Boolean;
    Message: array[0..1023] of Char;
    ResultCount: Integer;
    ResultData: Pointer; // Указатель на дополнительные данные результата
  end;

  // Определение типов для функций DLL
  TGetTaskCount = function: Integer; stdcall;
  TGetTaskInfo = function(Index: Integer; var TaskInfo: TTaskInfo): Boolean; stdcall;
  TGetTaskParamInfo = function(TaskIndex, ParamIndex: Integer; var ParamInfo: TTaskParamInfo): Boolean; stdcall;
  TStartTask = function(TaskName: PChar; TaskId: Integer; ParamCount: Integer; ParamNames, ParamValues: Pointer): Boolean; stdcall;
  TGetTaskProgress = function(TaskId: Integer; var Progress: TTaskProgress): Boolean; stdcall;
  TGetTaskResult = function(TaskId: Integer; var TaskResult: TTaskResult): Boolean; stdcall;
  TGetTaskResultDetail = function(TaskId, Index: Integer; Buffer: PChar; BufSize: Integer): Boolean; stdcall;
  TStopTask = function(TaskId: Integer): Boolean; stdcall;
  TFreeTask = function(TaskId: Integer): Boolean; stdcall;

  // Структура для хранения информации о DLL
  TDLLInfo = record
    Handle: THandle;
    FileName: string;
    GetTaskCount: TGetTaskCount;
    GetTaskInfo: TGetTaskInfo;
    GetTaskParamInfo: TGetTaskParamInfo;
    StartTask: TStartTask;
    GetTaskProgress: TGetTaskProgress;
    GetTaskResult: TGetTaskResult;
    GetTaskResultDetail: TGetTaskResultDetail;
    StopTask: TStopTask;
    FreeTask: TFreeTask;
  end;

  // Структура для хранения информации о задаче из DLL
  TAvailableTask = record
    DLLIndex: Integer;
    TaskIndex: Integer;
    Name: string;
    Description: string;
    ParamCount: Integer;
    Params: array of TTaskParamInfo;
  end;

  // Структура для хранения информации о запущенных задачах
  TRunningTask = record
    TaskId: Integer;
    DLLIndex: Integer;
    TaskName: string;
    Progress: Integer;
    Status: Integer;
    StartTime: TDateTime;
  end;

  TMainForm = class(TForm)
    tmrProgress: TTimer;
    dlgOpenDLL: TOpenDialog;
    pmTaskMenu: TPopupMenu;
    miGetResult: TMenuItem;
    miShowResults: TMenuItem;
    miStopTask: TMenuItem;
    miFreeTask: TMenuItem;
    pnlMain: TPanel;
    pnlTasks: TPanel;
    statBar: TStatusBar;
    spl1: TSplitter;
    btnLoadDLL: TButton;
    lbDLLs: TListBox;
    btnUnloadDLL: TButton;
    btnStartTask: TButton;
    lbTasks: TListBox;
    lblTaskDescription: TLabel;
    pnlParams: TPanel;
    pnlParams2: TScrollBox;
    spl2: TSplitter;
    pnlTaskRun: TPanel;
    sgTasks: TStringGrid;
    spl3: TSplitter;
    btnStopTask: TButton;
    btnShowResults: TButton;
    btnFreeTask: TButton;
    pnlTaskLog: TPanel;
    memoLog: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnLoadDLLClick(Sender: TObject);
    procedure btnUnloadDLLClick(Sender: TObject);
    procedure lbDLLsClick(Sender: TObject);
    procedure lbTasksClick(Sender: TObject);
    procedure btnStartTaskClick(Sender: TObject);
    procedure tmrProgressTimer(Sender: TObject);
    procedure btnGetResultClick(Sender: TObject);
    procedure btnFreeTaskClick(Sender: TObject);
    procedure btnStopTaskClick(Sender: TObject);
    procedure btnShowResultsClick(Sender: TObject);
  private
    { Private declarations }
    DLLs: array of TDLLInfo;
    AvailableTasks: array of TAvailableTask;
    RunningTasks: array of TRunningTask;
    SelectedDLLIndex: Integer;
    SelectedTaskIndex: Integer;
    procedure LoadDLLFile(const FileName: string);
    procedure UnloadDLL(DLLIndex: Integer);
    procedure UnloadAllDLLs;
    procedure UpdateDLLList;
    procedure UpdateTaskList;
    procedure UpdateTasksStatus;
    procedure UpdateTaskGrid;
    procedure CreateParamControls(TaskIndex: Integer);
    procedure ClearParamControls;
    function GetParamValues(TaskIndex: Integer; var ParamNames: array of string; var ParamValues: array of TTaskParamValue): Boolean;
    procedure ShowTaskResults(TaskId: Integer);
    procedure Log(const Msg: string);
    procedure BrowseFileClick(Sender: TObject);
    procedure BrowseFolderClick(Sender: TObject);
    procedure EditNumberKeyPress(Sender: TObject; var Key: Char);
    procedure SelectDirectory(const Caption: string; var Directory: string);
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;
  NextTaskId: Integer;

implementation

{$R *.dfm}

const
  PARAM_TAG_BASE = 1000;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  SetLength(DLLs, 0);
  SetLength(AvailableTasks, 0);
  SetLength(RunningTasks, 0);
  SelectedDLLIndex := -1;
  SelectedTaskIndex := -1;
  tmrProgress.Enabled := True;

  // Инициализируем таблицу задач
  sgTasks.Cells[0, 0] := 'ID';
  sgTasks.Cells[1, 0] := 'Задача';
  sgTasks.Cells[2, 0] := 'Прогресс';
  sgTasks.Cells[3, 0] := 'Статус';
  sgTasks.Cells[4, 0] := 'Время запуска';
  sgTasks.RowCount := 1;
  NextTaskId := 1;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  UnloadAllDLLs;
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  i: Integer;
  HasRunningTasks: Boolean;
begin
  HasRunningTasks := False;
  for i := 0 to Length(RunningTasks) - 1 do
  begin
    if (RunningTasks[i].TaskId >= 0) and (RunningTasks[i].Status = 1) then
    begin
      HasRunningTasks := True;
      Break;
    end;
  end;

  if HasRunningTasks then
  begin
    if MessageDlg('Имеются выполняющиеся задачи. Остановить их и закрыть приложение?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      for i := 0 to Length(RunningTasks) - 1 do
      begin
        if (RunningTasks[i].TaskId >= 0) and (RunningTasks[i].Status = 1) and (RunningTasks[i].DLLIndex >= 0) and (RunningTasks[i].DLLIndex < Length(DLLs)) then
        begin
          DLLs[RunningTasks[i].DLLIndex].StopTask(RunningTasks[i].TaskId);
        end;
      end;
      CanClose := True;
    end
    else
      CanClose := False;
  end
  else
    CanClose := True;
end;

procedure TMainForm.LoadDLLFile(const FileName: string);
var
  DLLIndex: Integer;
  DLLInfo: TDLLInfo;
begin
  for DLLIndex := 0 to Length(DLLs) - 1 do
  begin
    if CompareText(DLLs[DLLIndex].FileName, FileName) = 0 then
    begin
      Log(Format('DLL "%s" уже загружена', [FileName]));
      Exit;
    end;
  end;

  DLLInfo.Handle := LoadLibrary(PChar(FileName));
  if DLLInfo.Handle = 0 then
  begin
    Log(Format('Ошибка загрузки DLL "%s": %s', [FileName, SysErrorMessage(GetLastError)]));
    Exit;
  end;

  @DLLInfo.GetTaskCount := GetProcAddress(DLLInfo.Handle, 'GetTaskCount');
  @DLLInfo.GetTaskInfo := GetProcAddress(DLLInfo.Handle, 'GetTaskInfo');
  @DLLInfo.GetTaskParamInfo := GetProcAddress(DLLInfo.Handle, 'GetTaskParamInfo');
  @DLLInfo.StartTask := GetProcAddress(DLLInfo.Handle, 'StartTask');
  @DLLInfo.GetTaskProgress := GetProcAddress(DLLInfo.Handle, 'GetTaskProgress');
  @DLLInfo.GetTaskResult := GetProcAddress(DLLInfo.Handle, 'GetTaskResult');
  @DLLInfo.GetTaskResultDetail := GetProcAddress(DLLInfo.Handle, 'GetTaskResultDetail');
  @DLLInfo.StopTask := GetProcAddress(DLLInfo.Handle, 'StopTask');
  @DLLInfo.FreeTask := GetProcAddress(DLLInfo.Handle, 'FreeTask');

  if not Assigned(DLLInfo.GetTaskCount) or not Assigned(DLLInfo.GetTaskInfo) or not Assigned(DLLInfo.GetTaskParamInfo) or
     not Assigned(DLLInfo.StartTask) or not Assigned(DLLInfo.GetTaskProgress) or not Assigned(DLLInfo.GetTaskResult) or
     not Assigned(DLLInfo.GetTaskResultDetail) or not Assigned(DLLInfo.StopTask) or not Assigned(DLLInfo.FreeTask) then
  begin
    Log(Format('DLL "%s" не содержит необходимых функций', [FileName]));
    FreeLibrary(DLLInfo.Handle);
    Exit;
  end;

  DLLInfo.FileName := FileName;
  DLLIndex := Length(DLLs);
  SetLength(DLLs, DLLIndex + 1);
  DLLs[DLLIndex] := DLLInfo;
  Log(Format('DLL "%s" успешно загружена', [ExtractFileName(FileName)]));

  UpdateDLLList;
end;

procedure TMainForm.btnLoadDLLClick(Sender: TObject);
begin
  if dlgOpenDLL.Execute then
  begin
    LoadDLLFile(dlgOpenDLL.FileName);
  end;
end;

procedure TMainForm.btnUnloadDLLClick(Sender: TObject);
var
  DLLIndex: Integer;
begin
  DLLIndex := lbDLLs.ItemIndex;
  if (DLLIndex >= 0) and (DLLIndex < Length(DLLs)) then
  begin
    UnloadDLL(DLLIndex);
    UpdateDLLList;
    UpdateTaskList;
    ClearParamControls;
  end;
end;

procedure TMainForm.UnloadDLL(DLLIndex: Integer);
var
  i, j: Integer;
  TasksToRemove: TList;
  DLLFileName: string;
begin
  if (DLLIndex < 0) or (DLLIndex >= Length(DLLs)) then
    Exit;

  DLLFileName := ExtractFileName(DLLs[DLLIndex].FileName);
  TasksToRemove := TList.Create;
  try
    for i := 0 to Length(RunningTasks) - 1 do
    begin
      if RunningTasks[i].DLLIndex = DLLIndex then
      begin
        if RunningTasks[i].Status = 1 then
          DLLs[DLLIndex].StopTask(RunningTasks[i].TaskId);
        DLLs[DLLIndex].FreeTask(RunningTasks[i].TaskId);
        TasksToRemove.Add(Pointer(i));
      end;
    end;

    for i := TasksToRemove.Count - 1 downto 0 do
    begin
      j := Integer(TasksToRemove[i]);
      if j < Length(RunningTasks) - 1 then
        RunningTasks[j] := RunningTasks[Length(RunningTasks) - 1];
      SetLength(RunningTasks, Length(RunningTasks) - 1);
    end;
  finally
    TasksToRemove.Free;
  end;

  FreeLibrary(DLLs[DLLIndex].Handle);
  Log(Format('DLL "%s" выгружена', [DLLFileName]));

  if DLLIndex < Length(DLLs) - 1 then
    DLLs[DLLIndex] := DLLs[Length(DLLs) - 1];
  SetLength(DLLs, Length(DLLs) - 1);

  i := 0;
  while i < Length(AvailableTasks) do
  begin
    if AvailableTasks[i].DLLIndex = DLLIndex then
    begin
      if i < High(AvailableTasks) then
        AvailableTasks[i] := AvailableTasks[High(AvailableTasks)];
      SetLength(AvailableTasks, Length(AvailableTasks) - 1);
    end
    else
    begin
      if AvailableTasks[i].DLLIndex > DLLIndex then
        Dec(AvailableTasks[i].DLLIndex);
      Inc(i);
    end;
  end;

  UpdateTaskGrid;
end;

procedure TMainForm.UnloadAllDLLs;
var
  i: Integer;
begin
  for i := Length(DLLs) - 1 downto 0 do
    UnloadDLL(i);
  SetLength(DLLs, 0);
  SetLength(AvailableTasks, 0);
  SetLength(RunningTasks, 0);
  lbDLLs.Clear;
  lbTasks.Clear;
  ClearParamControls;
  btnUnloadDLL.Enabled := False;
  btnStartTask.Enabled := False;
end;

procedure TMainForm.UpdateDLLList;
var
  i: Integer;
begin
  lbDLLs.Clear;
  for i := 0 to Length(DLLs) - 1 do
    lbDLLs.Items.Add(ExtractFileName(DLLs[i].FileName));
  btnUnloadDLL.Enabled := lbDLLs.Items.Count > 0;
end;

procedure TMainForm.UpdateTaskList;
var
  i, j, k, TaskCount: Integer;
  TaskInfo: TTaskInfo;
  ParamInfo: TTaskParamInfo;
  TaskIndexInAvailable: Integer;
begin
  lbTasks.Clear;
  if lbDLLs.ItemIndex < 0 then
    Exit;

  SelectedDLLIndex := lbDLLs.ItemIndex;

  TaskCount := DLLs[SelectedDLLIndex].GetTaskCount;
  for j := 0 to TaskCount - 1 do
  begin
    if DLLs[SelectedDLLIndex].GetTaskInfo(j, TaskInfo) then
    begin
      TaskIndexInAvailable := Length(AvailableTasks);
      SetLength(AvailableTasks, TaskIndexInAvailable + 1);
      AvailableTasks[TaskIndexInAvailable].DLLIndex := SelectedDLLIndex;
      AvailableTasks[TaskIndexInAvailable].TaskIndex := j;
      AvailableTasks[TaskIndexInAvailable].Name := TaskInfo.Name;
      AvailableTasks[TaskIndexInAvailable].Description := TaskInfo.Description;
      AvailableTasks[TaskIndexInAvailable].ParamCount := TaskInfo.ParamCount;
      SetLength(AvailableTasks[TaskIndexInAvailable].Params, TaskInfo.ParamCount);

      for k := 0 to TaskInfo.ParamCount - 1 do
      begin
        if DLLs[SelectedDLLIndex].GetTaskParamInfo(j, k, ParamInfo) then
          AvailableTasks[TaskIndexInAvailable].Params[k] := ParamInfo;
      end;

      lbTasks.Items.AddObject(TaskInfo.Name, TObject(TaskIndexInAvailable));
    end;
  end;

  if lbTasks.Items.Count > 0 then
  begin
    lbTasks.ItemIndex := 0;
    lbTasksClick(nil);
    btnStartTask.Enabled := True;
  end
  else
  begin
    btnStartTask.Enabled := False;
    lblTaskDescription.Caption := '';
    ClearParamControls;
  end;
end;

procedure TMainForm.lbDLLsClick(Sender: TObject);
begin
  UpdateTaskList;
end;

procedure TMainForm.lbTasksClick(Sender: TObject);
var
  TaskIndexInAvailable: Integer;
begin
  ClearParamControls;

  if lbTasks.ItemIndex < 0 then
    Exit;

  TaskIndexInAvailable := Integer(lbTasks.Items.Objects[lbTasks.ItemIndex]);

  if (TaskIndexInAvailable >= 0) and (TaskIndexInAvailable < Length(AvailableTasks)) then
  begin
    SelectedTaskIndex := TaskIndexInAvailable;
    lblTaskDescription.Caption := AvailableTasks[SelectedTaskIndex].Description;
    CreateParamControls(SelectedTaskIndex);
  end
  else
  begin
    SelectedTaskIndex := -1;
    lblTaskDescription.Caption := '';
  end;
end;

procedure TMainForm.CreateParamControls(TaskIndex: Integer);
var
  i: Integer;
  ParamInfo: TTaskParamInfo;
  Label1, DescLabel: TLabel;
  Edit1: TEdit;
  CheckBox1: TCheckBox;
  ComboBox1: TComboBox;
  Button1: TButton;
  Y: Integer;
  CommentX: Integer;
  RequiredStr: string;
  s: string;
begin
  ClearParamControls;
  if (TaskIndex < 0) or (TaskIndex >= Length(AvailableTasks)) then
    Exit;

  Y := 10;

  if AvailableTasks[TaskIndex].ParamCount > 0 then
  begin
    Label1 := TLabel.Create(pnlParams2);
    Label1.Parent := pnlParams2;
    Label1.Left := 10;
    Label1.Top := Y;
    Label1.Caption := 'Параметры задачи:';
    Label1.Font.Style := [fsBold];
    Label1.Tag := PARAM_TAG_BASE;
    Inc(Y, 25);
  end;

  for i := 0 to AvailableTasks[TaskIndex].ParamCount - 1 do
  begin
    ParamInfo := AvailableTasks[TaskIndex].Params[i];
    CommentX := 0;

    if ParamInfo.Required then
      RequiredStr := ' *'
    else
      RequiredStr := '';

    Label1 := TLabel.Create(pnlParams2);
    Label1.Parent := pnlParams2;
    Label1.Left := 10;
    Label1.Top := Y;
    Label1.Caption := ParamInfo.Name + RequiredStr + ':';
    Label1.Hint := ParamInfo.Description;
    Label1.ShowHint := True;
    Label1.Tag := PARAM_TAG_BASE + i;

    case ParamInfo.ParamType of
      ptString:
        begin
        Edit1 := TEdit.Create(pnlParams2);
        Edit1.Parent := pnlParams2;
        Edit1.Left := 150;
        Edit1.Top := Y;
        Edit1.Width := 250;
        Edit1.Tag := PARAM_TAG_BASE + i;
        CommentX := Edit1.Left + Edit1.Width + 5;
        s := LowerCase(string(ParamInfo.Name));
        if (Pos('путь', s) > 0) or (Pos('файл', s) > 0) or (Pos('path', s) > 0) or (Pos('dir', s) > 0) then
        begin
          Edit1.Width := 220;
          Button1 := TButton.Create(pnlParams2);
          Button1.Parent := pnlParams2;
          Button1.Left := 375;
          Button1.Top := Y - 1;
          Button1.Width := 25;
          Button1.Height := 25;
          Button1.Caption := '...';
          Button1.Tag := PARAM_TAG_BASE + i;
          CommentX := Button1.Left + Button1.Width + 5;
          if (Pos('file', s) > 0) or (Pos('файл', s) > 0) then
            Button1.OnClick := BrowseFileClick
          else
            Button1.OnClick := BrowseFolderClick;
        end;
      end;
      ptInteger:
        begin
          Edit1 := TEdit.Create(pnlParams2);
          Edit1.Parent := pnlParams2;
          Edit1.Left := 150;
          Edit1.Top := Y;
          Edit1.Width := 100;
          Edit1.Tag := PARAM_TAG_BASE + i;
          CommentX := Edit1.Left + Edit1.Width + 5;
          Edit1.OnKeyPress := EditNumberKeyPress;
        end;
      ptBoolean:
        begin
          CheckBox1 := TCheckBox.Create(pnlParams2);
          CheckBox1.Parent := pnlParams2;
          CheckBox1.Left := 150;
          CheckBox1.Top := Y;
          CheckBox1.Width := 25;
          CheckBox1.Caption := '';
          CheckBox1.Tag := PARAM_TAG_BASE + i;
          CommentX := CheckBox1.Left + CheckBox1.Width + 5;
        end;
      ptStringList:
        begin
          ComboBox1 := TComboBox.Create(pnlParams2);
          ComboBox1.Parent := pnlParams2;
          ComboBox1.Left := 150;
          ComboBox1.Top := Y;
          ComboBox1.Width := 250;
          ComboBox1.Tag := PARAM_TAG_BASE + i;
          CommentX := ComboBox1.Left + ComboBox1.Width + 5;
        end;
    end;

    DescLabel := TLabel.Create(pnlParams2);
    DescLabel.Parent := pnlParams2;
    DescLabel.Left := CommentX;
    DescLabel.Top := Y;
    DescLabel.Caption := ParamInfo.Description;
    DescLabel.Font.Style := [fsItalic];
    DescLabel.Font.Size := 8;
    DescLabel.Font.Color := clGrayText;
    DescLabel.Tag := PARAM_TAG_BASE + i;

    Inc(Y, 23);
  end;

  if AvailableTasks[TaskIndex].ParamCount = 0 then
  begin
    Label1 := TLabel.Create(pnlParams2);
    Label1.Parent := pnlParams2;
    Label1.Left := 10;
    Label1.Top := Y;
    Label1.Caption := 'Эта задача не имеет настраиваемых параметров';
    Label1.Font.Style := [fsItalic];
    Label1.Tag := PARAM_TAG_BASE;
  end;
end;

procedure TMainForm.ClearParamControls;
var
  i: Integer;
begin
  for i := pnlParams2.ControlCount - 1 downto 0 do
  begin
    if pnlParams2.Controls[i].Tag >= PARAM_TAG_BASE then
      pnlParams2.Controls[i].Free;
  end;
end;

function TMainForm.GetParamValues(TaskIndex: Integer; var ParamNames: array of string; var ParamValues: array of TTaskParamValue): Boolean;
var
  i, ControlIndex: Integer;
  ParamInfo: TTaskParamInfo;
  Control: TControl;
  Edit1: TEdit;
  CheckBox1: TCheckBox;
  ComboBox1: TComboBox;
  IntValue: Integer;
  StringList: TStringList;
begin
  Result := True;
  if (TaskIndex < 0) or (TaskIndex >= Length(AvailableTasks)) then
    Exit;

  for i := 0 to AvailableTasks[TaskIndex].ParamCount - 1 do
  begin
    ParamInfo := AvailableTasks[TaskIndex].Params[i];
    ParamNames[i] := ParamInfo.Name;
    ParamValues[i].ParamType := ParamInfo.ParamType;

    Control := nil;
    for ControlIndex := 0 to pnlParams2.ControlCount - 1 do
    begin
      if (pnlParams2.Controls[ControlIndex].Tag = PARAM_TAG_BASE + i) and
         not (pnlParams2.Controls[ControlIndex] is TLabel) and
         not (pnlParams2.Controls[ControlIndex] is TButton) then
      begin
        Control := pnlParams2.Controls[ControlIndex];
        Break;
      end;
    end;

    if Control = nil then
      Continue;

    case ParamInfo.ParamType of
      ptString:
        begin
          Edit1 := TEdit(Control);
          StrPCopy(ParamValues[i].StringValue, Edit1.Text);
          if ParamInfo.Required and (Edit1.Text = '') then
          begin
            MessageDlg(Format('Параметр "%s" является обязательным', [ParamInfo.Name]), mtError, [mbOK], 0);
            Edit1.SetFocus;
            Result := False;
            Exit;
          end;
        end;
      ptInteger:
        begin
          Edit1 := TEdit(Control);
          if ParamInfo.Required and (Edit1.Text = '') then
          begin
            MessageDlg(Format('Параметр "%s" является обязательным', [ParamInfo.Name]), mtError, [mbOK], 0);
            Edit1.SetFocus;
            Result := False;
            Exit;
          end;

          if Edit1.Text <> '' then
          begin
            try
              IntValue := StrToInt(Edit1.Text);
              ParamValues[i].IntValue := IntValue;
            except
              on E: Exception do
              begin
                MessageDlg(Format('Ошибка преобразования: %s', [E.Message]), mtError, [mbOK], 0);
                Edit1.SetFocus;
                Result := False;
                Exit;
              end;
            end;
          end
          else
            ParamValues[i].IntValue := 0;
        end;
      ptBoolean:
        begin
          CheckBox1 := TCheckBox(Control);
          ParamValues[i].BoolValue := CheckBox1.Checked;
        end;
      ptStringList:
        begin
          ComboBox1 := TComboBox(Control);
          if ParamInfo.Required and (ComboBox1.Text = '') then
          begin
            MessageDlg(Format('Параметр "%s" является обязательным', [ParamInfo.Name]), mtError, [mbOK], 0);
            ComboBox1.SetFocus;
            Result := False;
            Exit;
          end;

          StringList := TStringList.Create;
          StringList.Add(ComboBox1.Text);
          ParamValues[i].StringListPtr := Pointer(StringList);
        end;
    end;
  end;
end;

procedure TMainForm.btnStartTaskClick(Sender: TObject);
var
  TaskIndexInAvailable, DLLIndex, TaskId, RunningTaskIndex, i: Integer;
  ParamNames: array of string;
  ParamValues: array of TTaskParamValue;
  StringLists: array of TStringList;
begin
  if lbTasks.ItemIndex < 0 then
    Exit;

  TaskIndexInAvailable := Integer(lbTasks.Items.Objects[lbTasks.ItemIndex]);
  if (TaskIndexInAvailable < 0) or (TaskIndexInAvailable >= Length(AvailableTasks)) then
    Exit;

  DLLIndex := AvailableTasks[TaskIndexInAvailable].DLLIndex;
  if (DLLIndex < 0) or (DLLIndex >= Length(DLLs)) then
    Exit;

  SetLength(ParamNames, AvailableTasks[TaskIndexInAvailable].ParamCount);
  SetLength(ParamValues, AvailableTasks[TaskIndexInAvailable].ParamCount);
  SetLength(StringLists, AvailableTasks[TaskIndexInAvailable].ParamCount);

  for i := 0 to AvailableTasks[TaskIndexInAvailable].ParamCount - 1 do
    StringLists[i] := nil;

  try
    if AvailableTasks[TaskIndexInAvailable].ParamCount > 0 then
    begin
      if not GetParamValues(TaskIndexInAvailable, ParamNames, ParamValues) then
      begin
        for i := 0 to AvailableTasks[TaskIndexInAvailable].ParamCount - 1 do
        begin
          if (ParamValues[i].ParamType = ptStringList) and (StringLists[i] <> nil) then
            StringLists[i].Free;
        end;
        Exit;
      end;

      for i := 0 to AvailableTasks[TaskIndexInAvailable].ParamCount - 1 do
      begin
        if ParamValues[i].ParamType = ptStringList then
          StringLists[i] := TStringList(ParamValues[i].StringListPtr);
      end;
    end;

    TaskId := NextTaskId;
    Inc(NextTaskId);

    if DLLs[DLLIndex].StartTask(PChar(AvailableTasks[TaskIndexInAvailable].Name), TaskId, AvailableTasks[TaskIndexInAvailable].ParamCount, @ParamNames[0], @ParamValues[0]) then
    begin
      RunningTaskIndex := Length(RunningTasks);
      SetLength(RunningTasks, RunningTaskIndex + 1);
      RunningTasks[RunningTaskIndex].TaskId := TaskId;
      RunningTasks[RunningTaskIndex].DLLIndex := DLLIndex;
      RunningTasks[RunningTaskIndex].TaskName := AvailableTasks[TaskIndexInAvailable].Name;
      RunningTasks[RunningTaskIndex].Progress := 0;
      RunningTasks[RunningTaskIndex].Status := 0;
      RunningTasks[RunningTaskIndex].StartTime := Now;

      UpdateTaskGrid;

      if AvailableTasks[TaskIndexInAvailable].ParamCount > 0 then
      begin
        Log(Format('Задача "%s" запущена с ID: %d и следующими параметрами:', [AvailableTasks[TaskIndexInAvailable].Name, TaskId]));
        for i := 0 to AvailableTasks[TaskIndexInAvailable].ParamCount - 1 do
        begin
          case ParamValues[i].ParamType of
            ptString:
              Log(Format('  - %s: %s', [ParamNames[i], ParamValues[i].StringValue]));
            ptInteger:
              Log(Format('  - %s: %d', [ParamNames[i], ParamValues[i].IntValue]));
            ptBoolean:
              Log(Format('  - %s: %s', [ParamNames[i], BoolToStr(ParamValues[i].BoolValue, True)]));
            ptStringList:
              if StringLists[i] <> nil then
                Log(Format('  - %s: %s', [ParamNames[i], StringReplace(StringLists[i].Text, #13#10, ' ', [rfReplaceAll])]));
          end;
        end;
      end
      else
        Log(Format('Задача "%s" запущена с ID: %d (без параметров)', [AvailableTasks[TaskIndexInAvailable].Name, TaskId]));
    end
    else
      Log(Format('Ошибка запуска задачи "%s"', [AvailableTasks[TaskIndexInAvailable].Name]));

  finally
    for i := 0 to AvailableTasks[TaskIndexInAvailable].ParamCount - 1 do
    begin
      if (ParamValues[i].ParamType = ptStringList) and (StringLists[i] <> nil) then
        StringLists[i].Free;
    end;
    SetLength(ParamNames, 0);
    SetLength(ParamValues, 0);
    SetLength(StringLists, 0);
  end;
end;

procedure TMainForm.UpdateTasksStatus;
var
  i: Integer;
  Progress: TTaskProgress;
begin
  for i := 0 to Length(RunningTasks) - 1 do
  begin
    if (RunningTasks[i].DLLIndex >= 0) and (RunningTasks[i].DLLIndex < Length(DLLs)) and
       Assigned(DLLs[RunningTasks[i].DLLIndex].GetTaskProgress) then
    begin
      if DLLs[RunningTasks[i].DLLIndex].GetTaskProgress(RunningTasks[i].TaskId, Progress) then
      begin
        RunningTasks[i].Progress := Progress.Progress;
        RunningTasks[i].Status := Progress.Status;
      end;
    end;
  end;
end;

procedure TMainForm.tmrProgressTimer(Sender: TObject);
var
  i, TaskId: Integer;
begin
  UpdateTasksStatus;
  UpdateTaskGrid;

  if (sgTasks.Row > 0) and (sgTasks.Row < sgTasks.RowCount) then
  begin
    try
      TaskId := StrToInt(sgTasks.Cells[0, sgTasks.Row]);
      for i := 0 to Length(RunningTasks) - 1 do
      begin
        if RunningTasks[i].TaskId = TaskId then
        begin
          case RunningTasks[i].Status of
            0: statBar.Panels[1].Text := 'Ожидание';
            1: statBar.Panels[1].Text := Format('Выполняется: %d%%', [RunningTasks[i].Progress]);
            2: statBar.Panels[1].Text := 'Завершена';
            3: statBar.Panels[1].Text := 'Ошибка';
          else
            statBar.Panels[1].Text := Format('Неизвестный статус: %d', [RunningTasks[i].Status]);
          end;
          btnShowResults.Enabled := (RunningTasks[i].Status = 2) or (RunningTasks[i].Status = 3);
          btnStopTask.Enabled := (RunningTasks[i].Status = 0) or (RunningTasks[i].Status = 1);
          btnFreeTask.Enabled := (RunningTasks[i].Status <> 1);
          Break;
        end;
      end;
    except
      statBar.Panels[1].Text := '';
      btnShowResults.Enabled := False;
      btnStopTask.Enabled := False;
      btnFreeTask.Enabled := False;
    end;
  end
  else
  begin
    statBar.Panels[1].Text := '';
    btnShowResults.Enabled := False;
    btnStopTask.Enabled := False;
    btnFreeTask.Enabled := False;
  end;
end;

procedure TMainForm.UpdateTaskGrid;
var
  i, Row: Integer;
  StatusText: string;
begin
  sgTasks.RowCount := Length(RunningTasks) + 1;
  for i := 0 to Length(RunningTasks) - 1 do
  begin
    Row := i + 1;
    sgTasks.Cells[0, Row] := IntToStr(RunningTasks[i].TaskId);
    sgTasks.Cells[1, Row] := RunningTasks[i].TaskName;
    sgTasks.Cells[2, Row] := Format('%d%%', [RunningTasks[i].Progress]);

    case RunningTasks[i].Status of
      0: StatusText := 'Ожидание';
      1: StatusText := 'Выполняется';
      2: StatusText := 'Завершена';
      3: StatusText := 'Ошибка';
    else
      StatusText := Format('Неизвестный (%d)', [RunningTasks[i].Status]);
    end;

    sgTasks.Cells[3, Row] := StatusText;
    sgTasks.Cells[4, Row] := FormatDateTime('yyyy-mm-dd hh:nn:ss', RunningTasks[i].StartTime);
  end;

  if Length(RunningTasks) = 0 then
  begin
    sgTasks.RowCount := 2;
    sgTasks.Cells[0, 1] := '';
    sgTasks.Cells[1, 1] := '';
    sgTasks.Cells[2, 1] := '';
    sgTasks.Cells[3, 1] := '';
    sgTasks.Cells[4, 1] := '';
  end;
end;

procedure TMainForm.btnGetResultClick(Sender: TObject);
var
  i, TaskId: Integer;
  TaskResult: TTaskResult;
  ResultMessage: string;
begin
  if (sgTasks.Row <= 0) or (sgTasks.Row >= sgTasks.RowCount) then
    Exit;

  try
    TaskId := StrToInt(sgTasks.Cells[0, sgTasks.Row]);
    for i := 0 to Length(RunningTasks) - 1 do
    begin
      if (RunningTasks[i].TaskId = TaskId) and (RunningTasks[i].DLLIndex >= 0) and (RunningTasks[i].DLLIndex < Length(DLLs)) then
      begin
        if DLLs[RunningTasks[i].DLLIndex].GetTaskResult(TaskId, TaskResult) then
        begin
          if TaskResult.Success then
            ResultMessage := 'Задача успешно выполнена'
          else
            ResultMessage := 'Ошибка выполнения задачи';
          ResultMessage := ResultMessage + ': ' + TaskResult.Message;
          MessageDlg(ResultMessage, mtInformation, [mbOK], 0);
          Log(Format('Результат задачи %d: %s', [TaskId, ResultMessage]));
        end
        else
          Log(Format('Ошибка получения результата задачи %d', [TaskId]));
        Break;
      end;
    end;
  except
    on E: Exception do
      Log(Format('Ошибка при получении результата: %s', [E.Message]));
  end;
end;

procedure TMainForm.btnStopTaskClick(Sender: TObject);
var
  i, TaskId: Integer;
begin
  if (sgTasks.Row <= 0) or (sgTasks.Row >= sgTasks.RowCount) then
    Exit;

  try
    TaskId := StrToInt(sgTasks.Cells[0, sgTasks.Row]);
    for i := 0 to Length(RunningTasks) - 1 do
    begin
      if (RunningTasks[i].TaskId = TaskId) and (RunningTasks[i].DLLIndex >= 0) and (RunningTasks[i].DLLIndex < Length(DLLs)) then
      begin
        if DLLs[RunningTasks[i].DLLIndex].StopTask(TaskId) then
        begin
          Log(Format('Задача %d остановлена', [TaskId]));
          RunningTasks[i].Status := 3;
          UpdateTaskGrid;
          btnShowResults.Enabled := True;
          btnStopTask.Enabled := False;
          btnFreeTask.Enabled := True;
        end
        else
          Log(Format('Ошибка остановки задачи %d', [TaskId]));
        Break;
      end;
    end;
  except
    on E: Exception do
      Log(Format('Ошибка при остановке задачи: %s', [E.Message]));
  end;
end;

procedure TMainForm.btnFreeTaskClick(Sender: TObject);
var
  i, TaskId: Integer;
begin
  if (sgTasks.Row <= 0) or (sgTasks.Row >= sgTasks.RowCount) then
    Exit;

  try
    TaskId := StrToInt(sgTasks.Cells[0, sgTasks.Row]);
    for i := 0 to Length(RunningTasks) - 1 do
    begin
      if (RunningTasks[i].TaskId = TaskId) and (RunningTasks[i].DLLIndex >= 0) and (RunningTasks[i].DLLIndex < Length(DLLs)) then
      begin
        if RunningTasks[i].Status = 1 then
        begin
          MessageDlg('Невозможно освободить выполняющуюся задачу. Сначала остановите её.', mtWarning, [mbOK], 0);
          Exit;
        end;

        if DLLs[RunningTasks[i].DLLIndex].FreeTask(TaskId) then
        begin
          Log(Format('Задача %d освобождена', [TaskId]));
          if i < Length(RunningTasks) - 1 then
            RunningTasks[i] := RunningTasks[Length(RunningTasks) - 1];
          SetLength(RunningTasks, Length(RunningTasks) - 1);
          UpdateTaskGrid;
          statBar.Panels[1].Text := '';
          btnShowResults.Enabled := False;
          btnStopTask.Enabled := False;
          btnFreeTask.Enabled := False;
        end
        else
          Log(Format('Ошибка освобождения задачи %d', [TaskId]));
        Break;
      end;
    end;
  except
    on E: Exception do
      Log(Format('Ошибка при освобождении задачи: %s', [E.Message]));
  end;
end;

procedure TMainForm.btnShowResultsClick(Sender: TObject);
var
  i, TaskId: Integer;
begin
  if (sgTasks.Row <= 0) or (sgTasks.Row >= sgTasks.RowCount) then
    Exit;

  try
    TaskId := StrToInt(sgTasks.Cells[0, sgTasks.Row]);
    for i := 0 to Length(RunningTasks) - 1 do
    begin
      if (RunningTasks[i].TaskId = TaskId) and (RunningTasks[i].DLLIndex >= 0) and (RunningTasks[i].DLLIndex < Length(DLLs)) then
      begin
        ShowTaskResults(TaskId);
        Break;
      end;
    end;
  except
    on E: Exception do
      Log(Format('Ошибка при показе результатов: %s', [E.Message]));
  end;
end;

procedure TMainForm.ShowTaskResults(TaskId: Integer);
var
  i, j, ResultCount: Integer;
  TaskResult: TTaskResult;
  ResultForm: TForm;
  Memo: TMemo;
  Buffer: array[0..1023] of Char;
begin
  for i := 0 to Length(RunningTasks) - 1 do
  begin
    if (RunningTasks[i].TaskId = TaskId) and (RunningTasks[i].DLLIndex >= 0) and (RunningTasks[i].DLLIndex < Length(DLLs)) then
    begin
      if DLLs[RunningTasks[i].DLLIndex].GetTaskResult(TaskId, TaskResult) then
      begin
        ResultForm := TForm.Create(Application);
        try
          ResultForm.Caption := Format('Результаты задачи %s (ID: %d)', [RunningTasks[i].TaskName, TaskId]);
          ResultForm.Width := 600;
          ResultForm.Height := 400;
          ResultForm.Position := poScreenCenter;
          Memo := TMemo.Create(ResultForm);
          Memo.Parent := ResultForm;
          Memo.Align := alClient;
          Memo.ScrollBars := ssBoth;
          Memo.ReadOnly := True;

          Memo.Lines.Add('Статус: ' + BoolToStr(TaskResult.Success, True));
          Memo.Lines.Add('Сообщение: ' + TaskResult.Message);
          Memo.Lines.Add('');
          Memo.Lines.Add('Детальные результаты:');
          ResultCount := TaskResult.ResultCount;
          for j := 0 to ResultCount - 1 do
          begin
            FillChar(Buffer, SizeOf(Buffer), 0);
            if DLLs[RunningTasks[i].DLLIndex].GetTaskResultDetail(TaskId, j, Buffer, SizeOf(Buffer)) then
              Memo.Lines.Add(Buffer);
          end;
          ResultForm.ShowModal;
        finally
          ResultForm.Free;
        end;
      end
      else
        MessageDlg('Не удалось получить результаты задачи', mtError, [mbOK], 0);
      Break;
    end;
  end;
end;

procedure TMainForm.BrowseFileClick(Sender: TObject);
var
  Button: TButton;
  Edit: TEdit;
  i: Integer;
  ParamTag: Integer;
  OpenDialog: TOpenDialog;
begin
  Button := TButton(Sender);
  ParamTag := Button.Tag;
  for i := 0 to pnlParams2.ControlCount - 1 do
  begin
    if (pnlParams2.Controls[i] is TEdit) and (pnlParams2.Controls[i].Tag = ParamTag) then
    begin
      Edit := TEdit(pnlParams2.Controls[i]);
      OpenDialog := TOpenDialog.Create(Self);
      try
        OpenDialog.Title := 'Выберите файл';
        OpenDialog.Filter := 'Все файлы (*.*)|*.*';
        if OpenDialog.Execute then
          Edit.Text := OpenDialog.FileName;
      finally
        OpenDialog.Free;
      end;
      Break;
    end;
  end;
end;

procedure TMainForm.BrowseFolderClick(Sender: TObject);
var
  Button: TButton;
  Edit: TEdit;
  i: Integer;
  ParamTag: Integer;
  Directory: string;
begin
  Button := TButton(Sender);
  ParamTag := Button.Tag;
  for i := 0 to pnlParams2.ControlCount - 1 do
  begin
    if (pnlParams2.Controls[i] is TEdit) and (pnlParams2.Controls[i].Tag = ParamTag) then
    begin
      Edit := TEdit(pnlParams2.Controls[i]);
      Directory := Edit.Text;
      SelectDirectory('Выберите каталог', Directory);
      Edit.Text := Directory;
      Break;
    end;
  end;
end;

procedure TMainForm.SelectDirectory(const Caption: string; var Directory: string);
var
  BrowseInfo: TBrowseInfo;
  Buffer: array[0..MAX_PATH] of Char;
  ItemIdList: PItemIDList;
  ShellMalloc: IMalloc;
begin
  if SHGetMalloc(ShellMalloc) = S_OK then
  begin
    FillChar(BrowseInfo, SizeOf(BrowseInfo), 0);
    BrowseInfo.hwndOwner := Handle;
    BrowseInfo.pszDisplayName := Buffer;
    BrowseInfo.lpszTitle := PChar(Caption);
    BrowseInfo.ulFlags := BIF_RETURNONLYFSDIRS or BIF_NEWDIALOGSTYLE;
    ItemIdList := SHBrowseForFolder(BrowseInfo);
    if ItemIdList <> nil then
    begin
      SHGetPathFromIDList(ItemIdList, Buffer);
      Directory := Buffer;
      ShellMalloc.Free(ItemIdList);
    end;
  end;
end;

procedure TMainForm.EditNumberKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['0'..'9', #8, #127]) and not ((Key = '-') and (TEdit(Sender).SelStart = 0) and (Pos('-', TEdit(Sender).Text) = 0)) then
    Key := #0;
end;

procedure TMainForm.Log(const Msg: string);
begin
  memoLog.Lines.Add(FormatDateTime('[yyyy-mm-dd hh:nn:ss] ', Now) + Msg);
  SendMessage(memoLog.Handle, EM_SCROLLCARET, 0, 0);
end;

end.
