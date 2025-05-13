library TaskLibrary;

uses
  Windows,
  SysUtils,
  Classes,
  Dialogs;

type
  // ���� ���������� �����
  TParamType = (ptString, ptInteger, ptBoolean, ptStringList);

  // �������� ��������� ������
  TTaskParamInfo = record
    Name: array[0..255] of Char;
    Description: array[0..1023] of Char;
    ParamType: TParamType;
    Required: Boolean;
  end;

  // �������� ���������
  TTaskParamValue = record
    ParamType: TParamType;
    case Integer of
      0:
        (StringValue: array[0..1023] of Char);
      1:
        (IntValue: Integer);
      2:
        (BoolValue: Boolean);
      3:
        (StringListPtr: Pointer); // ��������� �� TStringList
  end;

  // ���������� � ������
  TTaskInfo = record
    Name: array[0..255] of Char;
    Description: array[0..1023] of Char;
    ParamCount: Integer;
  end;

  // �������� ���������� ������
  TTaskProgress = record
    TaskId: Integer;
    Progress: Integer; // 0-100
    Status: Integer;   // 0 - � ��������, 1 - �����������, 2 - ���������, 3 - ������
    ErrorMessage: array[0..1023] of Char;
  end;

  // ��������� ���������� ������
  TTaskResult = record
    TaskId: Integer;
    Success: Boolean;
    Message: array[0..1023] of Char;
    ResultCount: Integer;
    ResultData: Pointer; // ��������� �� �������������� ������ ����������
  end;

  TParamArray = array of PChar;

  TParamValues = array of TTaskParamValue;

  // ����� ��� ���������� ������ � ��������� ������
  TTaskThread = class(TThread)
  private
    FTaskId: Integer;
    FTaskName: string;
    FProgress: Integer;
    FStatus: Integer;
    FErrorMessage: string;
    FResults: TStringList;
    FSuccess: Boolean;
    FResultMessage: string;
    FParameters: TStringList; // ��������� ��������� ��� ����������
  protected
    procedure Execute; override;
  public
    constructor Create(ATaskId: Integer; const ATaskName: string);
    destructor Destroy; override;
    property TaskId: Integer read FTaskId;
    property Progress: Integer read FProgress;
    property Status: Integer read FStatus;
    property ErrorMessage: string read FErrorMessage;
    property Results: TStringList read FResults;
    property Success: Boolean read FSuccess;
    property ResultMessage: string read FResultMessage;
    property Parameters: TStringList read FParameters write FParameters;
  end;

  // �������� �����
  TTaskManager = class
  private
    FTasks: TList;
    FCriticalSection: TRTLCriticalSection;
  public
    constructor Create;
    destructor Destroy; override;
    function AddTask(const TaskName: string; const pTaskId: Integer; ParamCount: Integer = 0; ParamNames: TParamArray = nil; ParamValues: TParamValues = nil): Integer;
    function GetTask(TaskId: Integer): TTaskThread;
    function GetTaskProgress(TaskId: Integer; var Progress: TTaskProgress): Boolean;
    function GetTaskResult(TaskId: Integer; var TaskResult: TTaskResult): Boolean;
    function GetTaskResultDetail(TaskId, Index: Integer; Buffer: PChar; BufSize: Integer): Boolean;
    function StopTask(TaskId: Integer): Boolean;
    function FreeTask(TaskId: Integer): Boolean;
  end;

var
  TaskManager: TTaskManager = nil;

// ���������� ������ TTaskThread
constructor TTaskThread.Create(ATaskId: Integer; const ATaskName: string);
begin
  inherited Create(True);
  FTaskId := ATaskId;
  FTaskName := ATaskName;
  FProgress := 0;
  FStatus := 0; // ��������
  FErrorMessage := '';
  FResults := TStringList.Create;
  FSuccess := False;
  FResultMessage := '';
  FParameters := TStringList.Create; // �������������� ������ ����������
  FreeOnTerminate := False;
end;

destructor TTaskThread.Destroy;
begin
  FParameters.Free; // ����������� ������ ��� ������ ����������
  FResults.Free;
  inherited Destroy;
end;

procedure TTaskThread.Execute;
var
  i: Integer;
  StartTime: Cardinal;
  ElapsedTime: Cardinal;
  TotalTime: Cardinal;
  DelayParam, CommentParam: string;
  ShowDetailsParam: string;
  CustomDelay: Integer;
  ShowDetails: Boolean;
begin
  try
    FStatus := 1; // �����������
    // ��������� ��������� ���������� ��� ������������
    FResults.Add('������ ���������� ������: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    FResults.Add('��� ������: ' + FTaskName);

    // �������� �������� ����������
    DelayParam := FParameters.Values['Delay'];
//    showmessage('DelayParam='+DelayParam);
    ShowDetailsParam := FParameters.Values['ShowDetails'];
    CommentParam := FParameters.Values['Comment'];

    // ���������, ���� �� �������� ��� ������� ����������
    CustomDelay := 20; // �� ��������� 20 ������
    if DelayParam <> '' then
    begin
      try
        CustomDelay := StrToInt(DelayParam);
        if CustomDelay <= 0 then
          CustomDelay := 1;
        if CustomDelay > 300 then
          CustomDelay := 300; // ������������ ������������ ����� 5 ��������
      except
        // ���� ������ ��������������, ���������� �������� �� ���������
      end;
    end;

    // ����������, ����� �� ���������� ������
    ShowDetails := False;
    if ShowDetailsParam <> '' then
    begin
      ShowDetails := (ShowDetailsParam = 'True');
    end;

    // ��������� ���������� � ����������
    FResults.Add('');
    FResults.Add('��������� ������:');
    FResults.Add(Format('- ����� ����������: %d ������', [CustomDelay]));
    FResults.Add(Format('- ���������� ������: %s', [BoolToStr(ShowDetails, True)]));

    if CommentParam <> '' then
    begin
      FResults.Add(Format('- �����������: %s', [CommentParam]));
    end
    else
    begin
      FResults.Add('- �����������: �� ������');
    end;

    // ��� �������� ������ ������ ��������� ������ � ������� ���������� �������
    StartTime := GetTickCount;
    TotalTime := CustomDelay * 1000; // ������������ � ������������

    while not Terminated and (FProgress < 100) do
    begin
      Sleep(100); // ��������� ��������

      ElapsedTime := GetTickCount - StartTime;
      if ElapsedTime >= TotalTime then
        FProgress := 100
      else
        FProgress := Round((ElapsedTime / TotalTime) * 100);

      // ��������� ���������� � ���������
      if ShowDetails then
      begin
        // ���� �������� ������, ��������� ���������� � ������ 5%
        if (FProgress mod 5 = 0) and (FProgress > 0) and (FProgress < 100) then
          FResults.Add(Format('�������� ����������: %d%% (������ %d ������)', [FProgress, ElapsedTime div 1000]));
      end
      else
      begin
        // ����� ������ � ������ 25%
        if (FProgress mod 25 = 0) and (FProgress > 0) and (FProgress < 100) then
          FResults.Add(Format('�������� ����������: %d%% (������ %d ������)', [FProgress, ElapsedTime div 1000]));
      end;
    end;

    if Terminated then
    begin
      FStatus := 3; // ������
      FErrorMessage := '������ ���� ������������� �����������';
      FSuccess := False;
      FResultMessage := '������ �� ���������';
      FResults.Add('������ ���� �������� �������������');
    end
    else
    begin
      FStatus := 2; // ���������
      FSuccess := True;
      FResultMessage := '������ ������� ���������';
      FResults.Add('������ ������� ���������: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
      FResults.Add(Format('����� ����� ����������: %d ������', [ElapsedTime div 1000]));

      // ��������� ����������� � ���������, ���� �� ��� ������
      if CommentParam <> '' then
      begin
        FResults.Add('');
        FResults.Add('����������� ������������: ' + CommentParam);
      end;
    end;
  except
    on E: Exception do
    begin
      FStatus := 3; // ������
      FErrorMessage := E.Message;
      FSuccess := False;
      FResultMessage := '������ ��� ���������� ������: ' + E.Message;
      FResults.Add('��������� ������: ' + E.Message);
    end;
  end;
end;

// ���������� ������ TTaskManager
constructor TTaskManager.Create;
begin
  inherited Create;
  FTasks := TList.Create;
  InitializeCriticalSection(FCriticalSection);
end;

destructor TTaskManager.Destroy;
var
  i: Integer;
  Task: TTaskThread;
begin
  EnterCriticalSection(FCriticalSection);
  try
    for i := 0 to FTasks.Count - 1 do
    begin
      Task := TTaskThread(FTasks[i]);
      if Task.Status = 1 then // ���� ������ �����������
        Task.Terminate;
      Task.WaitFor;
      Task.Free;
    end;
    FTasks.Free;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
  DeleteCriticalSection(FCriticalSection);
  inherited Destroy;
end;

function TTaskManager.AddTask(const TaskName: string; const pTaskId: Integer; ParamCount: Integer = 0; ParamNames: TParamArray = nil; ParamValues: TParamValues = nil): Integer;
var
  Task: TTaskThread;
  TaskId: Integer;
  i: Integer;
  ParamName: string;
  ParamValue: string;
begin
  EnterCriticalSection(FCriticalSection);
  try
    // ���������� ���������� ID ��� ������
    TaskId := pTaskId;

    // ������� � ��������� ����� ������
    Task := TTaskThread.Create(TaskId, TaskName);

    // ��������� ��������� � ������
    if ParamCount > 0 then
    begin
      for i := 0 to ParamCount - 1 do
      begin
        ParamName := ParamNames[i];

        case ParamValues[i].ParamType of
          ptString:
            ParamValue := ParamValues[i].StringValue;
          ptInteger:
            ParamValue := IntToStr(ParamValues[i].IntValue);
          ptBoolean:
            ParamValue := BoolToStr(ParamValues[i].BoolValue, True);
          ptStringList:
            begin
              if ParamValues[i].StringListPtr <> nil then
                ParamValue := TStringList(ParamValues[i].StringListPtr).Text
              else
                ParamValue := '(������ ������)';
            end;
        end;

        // ��������� �������� � ������ ���������� ������
        Task.Parameters.Values[ParamName] := ParamValue;
      end;
    end;

    FTasks.Add(Task);
    Task.Resume; // ��������� �����

    Result := TaskId;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
end;

function TTaskManager.GetTask(TaskId: Integer): TTaskThread;
var
  i: Integer;
  Task: TTaskThread;
begin
  Result := nil;

  EnterCriticalSection(FCriticalSection);
  try
    for i := 0 to FTasks.Count - 1 do
    begin
      Task := TTaskThread(FTasks[i]);
      if Task.TaskId = TaskId then
      begin
        Result := Task;
        Break;
      end;
    end;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
end;

function TTaskManager.GetTaskProgress(TaskId: Integer; var Progress: TTaskProgress): Boolean;
var
  Task: TTaskThread;
begin
  Result := False;

  EnterCriticalSection(FCriticalSection);
  try
    Task := GetTask(TaskId);
    if Task <> nil then
    begin
      Progress.TaskId := Task.TaskId;
      Progress.Progress := Task.Progress;
      Progress.Status := Task.Status;
      StrPCopy(Progress.ErrorMessage, Task.ErrorMessage);
      Result := True;
    end;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
end;

function TTaskManager.GetTaskResult(TaskId: Integer; var TaskResult: TTaskResult): Boolean;
var
  Task: TTaskThread;
begin
  Result := False;

  EnterCriticalSection(FCriticalSection);
  try
    Task := GetTask(TaskId);
    if Task <> nil then
    begin
      TaskResult.TaskId := Task.TaskId;
      TaskResult.Success := Task.Success;
      StrPCopy(TaskResult.Message, Task.ResultMessage);
      TaskResult.ResultCount := Task.Results.Count;
      TaskResult.ResultData := Task.Results; // �������� ��������� �� ������ �����������
      Result := True;
    end;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
end;

function TTaskManager.GetTaskResultDetail(TaskId, Index: Integer; Buffer: PChar; BufSize: Integer): Boolean;
var
  Task: TTaskThread;
begin
  Result := False;

  EnterCriticalSection(FCriticalSection);
  try
    Task := GetTask(TaskId);
    if (Task <> nil) and (Index >= 0) and (Index < Task.Results.Count) then
    begin
      StrLCopy(Buffer, PChar(Task.Results[Index]), BufSize - 1);
      Result := True;
    end;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
end;

function TTaskManager.StopTask(TaskId: Integer): Boolean;
var
  Task: TTaskThread;
begin
  Result := False;

  EnterCriticalSection(FCriticalSection);
  try
    Task := GetTask(TaskId);
    if (Task <> nil) and (Task.Status = 1) then // ���� ������ �����������
    begin
      Task.Terminate;
      Result := True;
    end;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
end;

function TTaskManager.FreeTask(TaskId: Integer): Boolean;
var
  i: Integer;
  Task: TTaskThread;
begin
  Result := False;

  EnterCriticalSection(FCriticalSection);
  try
    for i := 0 to FTasks.Count - 1 do
    begin
      Task := TTaskThread(FTasks[i]);
      if Task.TaskId = TaskId then
      begin
        // ���� ������ ��� ��� �����������, ������������� ��
        if Task.Status = 1 then
          Task.Terminate;

        // ���� ���������� ������
        Task.WaitFor;

        // ������� ������ �� ������ � ����������� �������
        FTasks.Delete(i);
        Task.Free;

        Result := True;
        Break;
      end;
    end;
  finally
    LeaveCriticalSection(FCriticalSection);
  end;
end;

// �������������� ������� DLL

// ���������� ���������� ��������� �����
function GetTaskCount: Integer; stdcall;
begin
  // � �������� DLL � ��� ������ ���� ������
  Result := 1;
end;

// ���������� ���������� � ������ �� �������
function GetTaskInfo(Index: Integer; var TaskInfo: TTaskInfo): Boolean; stdcall;
begin
  Result := False;

  if Index = 0 then
  begin
    StrPCopy(TaskInfo.Name, 'TestTask');
    StrPCopy(TaskInfo.Description, '�������� ������, ������� ������ ��������� �������� �� 0% �� 100% � ������� ���������� �������');
    TaskInfo.ParamCount := 3; // ������ � �������� ������ ���� 3 ���������
    Result := True;
  end;
end;

// ���������� ���������� � ��������� ������
function GetTaskParamInfo(TaskIndex, ParamIndex: Integer; var ParamInfo: TTaskParamInfo): Boolean; stdcall;
begin
  Result := False;

  // ��������� ������ ������
  if TaskIndex <> 0 then
    Exit;

  // ���������� ��������� ��� �������� ������
  case ParamIndex of
    0: // �������� "Delay"
      begin
        StrPCopy(ParamInfo.Name, 'Delay');
        StrPCopy(ParamInfo.Description, '����� ���������� ������ � ��������');
        ParamInfo.ParamType := ptInteger;
        ParamInfo.Required := False;
        Result := True;
      end;
    1: // �������� "ShowDetails"
      begin
        StrPCopy(ParamInfo.Name, 'ShowDetails');
        StrPCopy(ParamInfo.Description, '���������� ��������� ���������� � ����������');
        ParamInfo.ParamType := ptBoolean;
        ParamInfo.Required := False;
        Result := True;
      end;
    2: // �������� "Comment"
      begin
        StrPCopy(ParamInfo.Name, 'Comment');
        StrPCopy(ParamInfo.Description, '����������� � ������');
        ParamInfo.ParamType := ptString;
        ParamInfo.Required := False;
        Result := True;
      end;
  end;
end;

// ��������� ������ � ���������� � ID
function StartTask(TaskName: PChar; pTaskId: Integer; ParamCount: Integer; ParamNames, ParamValues: Pointer): Boolean; stdcall;
var
  i: Integer;
  ParamNamesArray: TParamArray;
  ParamValuesArray: TParamValues;
  ii: Integer;
  ss: string;
begin
  Result := False;

  // ���������, ������ �� �������� �����
  if TaskManager = nil then
    Exit;

  // ���������, �������������� �� ������
  if CompareText(TaskName, 'TestTask') = 0 then
  begin
    // �������������� ������� ��� ����������
    if ParamCount > 0 then
    begin
      SetLength(ParamNamesArray, ParamCount);
      SetLength(ParamValuesArray, ParamCount);

      // �������� ��������� �� ����� � �������� ����������
      Move(ParamNames^, ParamNamesArray[0], ParamCount * SizeOf(PChar));
      Move(ParamValues^, ParamValuesArray[0], ParamCount * SizeOf(TTaskParamValue));


//      ShowMessage('ss:');
//      for ii := 1 to ParamCount do
//        ss:=ss+ParamNamesArray[ii]+'='+IntToStr(ParamValuesArray[ii].IntValue)+'; ';
//      ShowMessage(ss);
      // ��������� ������ � �����������
      TaskManager.AddTask(TaskName, pTaskId, ParamCount, ParamNamesArray, ParamValuesArray);

    end
    else
    begin
      // ��������� ������ ��� ����������
      //ShowMessage('No params');
      TaskManager.AddTask(TaskName, pTaskId);
    end;

    Result := True;
  end;
end;

// ���������� ������� �������� ���������� ������
function GetTaskProgress(TaskId: Integer; var Progress: TTaskProgress): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.GetTaskProgress(TaskId, Progress);
end;

// ���������� ��������� ���������� ������
function GetTaskResult(TaskId: Integer; var TaskResult: TTaskResult): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.GetTaskResult(TaskId, TaskResult);
end;

// ���������� ��������� ���������� � ���������� ������
function GetTaskResultDetail(TaskId, Index: Integer; Buffer: PChar; BufSize: Integer): Boolean; stdcall;
begin
  Result := False;
  if TaskManager <> nil then
    Result := TaskManager.GetTaskResultDetail(TaskId, Index, Buffer, BufSize);
end;

// ������������� ���������� ������
function StopTask(TaskId: Integer): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.StopTask(TaskId);
end;

// ����������� �������, ���������� ��� ������
function FreeTask(TaskId: Integer): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.FreeTask(TaskId);
end;

// ������������� � ����������� DLL
procedure DLLEntryPoint(dwReason: DWORD);
begin
  case dwReason of
    DLL_PROCESS_ATTACH:
      begin
        // ������� �������� ����� ��� �������� DLL
        TaskManager := TTaskManager.Create;
      end;
    DLL_PROCESS_DETACH:
      begin
        // ����������� �������� ����� ��� �������� DLL
        if TaskManager <> nil then
        begin
          TaskManager.Free;
          TaskManager := nil;
        end;
      end;
  end;
end;

exports
  GetTaskCount,
  GetTaskInfo,
  GetTaskParamInfo,
  StartTask,
  GetTaskProgress,
  GetTaskResult,
  GetTaskResultDetail,
  StopTask,
  FreeTask;

begin
  DLLProc := @DLLEntryPoint;
  DLLEntryPoint(DLL_PROCESS_ATTACH);
end.

