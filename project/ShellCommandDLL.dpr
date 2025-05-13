library ShellCommandDLL;

uses
  Windows,
  SysUtils,
  Classes,
  ShellAPI,
  Math,
  DateUtils;

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
      0: (StringValue: array[0..1023] of Char);
      1: (IntValue: Integer);
      2: (BoolValue: Boolean);
      3: (StringListPtr: Pointer); // ��������� �� TStringList
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

  // ���������� � ���������� ��������
  TProcessInfo = record
    hProcess: THandle;         // ���������� ��������
    hThread: THandle;          // ���������� ��������� ������
    dwProcessId: DWORD;        // ID ��������
    hStdOutRead: THandle;      // ����� ��� ������ StdOut
    hStdOutWrite: THandle;     // ����� ��� ������ StdOut
    hStdErrRead: THandle;      // ����� ��� ������ StdErr
    hStdErrWrite: THandle;     // ����� ��� ������ StdErr
    ExitCode: DWORD;           // ��� ���������� ��������
    StartTime: TDateTime;      // ����� ������� ��������
  end;

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
    FParameters: TStringList;
    
    // ���� ��� ������ � ���������
    FProcessInfo: TProcessInfo;
    FCommandLine: string;
    FWorkingDir: string;
    FShowWindow: Boolean;
    FTimeout: Integer;
    
    function ExecuteShellCommand(const CommandLine, WorkingDir: string; ShowWindow: Boolean): Boolean;
    procedure ReadProcessOutput;
    function IsProcessRunning: Boolean;
    procedure TerminateProcess;
    procedure CleanupProcess;
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
    function AddTask(const TaskName: string; pTaskId: Integer; ParamCount: Integer = 0; ParamNames: TParamArray = nil; ParamValues: TParamValues = nil): Integer;
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
  FParameters := TStringList.Create;
  FreeOnTerminate := False;
  
  // ������������� ����� ��� ��������
  FProcessInfo.hProcess := 0;
  FProcessInfo.hThread := 0;
  FProcessInfo.hStdOutRead := 0;
  FProcessInfo.hStdOutWrite := 0;
  FProcessInfo.hStdErrRead := 0;
  FProcessInfo.hStdErrWrite := 0;
end;

destructor TTaskThread.Destroy;
begin
  // ����������, ��� ������� �������� � ������� �����������
  CleanupProcess;
  
  FParameters.Free;
  FResults.Free;
  inherited Destroy;
end;

// ������� ������� �������� � ���������������� ������
function TTaskThread.ExecuteShellCommand(const CommandLine, WorkingDir: string; ShowWindow: Boolean): Boolean;
var
  StartupInfo: TStartupInfo;
  Security: TSecurityAttributes;
  ProcessCreated: Boolean;
  CommandLineStr: array[0..4095] of Char;
  WindowMode: DWORD;
  PI: TProcessInformation;
begin
  Result := False;
  
  // ������������� �������� ������������ ��� �������� �������
  FillChar(Security, SizeOf(Security), 0);
  Security.nLength := SizeOf(Security);
  Security.bInheritHandle := True;
  
  // ������� ����� ��� stdout
  if not CreatePipe(FProcessInfo.hStdOutRead, FProcessInfo.hStdOutWrite, @Security, 0) then
  begin
    FErrorMessage := '�� ������� ������� ����� ��� stdout';
    Exit;
  end;
  
  // ������� ����� ��� stderr
  if not CreatePipe(FProcessInfo.hStdErrRead, FProcessInfo.hStdErrWrite, @Security, 0) then
  begin
    CloseHandle(FProcessInfo.hStdOutRead);
    CloseHandle(FProcessInfo.hStdOutWrite);
    FErrorMessage := '�� ������� ������� ����� ��� stderr';
    Exit;
  end;
  
  // ������������� ������������� ������ ��� ������������ ������
  SetHandleInformation(FProcessInfo.hStdOutRead, HANDLE_FLAG_INHERIT, 0);
  SetHandleInformation(FProcessInfo.hStdErrRead, HANDLE_FLAG_INHERIT, 0);
  
  // ������������� ��������� StartupInfo
  FillChar(StartupInfo, SizeOf(StartupInfo), 0);
  StartupInfo.cb := SizeOf(StartupInfo);
  StartupInfo.dwFlags := STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW;
  StartupInfo.hStdInput := GetStdHandle(STD_INPUT_HANDLE);
  StartupInfo.hStdOutput := FProcessInfo.hStdOutWrite;
  StartupInfo.hStdError := FProcessInfo.hStdErrWrite;
  
  if ShowWindow then
    WindowMode := SW_SHOW
  else
    WindowMode := SW_HIDE;
    
  StartupInfo.wShowWindow := WindowMode;

  // �������������� ������ ��������� ������
  StrPCopy(CommandLineStr, CommandLine);
  
  // ������������� ��������� ���������� � ��������
  FillChar(FProcessInfo, SizeOf(TProcessInfo), 0);
  
  // ������� �������
  ProcessCreated := CreateProcess(
    nil,                  // ��� ����������
    CommandLineStr,       // ��������� ������
    nil,                  // �������� ������������ ��������
    nil,                  // �������� ������������ ������
    True,                 // ������������ ������������
    CREATE_NEW_CONSOLE,   // ����� ��������
    nil,                  // ��������� ������������� ��������
    PChar(WorkingDir),    // ������� �������
    StartupInfo,          // ���������� � ������
    PI // ���������� � ��������
  );

  FProcessInfo.hProcess := PI.hProcess;
  FProcessInfo.hThread := PI.hThread;
  FProcessInfo.dwProcessId := PI.dwProcessId;
  
  // ��������� ����������� ������, ��� ��� ��� ������ ������������ �������� ���������
  CloseHandle(FProcessInfo.hStdOutWrite);
  FProcessInfo.hStdOutWrite := 0;
  CloseHandle(FProcessInfo.hStdErrWrite);
  FProcessInfo.hStdErrWrite := 0;
  
  if ProcessCreated then
  begin
    FProcessInfo.StartTime := Now;
    Result := True;
  end
  else
  begin
    FErrorMessage := Format('�� ������� ������� �������: %s (��� ������: %d)', 
                          [SysErrorMessage(GetLastError), GetLastError]);
    CleanupProcess;
  end;
end;

// ��������, ������� �� �������
function TTaskThread.IsProcessRunning: Boolean;
var
  ExitCode: DWORD;
begin
  Result := False;
  
  if FProcessInfo.hProcess <> 0 then
  begin
    // �������� ������ ��������
    GetExitCodeProcess(FProcessInfo.hProcess, ExitCode);
    Result := (ExitCode = STILL_ACTIVE);
    FProcessInfo.ExitCode := ExitCode;
  end;
end;

// �������������� ���������� ��������
procedure TTaskThread.TerminateProcess;
begin
  if FProcessInfo.hProcess <> 0 then
  begin
    // ������������� ��������� �������
    Windows.TerminateProcess(FProcessInfo.hProcess, 1);
  end;
end;

// ������������ �������� ��������
procedure TTaskThread.CleanupProcess;
begin
  // ��������� ��� �������� �����������
  if FProcessInfo.hStdOutRead <> 0 then
  begin
    CloseHandle(FProcessInfo.hStdOutRead);
    FProcessInfo.hStdOutRead := 0;
  end;
  
  if FProcessInfo.hStdOutWrite <> 0 then
  begin
    CloseHandle(FProcessInfo.hStdOutWrite);
    FProcessInfo.hStdOutWrite := 0;
  end;
  
  if FProcessInfo.hStdErrRead <> 0 then
  begin
    CloseHandle(FProcessInfo.hStdErrRead);
    FProcessInfo.hStdErrRead := 0;
  end;
  
  if FProcessInfo.hStdErrWrite <> 0 then
  begin
    CloseHandle(FProcessInfo.hStdErrWrite);
    FProcessInfo.hStdErrWrite := 0;
  end;
  
  if FProcessInfo.hThread <> 0 then
  begin
    CloseHandle(FProcessInfo.hThread);
    FProcessInfo.hThread := 0;
  end;
  
  if FProcessInfo.hProcess <> 0 then
  begin
    CloseHandle(FProcessInfo.hProcess);
    FProcessInfo.hProcess := 0;
  end;
end;

// ������ ������ ��������
procedure TTaskThread.ReadProcessOutput;
const
  BUFFER_SIZE = 1024;
var
  Buffer: array[0..BUFFER_SIZE-1] of Char;
  BytesRead: DWORD;
  Output: string;
  StdOutAvailable, StdErrAvailable: Boolean;
begin
  if FProcessInfo.hStdOutRead = 0 then Exit;
  
  StdOutAvailable := True;
  StdErrAvailable := True;
  
  while (StdOutAvailable or StdErrAvailable) and not Terminated do
  begin
    // ��������� ������� ������ � stdout
    if StdOutAvailable then
    begin
      BytesRead := 0;
      if not PeekNamedPipe(FProcessInfo.hStdOutRead, nil, 0, nil, @BytesRead, nil) or (BytesRead = 0) then
      begin
        StdOutAvailable := False;
      end
      else if BytesRead > 0 then
      begin
        FillChar(Buffer, BUFFER_SIZE, 0);
        if ReadFile(FProcessInfo.hStdOutRead, Buffer, BUFFER_SIZE-1, BytesRead, nil) and (BytesRead > 0) then
        begin
          Buffer[BytesRead] := #0;
          Output := Buffer;
          if Output <> '' then
          begin
            FResults.Add('STDOUT: ' + Output);
          end;
        end;
      end;
    end;
    
    // ��������� ������� ������ � stderr
    if StdErrAvailable then
    begin
      BytesRead := 0;
      if not PeekNamedPipe(FProcessInfo.hStdErrRead, nil, 0, nil, @BytesRead, nil) or (BytesRead = 0) then
      begin
        StdErrAvailable := False;
      end
      else if BytesRead > 0 then
      begin
        FillChar(Buffer, BUFFER_SIZE, 0);
        if ReadFile(FProcessInfo.hStdErrRead, Buffer, BUFFER_SIZE-1, BytesRead, nil) and (BytesRead > 0) then
        begin
          Buffer[BytesRead] := #0;
          Output := Buffer;
          if Output <> '' then
          begin
            FResults.Add('STDERR: ' + Output);
          end;
        end;
      end;
    end;
    
    // ���� ��� ������ � ����� �������, �� ������� ��� ��������, ���� �������
    if (not StdOutAvailable) and (not StdErrAvailable) and IsProcessRunning then
    begin
      Sleep(50);
      StdOutAvailable := True;
      StdErrAvailable := True;
    end;
  end;
end;

procedure TTaskThread.Execute;
var
  StartTime: TDateTime;
  RunningTime: Integer;
  HadTimeout: Boolean;
  ElapsedSeconds: Integer;
begin
  try
    FStatus := 1; // �����������
    FResults.Add('������ ���������� �������: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    
    // �������� ���������
    FCommandLine := FParameters.Values['CommandLine'];
    FWorkingDir := FParameters.Values['WorkingDirectory'];
    FShowWindow := StrToBoolDef(FParameters.Values['ShowWindow'], False);
    FTimeout := StrToIntDef(FParameters.Values['Timeout'], 0); // 0 = ��� ��������
    
    // ��������� ������� ������������ ����������
    if FCommandLine = '' then
    begin
      FStatus := 3; // ������
      FErrorMessage := '�� ������� ������� ��� ����������';
      FSuccess := False;
      FResultMessage := '������: �� ������� �������';
      Exit;
    end;
    
    // ���� ������� ���������� �� �������, ���������� �������
    if FWorkingDir = '' then
      FWorkingDir := GetCurrentDir;
    
    // ������� ���������� � ����������
    FResults.Add('��������� �������:');
    FResults.Add('- ��������� ������: ' + FCommandLine);
    FResults.Add('- ������� ����������: ' + FWorkingDir);
    FResults.Add('- ���������� ����: ' + BoolToStr(FShowWindow, True));
    if FTimeout > 0 then
      FResults.Add('- �������: ' + IntToStr(FTimeout) + ' ������')
    else
      FResults.Add('- �������: �� ����������');
    
    // ��������� �������
    FResults.Add('������ �������...');
    if not ExecuteShellCommand(FCommandLine, FWorkingDir, FShowWindow) then
    begin
      FStatus := 3; // ������
      FSuccess := False;
      FResultMessage := '������ ������� �������: ' + FErrorMessage;
      Exit;
    end;
    
    // ������� ������� �������
    FResults.Add('������� �������� �������. PID: ' + IntToStr(FProcessInfo.dwProcessId));
    FProgress := 1; // �������� � 1%
    
    // ���������� ����� ������
    StartTime := Now;
    HadTimeout := False;
    
    // �������� ���� ����������� ��������
    while IsProcessRunning and not Terminated do
    begin
      // ������ ����� ��������
      ReadProcessOutput;
      
      // ��������� �������� �� ������ ������� ����������
      if FTimeout > 0 then
      begin
        // ���� ���������� �������, ������� ���������� ������� �� ����������� ���������� ������� � ��������
        RunningTime := SecondsBetween(Now, StartTime);
        FProgress := Min(99, Trunc((RunningTime / FTimeout) * 100));
        
        // ���������, �� ����� �� �������
        if RunningTime >= FTimeout then
        begin
          FResults.Add('�������� ������� ���������� ������� (' + IntToStr(FTimeout) + ' ������)');
          TerminateProcess;
          HadTimeout := True;
          Break;
        end;
      end
      else
      begin
        // ���� ������� �� ����������, ���������� �������� �������� 
        // �� ������ ������������ ����������, �� �� ����� 99%
        ElapsedSeconds := SecondsBetween(Now, StartTime);
        if ElapsedSeconds <= 60 then
          FProgress := Min(50, ElapsedSeconds) // ������ 60 ������ - �� 50%
        else
          FProgress := Min(99, 50 + Trunc(((ElapsedSeconds - 60) / 600) * 49)); // ��� 600 ������ �� 99%
      end;
      
      // ��������� ���������� � ������� ���������� ������ 5 ������
      if (SecondsBetween(Now, StartTime) mod 5 = 0) then
      begin
        FResults.Add('����� ����������: ' + IntToStr(SecondsBetween(Now, StartTime)) + ' ������');
      end;
      
      Sleep(100); // ��������� �������� ��� ���������� ��������
    end;
    
    // ��������� ������ ������
    ReadProcessOutput;
    
    // ��������� ��������� ����������
    if Terminated then
    begin
      // ������ ���� �������� �������������
      FResults.Add('���������� ������� �������� �������������');
      TerminateProcess;
      FStatus := 3; // ������
      FErrorMessage := '������ ���� ������������� �����������';
      FSuccess := False;
      FResultMessage := '������ �� ���������';
    end
    else if HadTimeout then
    begin
      // ������� ��� ���������� ��-�� ��������
      FStatus := 3; // ������
      FErrorMessage := '�������� ������� ����������';
      FSuccess := False;
      FResultMessage := '������ �� ��������� ��-�� ��������';
    end
    else
    begin
      // ������� ���������� ���
      FStatus := 2; // ���������
      FProgress := 100; // 100%
      
      if FProcessInfo.ExitCode = 0 then
      begin
        FSuccess := True;
        FResultMessage := '������� ������� ���������';
        FResults.Add('������� ������� ��������� � ����� 0');
      end
      else
      begin
        FSuccess := False;
        FResultMessage := '������� ��������� � �������, ���: ' + IntToStr(FProcessInfo.ExitCode);
        FResults.Add('������� ��������� � ����� ������: ' + IntToStr(FProcessInfo.ExitCode));
      end;
    end;
    
    // ��������� ���������� � ������� ����������
    FResults.Add('����� ����� ����������: ' + IntToStr(SecondsBetween(Now, StartTime)) + ' ������');
    FResults.Add('���������� ������: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    
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
  
  // ����������� ������� ��������
  CleanupProcess;
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

function TTaskManager.AddTask(const TaskName: string; pTaskId: Integer; ParamCount: Integer = 0; ParamNames: TParamArray = nil; ParamValues: TParamValues = nil): Integer;
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

    // ������� ����� ������
    Task := TTaskThread.Create(TaskId, TaskName);

    // ��������� ��� ������ ��� ��������
    Task.Parameters.Values['TaskType'] := TaskName;

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
                ParamValue := TStringList(ParamValues[i].StringListPtr).CommaText
              else
                ParamValue := '';
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
  // � ����� DLL ������ ���� ��� ������ - ���������� shell-�������
  Result := 1;
end;

// ���������� ���������� � ������ �� �������
function GetTaskInfo(Index: Integer; var TaskInfo: TTaskInfo): Boolean; stdcall;
begin
  Result := False;

  if Index = 0 then // ExecuteShellCommand
  begin
    StrPCopy(TaskInfo.Name, 'ExecuteShellCommand');
    StrPCopy(TaskInfo.Description, '���������� shell-������� � ������������� ���������');
    TaskInfo.ParamCount := 4; // CommandLine, WorkingDirectory, ShowWindow, Timeout
    Result := True;
  end;
end;

// ���������� ���������� � ��������� ������
function GetTaskParamInfo(TaskIndex, ParamIndex: Integer; var ParamInfo: TTaskParamInfo): Boolean; stdcall;
begin
  Result := False;

  if TaskIndex <> 0 then
    Exit;

  case ParamIndex of
    0: // CommandLine
      begin
        StrPCopy(ParamInfo.Name, 'CommandLine');
        StrPCopy(ParamInfo.Description, '��������� ������ ��� ����������');
        ParamInfo.ParamType := ptString;
        ParamInfo.Required := True;
        Result := True;
      end;
    1: // WorkingDirectory
      begin
        StrPCopy(ParamInfo.Name, 'WorkingDirectory');
        StrPCopy(ParamInfo.Description, '������� ���������� ��� ���������� �������');
        ParamInfo.ParamType := ptString;
        ParamInfo.Required := False;
        Result := True;
      end;
    2: // ShowWindow
      begin
        StrPCopy(ParamInfo.Name, 'ShowWindow');
        StrPCopy(ParamInfo.Description, '���������� ���� ��������');
        ParamInfo.ParamType := ptBoolean;
        ParamInfo.Required := False;
        Result := True;
      end;
    3: // Timeout
      begin
        StrPCopy(ParamInfo.Name, 'Timeout');
        StrPCopy(ParamInfo.Description, '������� ���������� ������� � �������� (0 - ��� ��������)');
        ParamInfo.ParamType := ptInteger;
        ParamInfo.Required := False;
        Result := True;
      end;
  end;
end;

// ��������� ������ � ���������� � ID
function StartTask(TaskName: PChar; TaskId: Integer; ParamCount: Integer; ParamNames, ParamValues: Pointer): Boolean; stdcall;
var
  ParamNamesArray: TParamArray;
  ParamValuesArray: TParamValues;
begin
  Result := False;

  // ���������, ������ �� �������� �����
  if TaskManager = nil then
    Exit;

  // ���������, �������������� �� ������
  if CompareText(TaskName, 'ExecuteShellCommand') <> 0 then
    Exit;

  // �������������� ������� ��� ����������
  if ParamCount > 0 then
  begin
    SetLength(ParamNamesArray, ParamCount);
    SetLength(ParamValuesArray, ParamCount);

    // �������� ��������� �� ����� � �������� ����������
    Move(ParamNames^, ParamNamesArray[0], ParamCount * SizeOf(PChar));
    Move(ParamValues^, ParamValuesArray[0], ParamCount * SizeOf(TTaskParamValue));

    // ��������� ������ � �����������
    TaskManager.AddTask(TaskName, TaskId, ParamCount, ParamNamesArray, ParamValuesArray);
    Result := True;
  end
  else
  begin
    // ��� ���������� ������ ��������� ���� ������
    Result := False;
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