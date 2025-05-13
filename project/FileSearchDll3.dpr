library FileSearchDLL3;

uses
  Windows,
  SysUtils,
  Classes,
  Masks,
  Math,
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
    
    // ��� ������ ������
    procedure FindFiles(const StartPath, FileMask: string);
    procedure FindFilesWithMultiMasks(const StartPath: string; Masks: TStringList);
    
    // ��� ������ ����� � �����
    procedure SearchStringInFile(const FilePath, SearchStr: string);
    procedure SearchMultipleStringsInFile(const FilePath: string; SearchStrings: TStringList);
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
  FParameters := TStringList.Create;
  FreeOnTerminate := False;
end;

destructor TTaskThread.Destroy;
begin
  // ������������� ���������, ��� ������ ���������
  if Self.Status = 1 then // ���� ������ ��� �����������
  begin
    Self.Terminate;
    // ���� ������� ������� �� ���������� ����������
    Sleep(1000);
  end;

  try
    FParameters.Free;
    FResults.Free;
  finally
    // ���������� ���������� ��� ������������ ��������
    //showmessage('(DEBUG): ���������� ��� ������������ ��������');
  end;
  inherited Destroy;
end;

// ����� ������ �� ����� �����
procedure TTaskThread.FindFiles(const StartPath, FileMask: string);
var
  SearchRec: TSearchRec;
  CurrentPath: string;
  FileCount: Integer;
  TotalFiles: Integer;
  DirList: TStringList;
  i: Integer;
  IncludeSubfolders: Boolean;
begin
  FileCount := 0;
  TotalFiles := 0;
  
  // ���������, ����� �� ������ � ���������
  IncludeSubfolders := StrToBoolDef(FParameters.Values['IncludeSubfolders'], True);
  
  FResults.Add('������ ������ ������ �� �����: ' + FileMask);
  FResults.Add('��������� �����: ' + StartPath);
  FResults.Add('�������� ��������: ' + BoolToStr(IncludeSubfolders, True));
  
  if not DirectoryExists(StartPath) then
  begin
    FResults.Add('������: ��������� ����� �� ����������');
    Exit;
  end;
  
  // ������� ������ ��� �������� �����
  DirList := TStringList.Create;
  try
    // ��������� ��������� �����
    DirList.Add(IncludeTrailingPathDelimiter(StartPath));
    
    // ���� ������ ����� �� ���� � ������ �� ��������
    i := 0;
    while (i < DirList.Count) and not Terminated do
    begin
      CurrentPath := DirList[i];
      
      // ���� ����� � ������� �����
      if FindFirst(CurrentPath + FileMask, faAnyFile - faDirectory, SearchRec) = 0 then
      begin
        try
          repeat
            Inc(FileCount);
            Inc(TotalFiles);
            // ��������� ���� � ����� � ����������
            FResults.Add('������ ����: ' + CurrentPath + SearchRec.Name);
            
            // ��������� �������� ������ 10 ������
            if (FileCount mod 10 = 0) then
            begin
              FProgress := Min(99, FProgress + 1); // �� ������� �� 100%, ���� �� ��������
              FResults.Add(Format('������� ������: %d', [TotalFiles]));
            end;
            
            if Terminated then Break;
          until FindNext(SearchRec) <> 0;
        finally
          FindClose(SearchRec);
        end;
      end;
      
      // ���� ����� ������ � ���������
      if IncludeSubfolders and not Terminated then
      begin
        // ������� ��� ��������
        if FindFirst(CurrentPath + '*', faDirectory, SearchRec) = 0 then
        begin
          try
            repeat
              // ���������� . � ..
              if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') and
                 ((SearchRec.Attr and faDirectory) = faDirectory) then
              begin
                // ��������� ���� � ����� � ������ ��� ���������
                DirList.Add(IncludeTrailingPathDelimiter(CurrentPath + SearchRec.Name));
              end;
              
              if Terminated then Break;
            until FindNext(SearchRec) <> 0;
          finally
            FindClose(SearchRec);
          end;
        end;
      end;
      
      Inc(i);
    end;
    
    if Terminated then
    begin
      FResults.Add('����� ��� ������� �������������');
    end
    else
    begin
      FResults.Add(Format('������� ����� ������: %d', [TotalFiles]));
    end;
    
  finally
    DirList.Free;
  end;
end;

// ����� ������ �� ���������� ������
procedure TTaskThread.FindFilesWithMultiMasks(const StartPath: string; Masks: TStringList);
var
  i: Integer;
  TotalFilesFound: Integer;
  FilePaths: TStringList;
  CurrentMask: string;
  SearchRec: TSearchRec;
  CurrentPath: string;
  DirList: TStringList;
  j: Integer;
  IncludeSubfolders: Boolean;
begin
  TotalFilesFound := 0;
  
  // ���������, ����� �� ������ � ���������
  IncludeSubfolders := StrToBoolDef(FParameters.Values['IncludeSubfolders'], True);
  
  FResults.Add('������ ������ ������ �� ���������� ������:');
  for i := 0 to Masks.Count - 1 do
  begin
    FResults.Add('- ' + Masks[i]);
  end;
  FResults.Add('��������� �����: ' + StartPath);
  FResults.Add('�������� ��������: ' + BoolToStr(IncludeSubfolders, True));
  
  if not DirectoryExists(StartPath) then
  begin
    FResults.Add('������: ��������� ����� �� ����������');
    Exit;
  end;
  
  // ������� ������ ��� �������� ��������� ������ � ����� ��� ������
  FilePaths := TStringList.Create;
  DirList := TStringList.Create;
  try
    // ��������� ��������� �����
    DirList.Add(IncludeTrailingPathDelimiter(StartPath));
    
    // ���� ������ ����� �� ���� � ������ �� ��������
    j := 0;
    while (j < DirList.Count) and not Terminated do
    begin
      CurrentPath := DirList[j];
      
      // ��������� ��� ����� � ������� �����
      for i := 0 to Masks.Count - 1 do
      begin
        if Terminated then Break;
        
        CurrentMask := Masks[i];
        if FindFirst(CurrentPath + CurrentMask, faAnyFile - faDirectory, SearchRec) = 0 then
        begin
          try
            repeat
              // ���������, �� �������� �� ��� ���� (����� �������� ����������)
              if FilePaths.IndexOf(CurrentPath + SearchRec.Name) < 0 then
              begin
                FilePaths.Add(CurrentPath + SearchRec.Name);
                Inc(TotalFilesFound);
                FResults.Add('������ ����: ' + CurrentPath + SearchRec.Name + ' (�� �����: ' + CurrentMask + ')');
                
                // ��������� ��������
                if (TotalFilesFound mod 10 = 0) then
                begin
                  FProgress := Min(99, FProgress + 1);
                  FResults.Add(Format('������� ������: %d', [TotalFilesFound]));
                end;
              end;
              
              if Terminated then Break;
            until FindNext(SearchRec) <> 0;
          finally
            FindClose(SearchRec);
          end;
        end;
      end;
      
      // ���� ����� ������ � ���������
      if IncludeSubfolders and not Terminated then
      begin
        // ������� ��� ��������
        if FindFirst(CurrentPath + '*', faDirectory, SearchRec) = 0 then
        begin
          try
            repeat
              // ���������� . � ..
              if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') and
                 ((SearchRec.Attr and faDirectory) = faDirectory) then
              begin
                // ��������� ���� � ����� � ������ ��� ���������
                DirList.Add(IncludeTrailingPathDelimiter(CurrentPath + SearchRec.Name));
              end;
              
              if Terminated then Break;
            until FindNext(SearchRec) <> 0;
          finally
            FindClose(SearchRec);
          end;
        end;
      end;
      
      Inc(j);
    end;
    
    if Terminated then
    begin
      FResults.Add('����� ��� ������� �������������');
    end
    else
    begin
      FResults.Add(Format('������� ����� ������: %d', [TotalFilesFound]));
      // ��������� ��� ���� � ������ � ����������
      FResults.Add('');
      FResults.Add('������ ��������� ������:');
      for i := 0 to FilePaths.Count - 1 do
      begin
        FResults.Add(FilePaths[i]);
      end;
    end;
    
  finally
    FilePaths.Free;
    DirList.Free;
  end;
end;

// ����� ����� ������ � �����
procedure TTaskThread.SearchStringInFile(const FilePath, SearchStr: string);
var
  F: file;
  Buffer: array[0..4095] of Byte;
  BytesRead: Integer;
  TotalBytesRead: Int64;
  FileSize: Int64;
  i, j: Integer;
  FoundCount: Integer;
  Match: Boolean;
  SearchBytes: array of Byte;
  Position: Int64;
begin
  FoundCount := 0;
  TotalBytesRead := 0;
  
  FResults.Add('������ ������ ������ � �����:');
  FResults.Add('����: ' + FilePath);
  FResults.Add('������� ������: ' + SearchStr);
  
  if not FileExists(FilePath) then
  begin
    FResults.Add('������: ��������� ���� �� ����������');
    Exit;
  end;
  
  // ����������� ������ ������ � ������ ������
  SetLength(SearchBytes, Length(SearchStr));
  for i := 1 to Length(SearchStr) do
  begin
    SearchBytes[i-1] := Byte(SearchStr[i]);
  end;
  
  try
    AssignFile(F, FilePath);
    FileMode := fmOpenRead; // ��������� ������ ��� ������
    Reset(F, 1);  // ��������� ���� � ���������� ������
    
    // �������� ������ �����
    FileSize := System.FileSize(F);
    
    // ������ ���� �������� � ���� � ��� ������
    while not Eof(F) and not Terminated do
    begin
      BlockRead(F, Buffer, SizeOf(Buffer), BytesRead);
      
      // �������� �� ���� ������ � ������
      i := 0;
      while (i <= BytesRead - Length(SearchBytes)) and not Terminated do
      begin
        // ���������, ��������� �� ������� ���� � ������ ������ ������� ������
        if Buffer[i] = SearchBytes[0] then
        begin
          // ���� ���������, ��������� ��������� �����
          Match := True;
          for j := 1 to Length(SearchBytes) - 1 do
          begin
            if Buffer[i + j] <> SearchBytes[j] then
            begin
              Match := False;
              Break;
            end;
          end;
          
          // ���� ��� ����� �������, ����������� ������� ��������� ���������
          if Match then
          begin
            Inc(FoundCount);
            // ��������� ������� ���������� ��������� � �����
            Position := TotalBytesRead + i;
            FResults.Add(Format('������� ��������� �� �������: %d', [Position]));
          end;
        end;
        Inc(i);
      end;
      
      Inc(TotalBytesRead, BytesRead);
      
      // ��������� ��������
      if FileSize > 0 then
      begin
        FProgress := Min(99, Trunc((TotalBytesRead / FileSize) * 100));
      end;
    end;
    
    CloseFile(F);
    
    if Terminated then
    begin
      FResults.Add('����� ��� ������� �������������');
    end
    else
    begin
      FResults.Add(Format('������� ���������: %d', [FoundCount]));
    end;
    
  except
    on E: Exception do
    begin
      FResults.Add('������ ��� ������: ' + E.Message);
    end;
  end;
end;

// ����� ���������� ����� � �����
procedure TTaskThread.SearchMultipleStringsInFile(const FilePath: string; SearchStrings: TStringList);
var
  F: file;
  Buffer: array[0..4095] of Byte;
  BytesRead: Integer;
  TotalBytesRead: Int64;
  FileSize: Int64;
  i, j, k: Integer;
  Match: Boolean;
  SearchBytes: array of array of Byte;
  FoundCounts: array of Integer;
  Positions: array of TStringList;
  Position: Int64;
begin
  // �������������� ������� ��� �������� ������ ������ � ���������
  SetLength(SearchBytes, SearchStrings.Count);
  SetLength(FoundCounts, SearchStrings.Count);
  SetLength(Positions, SearchStrings.Count);
  
  FResults.Add('������ ������ ���������� ����� � �����:');
  FResults.Add('����: ' + FilePath);
  FResults.Add('������� ������:');
  
  // �������������� �������� � ����������� ������ ������ � ������� ������
  for i := 0 to SearchStrings.Count - 1 do
  begin
    FResults.Add('- ' + SearchStrings[i]);
    FoundCounts[i] := 0;
    
    // ����������� ������ � ������ ������
    SetLength(SearchBytes[i], Length(SearchStrings[i]));
    for j := 1 to Length(SearchStrings[i]) do
    begin
      SearchBytes[i][j-1] := Byte(SearchStrings[i][j]);
    end;
    
    // ������� ������ ��� �������� ������� ���������
    Positions[i] := TStringList.Create;
  end;
  
  if not FileExists(FilePath) then
  begin
    FResults.Add('������: ��������� ���� �� ����������');
    // ����������� ��������� ������
    for i := 0 to High(Positions) do
      Positions[i].Free;
    Exit;
  end;
  
  TotalBytesRead := 0;
  
  try
    AssignFile(F, FilePath);
    FileMode := fmOpenRead; // ��������� ������ ��� ������
    Reset(F, 1);  // ��������� ���� � ���������� ������
    
    // �������� ������ �����
    FileSize := System.FileSize(F);
    
    // ������ ���� �������� � ���� � ��� ������
    while not Eof(F) and not Terminated do
    begin
      BlockRead(F, Buffer, SizeOf(Buffer), BytesRead);
      
      // ���� ������ ������
      for k := 0 to High(SearchBytes) do
      begin
        // �������� �� ���� ������ � ������
        i := 0;
        while (i <= BytesRead - Length(SearchBytes[k])) and not Terminated do
        begin
          // ���������, ��������� �� ������� ���� � ������ ������ ������� ������
          if Buffer[i] = SearchBytes[k][0] then
          begin
            // ���� ���������, ��������� ��������� �����
            Match := True;
            for j := 1 to Length(SearchBytes[k]) - 1 do
            begin
              if (i + j < BytesRead) and (Buffer[i + j] <> SearchBytes[k][j]) then
              begin
                Match := False;
                Break;
              end;
            end;
            
            // ���� ��� ����� �������, ����������� ������� ��������� ���������
            if Match then
            begin
              Inc(FoundCounts[k]);
              // ��������� ������� ���������� ��������� � �����
              Position := TotalBytesRead + i;
              Positions[k].Add(IntToStr(Position));
            end;
          end;
          Inc(i);
        end;
      end;
      
      Inc(TotalBytesRead, BytesRead);
      
      // ��������� ��������
      if FileSize > 0 then
      begin
        FProgress := Min(99, Trunc((TotalBytesRead / FileSize) * 100));
      end;
    end;
    
    CloseFile(F);
    
    if Terminated then
    begin
      FResults.Add('����� ��� ������� �������������');
    end
    else
    begin
      // ������� ���������� ������
      FResults.Add('');
      FResults.Add('���������� ������:');
      for i := 0 to High(FoundCounts) do
      begin
        FResults.Add(Format('������ "%s": ������� ��������� %d', [SearchStrings[i], FoundCounts[i]]));
        
        // ������� ������� ���������
        if FoundCounts[i] > 0 then
        begin
          FResults.Add('������� ���������:');
          for j := 0 to Min(999, Positions[i].Count - 1) do // ������������ ���������� �������
          begin
            FResults.Add('- ' + Positions[i][j]);
          end;
          
          if Positions[i].Count > 1000 then
            FResults.Add('... (� ��� ' + IntToStr(Positions[i].Count - 1000) + ' ���������)');
        end;
      end;
    end;
    
  except
    on E: Exception do
    begin
      FResults.Add('������ ��� ������: ' + E.Message);
    end;
  end;
  
  // ����������� ��������� ������
  for i := 0 to High(Positions) do
    Positions[i].Free;
end;

procedure TTaskThread.Execute;
var
  TaskType: string;
  StartPath, FileMask: string;
  FilePath, SearchStr: string;
  MultiMasks, MultiSearchStrings: TStringList;
begin
  try
    FStatus := 1; // �����������
    FResults.Add('������ ���������� ������: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    FResults.Add('��� ������: ' + FTaskName);
    
    // ���������� ��� ������
    TaskType := FParameters.Values['TaskType'];
    
    if TaskType = 'FindFiles' then
    begin
      // ����� ������ �� �����
      StartPath := FParameters.Values['StartPath'];
      FileMask := FParameters.Values['FileMask'];
      
      if (StartPath <> '') and (FileMask <> '') then
      begin
        FindFiles(StartPath, FileMask);
      end
      else
      begin
        FResults.Add('������: �� ������� ������������ ��������� (StartPath ��� FileMask)');
        FStatus := 3; // ������
        FErrorMessage := '�� ������� ������������ ���������';
        FSuccess := False;
        Exit;
      end;
    end
    else if TaskType = 'FindFilesMultiMask' then
    begin
      // ����� ������ �� ���������� ������
      StartPath := FParameters.Values['StartPath'];
      
      // �������� ������ �����
      MultiMasks := TStringList.Create;
      try
        // ���� �������� MultiMasks ������� ��� ������
        MultiMasks.CommaText := FParameters.Values['MultiMasks'];
        
        if (StartPath <> '') and (MultiMasks.Count > 0) then
        begin
          FindFilesWithMultiMasks(StartPath, MultiMasks);
        end
        else
        begin
          FResults.Add('������: �� ������� ������������ ��������� (StartPath ��� MultiMasks)');
          FStatus := 3; // ������
          FErrorMessage := '�� ������� ������������ ���������';
          FSuccess := False;
        end;
      finally
        MultiMasks.Free;
      end;
    end
    else if TaskType = 'SearchInFile' then
    begin
      // ����� ������ � �����
      FilePath := FParameters.Values['FilePath'];
      SearchStr := FParameters.Values['SearchString'];
      
      if (FilePath <> '') and (SearchStr <> '') then
      begin
        SearchStringInFile(FilePath, SearchStr);
      end
      else
      begin
        FResults.Add('������: �� ������� ������������ ��������� (FilePath ��� SearchString)');
        FStatus := 3; // ������
        FErrorMessage := '�� ������� ������������ ���������';
        FSuccess := False;
        Exit;
      end;
    end
    else if TaskType = 'SearchMultipleInFile' then
    begin
      // ����� ���������� ����� � �����
      FilePath := FParameters.Values['FilePath'];
      
      // �������� ������ ����� ��� ������
      MultiSearchStrings := TStringList.Create;
      try
        // ���� ������ ������� ����� �����������
        MultiSearchStrings.CommaText := FParameters.Values['MultiSearchStrings'];
        
        if (FilePath <> '') and (MultiSearchStrings.Count > 0) then
        begin
          SearchMultipleStringsInFile(FilePath, MultiSearchStrings);
        end
        else
        begin
          FResults.Add('������: �� ������� ������������ ��������� (FilePath ��� MultiSearchStrings)');
          FStatus := 3; // ������
          FErrorMessage := '�� ������� ������������ ���������';
          FSuccess := False;
        end;
      finally
        MultiSearchStrings.Free;
      end;
    end
    else
    begin
      FResults.Add('������: ����������� ��� ������ - ' + TaskType);
      FStatus := 3; // ������
      FErrorMessage := '����������� ��� ������';
      FSuccess := False;
      Exit;
    end;
    
    // ��������� ������
    if FStatus <> 3 then // ���� �� ���� ������
    begin
      if Terminated then
      begin
        FStatus := 3; // ������
        FErrorMessage := '������ ���� ������������� �����������';
        FSuccess := False;
        FResultMessage := '������ �� ���������';
      end
      else
      begin
        FStatus := 2; // ���������
        FProgress := 100; // 100%
        FSuccess := True;
        FResultMessage := '������ ������� ���������';
        FResults.Add('������ ������� ���������: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
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
  // � ����� DLL 2 ���� �����
  Result := 2;
end;

// ���������� ���������� � ������ �� �������
function GetTaskInfo(Index: Integer; var TaskInfo: TTaskInfo): Boolean; stdcall;
begin
  Result := False;

  case Index of
    0: // FindFilesMultiMask
      begin
        StrPCopy(TaskInfo.Name, 'FindFilesMultiMask');
        StrPCopy(TaskInfo.Description, '����� ������ �� ���������� ������ � ��������� ��������');
        TaskInfo.ParamCount := 3; // StartPath, MultiMasks, IncludeSubfolders
        Result := True;
      end;
    1: // SearchMultipleInFile
      begin
        StrPCopy(TaskInfo.Name, 'SearchMultipleInFile');
        StrPCopy(TaskInfo.Description, '����� ��������� ���������� ����� � �����');
        TaskInfo.ParamCount := 2; // FilePath, MultiSearchStrings
        Result := True;
      end;
  end;
end;

// ���������� ���������� � ��������� ������
function GetTaskParamInfo(TaskIndex, ParamIndex: Integer; var ParamInfo: TTaskParamInfo): Boolean; stdcall;
begin
  Result := False;

  case TaskIndex of
    0: // FindFilesMultiMask
      case ParamIndex of
        0: // StartPath
          begin
            StrPCopy(ParamInfo.Name, 'StartPath');
            StrPCopy(ParamInfo.Description, '������� ��� ������ ������');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
        1: // MultiMasks
          begin
            StrPCopy(ParamInfo.Name, 'MultiMasks');
            StrPCopy(ParamInfo.Description, '����� ��� ������ ������, ����������� �������� (��������, *.txt,*.doc)');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
        2: // IncludeSubfolders
          begin
            StrPCopy(ParamInfo.Name, 'IncludeSubfolders');
            StrPCopy(ParamInfo.Description, '�������� ����������� � �����');
            ParamInfo.ParamType := ptBoolean;
            ParamInfo.Required := False;
            Result := True;
          end;
      end;
    1: // SearchMultipleInFile
      case ParamIndex of
        0: // FilePath
          begin
            StrPCopy(ParamInfo.Name, 'FilePath');
            StrPCopy(ParamInfo.Description, '���� � ����� ��� ������');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
        1: // MultiSearchStrings
          begin
            StrPCopy(ParamInfo.Name, 'MultiSearchStrings');
            StrPCopy(ParamInfo.Description, '������ ��� ������, ����������� ��������');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
      end;
  end;
end;

// ��������� ������ � ���������� � ID
function StartTask(TaskName: PChar; TaskId: Integer; ParamCount: Integer; ParamNames, ParamValues: Pointer): Boolean; stdcall;
var
  ParamNamesArray: TParamArray;
  ParamValuesArray: TParamValues;
  i: Integer;
begin
  Result := False;

  // ���������, ������ �� �������� �����
  if TaskManager = nil then
    Exit;

  // �������������� ������� ��� ����������
  if ParamCount > 0 then
  begin
    SetLength(ParamNamesArray, ParamCount);
    SetLength(ParamValuesArray, ParamCount);

    // �������� ��������� �� ����� � �������� ����������
    Move(ParamNames^, ParamNamesArray[0], ParamCount * SizeOf(PChar));
    Move(ParamValues^, ParamValuesArray[0], ParamCount * SizeOf(TTaskParamValue));

    // ���������, ��� ��� ��������� �����
    for i := 0 to ParamCount - 1 do
    begin
      if ParamNamesArray[i] = nil then
        Exit; // ���� ��� ��������� �� �������, �������
    end;

    // ��������� ������ � �����������
    if (CompareText(TaskName, 'FindFiles') = 0) or
       (CompareText(TaskName, 'FindFilesMultiMask') = 0) or
       (CompareText(TaskName, 'SearchInFile') = 0) or
       (CompareText(TaskName, 'SearchMultipleInFile') = 0) then
    begin
      TaskManager.AddTask(TaskName, TaskId, ParamCount, ParamNamesArray, ParamValuesArray);
      Result := True;
    end;
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