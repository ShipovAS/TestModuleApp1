library ShellCommandDLL;

uses
  Windows,
  SysUtils,
  Classes,
  ShellAPI,
  Math,
  DateUtils;

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
    Status: Integer;   // 0 - в ожидании, 1 - выполняется, 2 - завершена, 3 - ошибка
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

  TParamArray = array of PChar;
  TParamValues = array of TTaskParamValue;

  // Информация о запущенном процессе
  TProcessInfo = record
    hProcess: THandle;         // Дескриптор процесса
    hThread: THandle;          // Дескриптор основного потока
    dwProcessId: DWORD;        // ID процесса
    hStdOutRead: THandle;      // Канал для чтения StdOut
    hStdOutWrite: THandle;     // Канал для записи StdOut
    hStdErrRead: THandle;      // Канал для чтения StdErr
    hStdErrWrite: THandle;     // Канал для записи StdErr
    ExitCode: DWORD;           // Код завершения процесса
    StartTime: TDateTime;      // Время запуска процесса
  end;

  // Класс для выполнения задачи в отдельном потоке
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
    
    // Поля для работы с процессом
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

  // Менеджер задач
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

// Реализация класса TTaskThread
constructor TTaskThread.Create(ATaskId: Integer; const ATaskName: string);
begin
  inherited Create(True);
  FTaskId := ATaskId;
  FTaskName := ATaskName;
  FProgress := 0;
  FStatus := 0; // Ожидание
  FErrorMessage := '';
  FResults := TStringList.Create;
  FSuccess := False;
  FResultMessage := '';
  FParameters := TStringList.Create;
  FreeOnTerminate := False;
  
  // Инициализация полей для процесса
  FProcessInfo.hProcess := 0;
  FProcessInfo.hThread := 0;
  FProcessInfo.hStdOutRead := 0;
  FProcessInfo.hStdOutWrite := 0;
  FProcessInfo.hStdErrRead := 0;
  FProcessInfo.hStdErrWrite := 0;
end;

destructor TTaskThread.Destroy;
begin
  // Убеждаемся, что процесс завершен и ресурсы освобождены
  CleanupProcess;
  
  FParameters.Free;
  FResults.Free;
  inherited Destroy;
end;

// Функция запуска процесса с перенаправлением вывода
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
  
  // Инициализация структур безопасности для создания каналов
  FillChar(Security, SizeOf(Security), 0);
  Security.nLength := SizeOf(Security);
  Security.bInheritHandle := True;
  
  // Создаем канал для stdout
  if not CreatePipe(FProcessInfo.hStdOutRead, FProcessInfo.hStdOutWrite, @Security, 0) then
  begin
    FErrorMessage := 'Не удалось создать канал для stdout';
    Exit;
  end;
  
  // Создаем канал для stderr
  if not CreatePipe(FProcessInfo.hStdErrRead, FProcessInfo.hStdErrWrite, @Security, 0) then
  begin
    CloseHandle(FProcessInfo.hStdOutRead);
    CloseHandle(FProcessInfo.hStdOutWrite);
    FErrorMessage := 'Не удалось создать канал для stderr';
    Exit;
  end;
  
  // Устанавливаем наследуемость только для дескрипторов записи
  SetHandleInformation(FProcessInfo.hStdOutRead, HANDLE_FLAG_INHERIT, 0);
  SetHandleInformation(FProcessInfo.hStdErrRead, HANDLE_FLAG_INHERIT, 0);
  
  // Инициализация структуры StartupInfo
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

  // Подготавливаем строку командной строки
  StrPCopy(CommandLineStr, CommandLine);
  
  // Инициализация структуры информации о процессе
  FillChar(FProcessInfo, SizeOf(TProcessInfo), 0);
  
  // Создаем процесс
  ProcessCreated := CreateProcess(
    nil,                  // Имя приложения
    CommandLineStr,       // Командная строка
    nil,                  // Атрибуты безопасности процесса
    nil,                  // Атрибуты безопасности потока
    True,                 // Наследование дескрипторов
    CREATE_NEW_CONSOLE,   // Флаги создания
    nil,                  // Окружение родительского процесса
    PChar(WorkingDir),    // Текущий каталог
    StartupInfo,          // Информация о старте
    PI // Информация о процессе
  );

  FProcessInfo.hProcess := PI.hProcess;
  FProcessInfo.hThread := PI.hThread;
  FProcessInfo.dwProcessId := PI.dwProcessId;
  
  // Закрываем дескрипторы записи, так как они теперь используются дочерним процессом
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
    FErrorMessage := Format('Не удалось создать процесс: %s (Код ошибки: %d)', 
                          [SysErrorMessage(GetLastError), GetLastError]);
    CleanupProcess;
  end;
end;

// Проверка, запущен ли процесс
function TTaskThread.IsProcessRunning: Boolean;
var
  ExitCode: DWORD;
begin
  Result := False;
  
  if FProcessInfo.hProcess <> 0 then
  begin
    // Получаем статус процесса
    GetExitCodeProcess(FProcessInfo.hProcess, ExitCode);
    Result := (ExitCode = STILL_ACTIVE);
    FProcessInfo.ExitCode := ExitCode;
  end;
end;

// Принудительное завершение процесса
procedure TTaskThread.TerminateProcess;
begin
  if FProcessInfo.hProcess <> 0 then
  begin
    // Принудительно завершаем процесс
    Windows.TerminateProcess(FProcessInfo.hProcess, 1);
  end;
end;

// Освобождение ресурсов процесса
procedure TTaskThread.CleanupProcess;
begin
  // Закрываем все открытые дескрипторы
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

// Чтение вывода процесса
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
    // Проверяем наличие данных в stdout
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
    
    // Проверяем наличие данных в stderr
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
    
    // Если нет данных в обоих каналах, но процесс еще работает, ждем немного
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
    FStatus := 1; // Выполняется
    FResults.Add('Начало выполнения команды: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    
    // Получаем параметры
    FCommandLine := FParameters.Values['CommandLine'];
    FWorkingDir := FParameters.Values['WorkingDirectory'];
    FShowWindow := StrToBoolDef(FParameters.Values['ShowWindow'], False);
    FTimeout := StrToIntDef(FParameters.Values['Timeout'], 0); // 0 = без таймаута
    
    // Проверяем наличие обязательных параметров
    if FCommandLine = '' then
    begin
      FStatus := 3; // Ошибка
      FErrorMessage := 'Не указана команда для выполнения';
      FSuccess := False;
      FResultMessage := 'Ошибка: не указана команда';
      Exit;
    end;
    
    // Если рабочая директория не указана, используем текущую
    if FWorkingDir = '' then
      FWorkingDir := GetCurrentDir;
    
    // Выводим информацию о параметрах
    FResults.Add('Параметры команды:');
    FResults.Add('- Командная строка: ' + FCommandLine);
    FResults.Add('- Рабочая директория: ' + FWorkingDir);
    FResults.Add('- Показывать окно: ' + BoolToStr(FShowWindow, True));
    if FTimeout > 0 then
      FResults.Add('- Таймаут: ' + IntToStr(FTimeout) + ' секунд')
    else
      FResults.Add('- Таймаут: не установлен');
    
    // Запускаем процесс
    FResults.Add('Запуск команды...');
    if not ExecuteShellCommand(FCommandLine, FWorkingDir, FShowWindow) then
    begin
      FStatus := 3; // Ошибка
      FSuccess := False;
      FResultMessage := 'Ошибка запуска команды: ' + FErrorMessage;
      Exit;
    end;
    
    // Процесс запущен успешно
    FResults.Add('Команда запущена успешно. PID: ' + IntToStr(FProcessInfo.dwProcessId));
    FProgress := 1; // Начинаем с 1%
    
    // Запоминаем время начала
    StartTime := Now;
    HadTimeout := False;
    
    // Основной цикл мониторинга процесса
    while IsProcessRunning and not Terminated do
    begin
      // Читаем вывод процесса
      ReadProcessOutput;
      
      // Вычисляем прогресс на основе времени выполнения
      if FTimeout > 0 then
      begin
        // Если установлен таймаут, процент выполнения основан на соотношении прошедшего времени к таймауту
        RunningTime := SecondsBetween(Now, StartTime);
        FProgress := Min(99, Trunc((RunningTime / FTimeout) * 100));
        
        // Проверяем, не истек ли таймаут
        if RunningTime >= FTimeout then
        begin
          FResults.Add('Превышен таймаут выполнения команды (' + IntToStr(FTimeout) + ' секунд)');
          TerminateProcess;
          HadTimeout := True;
          Break;
        end;
      end
      else
      begin
        // Если таймаут не установлен, показываем условный прогресс 
        // на основе длительности выполнения, но не более 99%
        ElapsedSeconds := SecondsBetween(Now, StartTime);
        if ElapsedSeconds <= 60 then
          FProgress := Min(50, ElapsedSeconds) // Первые 60 секунд - до 50%
        else
          FProgress := Min(99, 50 + Trunc(((ElapsedSeconds - 60) / 600) * 49)); // Ещё 600 секунд до 99%
      end;
      
      // Добавляем информацию о времени выполнения каждые 5 секунд
      if (SecondsBetween(Now, StartTime) mod 5 = 0) then
      begin
        FResults.Add('Время выполнения: ' + IntToStr(SecondsBetween(Now, StartTime)) + ' секунд');
      end;
      
      Sleep(100); // Небольшая задержка для уменьшения нагрузки
    end;
    
    // Завершаем чтение вывода
    ReadProcessOutput;
    
    // Проверяем результат выполнения
    if Terminated then
    begin
      // Задача была прервана пользователем
      FResults.Add('Выполнение команды прервано пользователем');
      TerminateProcess;
      FStatus := 3; // Ошибка
      FErrorMessage := 'Задача была принудительно остановлена';
      FSuccess := False;
      FResultMessage := 'Задача не завершена';
    end
    else if HadTimeout then
    begin
      // Процесс был остановлен из-за таймаута
      FStatus := 3; // Ошибка
      FErrorMessage := 'Превышен таймаут выполнения';
      FSuccess := False;
      FResultMessage := 'Задача не завершена из-за таймаута';
    end
    else
    begin
      // Процесс завершился сам
      FStatus := 2; // Завершена
      FProgress := 100; // 100%
      
      if FProcessInfo.ExitCode = 0 then
      begin
        FSuccess := True;
        FResultMessage := 'Команда успешно выполнена';
        FResults.Add('Команда успешно завершена с кодом 0');
      end
      else
      begin
        FSuccess := False;
        FResultMessage := 'Команда завершена с ошибкой, код: ' + IntToStr(FProcessInfo.ExitCode);
        FResults.Add('Команда завершена с кодом ошибки: ' + IntToStr(FProcessInfo.ExitCode));
      end;
    end;
    
    // Добавляем информацию о времени выполнения
    FResults.Add('Общее время выполнения: ' + IntToStr(SecondsBetween(Now, StartTime)) + ' секунд');
    FResults.Add('Завершение задачи: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    
  except
    on E: Exception do
    begin
      FStatus := 3; // Ошибка
      FErrorMessage := E.Message;
      FSuccess := False;
      FResultMessage := 'Ошибка при выполнении задачи: ' + E.Message;
      FResults.Add('Произошла ошибка: ' + E.Message);
    end;
  end;
  
  // Освобождаем ресурсы процесса
  CleanupProcess;
end;

// Реализация класса TTaskManager
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
      if Task.Status = 1 then // Если задача выполняется
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
    // Генерируем уникальный ID для задачи
    TaskId := pTaskId;

    // Создаем поток задачи
    Task := TTaskThread.Create(TaskId, TaskName);

    // Добавляем тип задачи как параметр
    Task.Parameters.Values['TaskType'] := TaskName;

    // Добавляем параметры в задачу
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

        // Добавляем параметр в список параметров задачи
        Task.Parameters.Values[ParamName] := ParamValue;
      end;
    end;

    FTasks.Add(Task);
    Task.Resume; // Запускаем поток

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
      TaskResult.ResultData := Task.Results; // Передаем указатель на список результатов
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
    if (Task <> nil) and (Task.Status = 1) then // Если задача выполняется
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
        // Если задача все еще выполняется, останавливаем ее
        if Task.Status = 1 then
          Task.Terminate;

        // Ждем завершения потока
        Task.WaitFor;

        // Удаляем задачу из списка и освобождаем ресурсы
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

// Экспортируемые функции DLL

// Возвращает количество доступных задач
function GetTaskCount: Integer; stdcall;
begin
  // В нашей DLL только один тип задачи - выполнение shell-команды
  Result := 1;
end;

// Возвращает информацию о задаче по индексу
function GetTaskInfo(Index: Integer; var TaskInfo: TTaskInfo): Boolean; stdcall;
begin
  Result := False;

  if Index = 0 then // ExecuteShellCommand
  begin
    StrPCopy(TaskInfo.Name, 'ExecuteShellCommand');
    StrPCopy(TaskInfo.Description, 'Выполнение shell-команды с отслеживанием прогресса');
    TaskInfo.ParamCount := 4; // CommandLine, WorkingDirectory, ShowWindow, Timeout
    Result := True;
  end;
end;

// Возвращает информацию о параметре задачи
function GetTaskParamInfo(TaskIndex, ParamIndex: Integer; var ParamInfo: TTaskParamInfo): Boolean; stdcall;
begin
  Result := False;

  if TaskIndex <> 0 then
    Exit;

  case ParamIndex of
    0: // CommandLine
      begin
        StrPCopy(ParamInfo.Name, 'CommandLine');
        StrPCopy(ParamInfo.Description, 'Командная строка для выполнения');
        ParamInfo.ParamType := ptString;
        ParamInfo.Required := True;
        Result := True;
      end;
    1: // WorkingDirectory
      begin
        StrPCopy(ParamInfo.Name, 'WorkingDirectory');
        StrPCopy(ParamInfo.Description, 'Рабочая директория для выполнения команды');
        ParamInfo.ParamType := ptString;
        ParamInfo.Required := False;
        Result := True;
      end;
    2: // ShowWindow
      begin
        StrPCopy(ParamInfo.Name, 'ShowWindow');
        StrPCopy(ParamInfo.Description, 'Показывать окно процесса');
        ParamInfo.ParamType := ptBoolean;
        ParamInfo.Required := False;
        Result := True;
      end;
    3: // Timeout
      begin
        StrPCopy(ParamInfo.Name, 'Timeout');
        StrPCopy(ParamInfo.Description, 'Таймаут выполнения команды в секундах (0 - без таймаута)');
        ParamInfo.ParamType := ptInteger;
        ParamInfo.Required := False;
        Result := True;
      end;
  end;
end;

// Запускает задачу и возвращает её ID
function StartTask(TaskName: PChar; TaskId: Integer; ParamCount: Integer; ParamNames, ParamValues: Pointer): Boolean; stdcall;
var
  ParamNamesArray: TParamArray;
  ParamValuesArray: TParamValues;
begin
  Result := False;

  // Проверяем, создан ли менеджер задач
  if TaskManager = nil then
    Exit;

  // Проверяем, поддерживается ли задача
  if CompareText(TaskName, 'ExecuteShellCommand') <> 0 then
    Exit;

  // Подготавливаем массивы для параметров
  if ParamCount > 0 then
  begin
    SetLength(ParamNamesArray, ParamCount);
    SetLength(ParamValuesArray, ParamCount);

    // Копируем указатели на имена и значения параметров
    Move(ParamNames^, ParamNamesArray[0], ParamCount * SizeOf(PChar));
    Move(ParamValues^, ParamValuesArray[0], ParamCount * SizeOf(TTaskParamValue));

    // Запускаем задачу с параметрами
    TaskManager.AddTask(TaskName, TaskId, ParamCount, ParamNamesArray, ParamValuesArray);
    Result := True;
  end
  else
  begin
    // Без параметров нельзя запустить нашу задачу
    Result := False;
  end;
end;

// Возвращает текущий прогресс выполнения задачи
function GetTaskProgress(TaskId: Integer; var Progress: TTaskProgress): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.GetTaskProgress(TaskId, Progress);
end;

// Возвращает результат выполнения задачи
function GetTaskResult(TaskId: Integer; var TaskResult: TTaskResult): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.GetTaskResult(TaskId, TaskResult);
end;

// Возвращает детальную информацию о результате задачи
function GetTaskResultDetail(TaskId, Index: Integer; Buffer: PChar; BufSize: Integer): Boolean; stdcall;
begin
  Result := False;
  if TaskManager <> nil then
    Result := TaskManager.GetTaskResultDetail(TaskId, Index, Buffer, BufSize);
end;

// Останавливает выполнение задачи
function StopTask(TaskId: Integer): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.StopTask(TaskId);
end;

// Освобождает ресурсы, выделенные для задачи
function FreeTask(TaskId: Integer): Boolean; stdcall;
begin
  Result := False;

  if TaskManager <> nil then
    Result := TaskManager.FreeTask(TaskId);
end;

// Инициализация и финализация DLL
procedure DLLEntryPoint(dwReason: DWORD);
begin
  case dwReason of
    DLL_PROCESS_ATTACH:
      begin
        // Создаем менеджер задач при загрузке DLL
        TaskManager := TTaskManager.Create;
      end;
    DLL_PROCESS_DETACH:
      begin
        // Освобождаем менеджер задач при выгрузке DLL
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