library TaskLibrary;

uses
  Windows,
  SysUtils,
  Classes,
  Dialogs;

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
      0:
        (StringValue: array[0..1023] of Char);
      1:
        (IntValue: Integer);
      2:
        (BoolValue: Boolean);
      3:
        (StringListPtr: Pointer); // Указатель на TStringList
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
    FParameters: TStringList; // Добавляем хранилище для параметров
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
  FParameters := TStringList.Create; // Инициализируем список параметров
  FreeOnTerminate := False;
end;

destructor TTaskThread.Destroy;
begin
  FParameters.Free; // Освобождаем память для списка параметров
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
    FStatus := 1; // Выполняется
    // Добавляем некоторые результаты для тестирования
    FResults.Add('Начало выполнения задачи: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    FResults.Add('Имя задачи: ' + FTaskName);

    // Получаем значения параметров
    DelayParam := FParameters.Values['Delay'];
//    showmessage('DelayParam='+DelayParam);
    ShowDetailsParam := FParameters.Values['ShowDetails'];
    CommentParam := FParameters.Values['Comment'];

    // Проверяем, есть ли параметр для времени выполнения
    CustomDelay := 20; // По умолчанию 20 секунд
    if DelayParam <> '' then
    begin
      try
        CustomDelay := StrToInt(DelayParam);
        if CustomDelay <= 0 then
          CustomDelay := 1;
        if CustomDelay > 300 then
          CustomDelay := 300; // Ограничиваем максимальное время 5 минутами
      except
        // Если ошибка преобразования, используем значение по умолчанию
      end;
    end;

    // Определяем, нужно ли показывать детали
    ShowDetails := False;
    if ShowDetailsParam <> '' then
    begin
      ShowDetails := (ShowDetailsParam = 'True');
    end;

    // Добавляем информацию о параметрах
    FResults.Add('');
    FResults.Add('Параметры задачи:');
    FResults.Add(Format('- Время выполнения: %d секунд', [CustomDelay]));
    FResults.Add(Format('- Показывать детали: %s', [BoolToStr(ShowDetails, True)]));

    if CommentParam <> '' then
    begin
      FResults.Add(Format('- Комментарий: %s', [CommentParam]));
    end
    else
    begin
      FResults.Add('- Комментарий: не указан');
    end;

    // Для тестовой задачи просто имитируем работу в течение указанного времени
    StartTime := GetTickCount;
    TotalTime := CustomDelay * 1000; // Конвертируем в миллисекунды

    while not Terminated and (FProgress < 100) do
    begin
      Sleep(100); // Небольшая задержка

      ElapsedTime := GetTickCount - StartTime;
      if ElapsedTime >= TotalTime then
        FProgress := 100
      else
        FProgress := Round((ElapsedTime / TotalTime) * 100);

      // Добавляем информацию о прогрессе
      if ShowDetails then
      begin
        // Если включены детали, добавляем информацию о каждых 5%
        if (FProgress mod 5 = 0) and (FProgress > 0) and (FProgress < 100) then
          FResults.Add(Format('Прогресс выполнения: %d%% (прошло %d секунд)', [FProgress, ElapsedTime div 1000]));
      end
      else
      begin
        // Иначе только о каждых 25%
        if (FProgress mod 25 = 0) and (FProgress > 0) and (FProgress < 100) then
          FResults.Add(Format('Прогресс выполнения: %d%% (прошло %d секунд)', [FProgress, ElapsedTime div 1000]));
      end;
    end;

    if Terminated then
    begin
      FStatus := 3; // Ошибка
      FErrorMessage := 'Задача была принудительно остановлена';
      FSuccess := False;
      FResultMessage := 'Задача не завершена';
      FResults.Add('Задача была прервана пользователем');
    end
    else
    begin
      FStatus := 2; // Завершена
      FSuccess := True;
      FResultMessage := 'Задача успешно выполнена';
      FResults.Add('Задача успешно завершена: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
      FResults.Add(Format('Общее время выполнения: %d секунд', [ElapsedTime div 1000]));

      // Добавляем комментарий в результат, если он был указан
      if CommentParam <> '' then
      begin
        FResults.Add('');
        FResults.Add('Комментарий пользователя: ' + CommentParam);
      end;
    end;
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
    // Генерируем уникальный ID для задачи
    TaskId := pTaskId;

    // Создаем и запускаем поток задачи
    Task := TTaskThread.Create(TaskId, TaskName);

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
                ParamValue := TStringList(ParamValues[i].StringListPtr).Text
              else
                ParamValue := '(пустой список)';
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
  // В тестовой DLL у нас только одна задача
  Result := 1;
end;

// Возвращает информацию о задаче по индексу
function GetTaskInfo(Index: Integer; var TaskInfo: TTaskInfo): Boolean; stdcall;
begin
  Result := False;

  if Index = 0 then
  begin
    StrPCopy(TaskInfo.Name, 'TestTask');
    StrPCopy(TaskInfo.Description, 'Тестовая задача, которая просто заполняет прогресс от 0% до 100% в течение указанного времени');
    TaskInfo.ParamCount := 3; // Теперь у тестовой задачи есть 3 параметра
    Result := True;
  end;
end;

// Возвращает информацию о параметре задачи
function GetTaskParamInfo(TaskIndex, ParamIndex: Integer; var ParamInfo: TTaskParamInfo): Boolean; stdcall;
begin
  Result := False;

  // Проверяем индекс задачи
  if TaskIndex <> 0 then
    Exit;

  // Определяем параметры для тестовой задачи
  case ParamIndex of
    0: // Параметр "Delay"
      begin
        StrPCopy(ParamInfo.Name, 'Delay');
        StrPCopy(ParamInfo.Description, 'Время выполнения задачи в секундах');
        ParamInfo.ParamType := ptInteger;
        ParamInfo.Required := False;
        Result := True;
      end;
    1: // Параметр "ShowDetails"
      begin
        StrPCopy(ParamInfo.Name, 'ShowDetails');
        StrPCopy(ParamInfo.Description, 'Показывать детальную информацию о выполнении');
        ParamInfo.ParamType := ptBoolean;
        ParamInfo.Required := False;
        Result := True;
      end;
    2: // Параметр "Comment"
      begin
        StrPCopy(ParamInfo.Name, 'Comment');
        StrPCopy(ParamInfo.Description, 'Комментарий к задаче');
        ParamInfo.ParamType := ptString;
        ParamInfo.Required := False;
        Result := True;
      end;
  end;
end;

// Запускает задачу и возвращает её ID
function StartTask(TaskName: PChar; pTaskId: Integer; ParamCount: Integer; ParamNames, ParamValues: Pointer): Boolean; stdcall;
var
  i: Integer;
  ParamNamesArray: TParamArray;
  ParamValuesArray: TParamValues;
  ii: Integer;
  ss: string;
begin
  Result := False;

  // Проверяем, создан ли менеджер задач
  if TaskManager = nil then
    Exit;

  // Проверяем, поддерживается ли задача
  if CompareText(TaskName, 'TestTask') = 0 then
  begin
    // Подготавливаем массивы для параметров
    if ParamCount > 0 then
    begin
      SetLength(ParamNamesArray, ParamCount);
      SetLength(ParamValuesArray, ParamCount);

      // Копируем указатели на имена и значения параметров
      Move(ParamNames^, ParamNamesArray[0], ParamCount * SizeOf(PChar));
      Move(ParamValues^, ParamValuesArray[0], ParamCount * SizeOf(TTaskParamValue));


//      ShowMessage('ss:');
//      for ii := 1 to ParamCount do
//        ss:=ss+ParamNamesArray[ii]+'='+IntToStr(ParamValuesArray[ii].IntValue)+'; ';
//      ShowMessage(ss);
      // Запускаем задачу с параметрами
      TaskManager.AddTask(TaskName, pTaskId, ParamCount, ParamNamesArray, ParamValuesArray);

    end
    else
    begin
      // Запускаем задачу без параметров
      //ShowMessage('No params');
      TaskManager.AddTask(TaskName, pTaskId);
    end;

    Result := True;
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

