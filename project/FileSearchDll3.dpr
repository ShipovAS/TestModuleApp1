library FileSearchDLL3;

uses
  Windows,
  SysUtils,
  Classes,
  Masks,
  Math,
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
    
    // Для поиска файлов
    procedure FindFiles(const StartPath, FileMask: string);
    procedure FindFilesWithMultiMasks(const StartPath: string; Masks: TStringList);
    
    // Для поиска строк в файле
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
  FParameters := TStringList.Create;
  FreeOnTerminate := False;
end;

destructor TTaskThread.Destroy;
begin
  // Дополнительно проверяем, что потоки завершены
  if Self.Status = 1 then // Если задача еще выполняется
  begin
    Self.Terminate;
    // Даем немного времени на корректное завершение
    Sleep(1000);
  end;

  try
    FParameters.Free;
    FResults.Free;
  finally
    // Обработать исключения при освобождении ресурсов
    //showmessage('(DEBUG): Исключения при освобождении ресурсов');
  end;
  inherited Destroy;
end;

// Поиск файлов по одной маске
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
  
  // Проверяем, нужно ли искать в подпапках
  IncludeSubfolders := StrToBoolDef(FParameters.Values['IncludeSubfolders'], True);
  
  FResults.Add('Начало поиска файлов по маске: ' + FileMask);
  FResults.Add('Начальная папка: ' + StartPath);
  FResults.Add('Включать подпапки: ' + BoolToStr(IncludeSubfolders, True));
  
  if not DirectoryExists(StartPath) then
  begin
    FResults.Add('Ошибка: указанная папка не существует');
    Exit;
  end;
  
  // Создаем список для хранения папок
  DirList := TStringList.Create;
  try
    // Добавляем начальную папку
    DirList.Add(IncludeTrailingPathDelimiter(StartPath));
    
    // Пока список папок не пуст и задача не отменена
    i := 0;
    while (i < DirList.Count) and not Terminated do
    begin
      CurrentPath := DirList[i];
      
      // Ищем файлы в текущей папке
      if FindFirst(CurrentPath + FileMask, faAnyFile - faDirectory, SearchRec) = 0 then
      begin
        try
          repeat
            Inc(FileCount);
            Inc(TotalFiles);
            // Добавляем путь к файлу в результаты
            FResults.Add('Найден файл: ' + CurrentPath + SearchRec.Name);
            
            // Обновляем прогресс каждые 10 файлов
            if (FileCount mod 10 = 0) then
            begin
              FProgress := Min(99, FProgress + 1); // Не доходим до 100%, пока не закончим
              FResults.Add(Format('Найдено файлов: %d', [TotalFiles]));
            end;
            
            if Terminated then Break;
          until FindNext(SearchRec) <> 0;
        finally
          FindClose(SearchRec);
        end;
      end;
      
      // Если нужно искать в подпапках
      if IncludeSubfolders and not Terminated then
      begin
        // Находим все подпапки
        if FindFirst(CurrentPath + '*', faDirectory, SearchRec) = 0 then
        begin
          try
            repeat
              // Пропускаем . и ..
              if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') and
                 ((SearchRec.Attr and faDirectory) = faDirectory) then
              begin
                // Добавляем путь к папке в список для обработки
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
      FResults.Add('Поиск был прерван пользователем');
    end
    else
    begin
      FResults.Add(Format('Найдено всего файлов: %d', [TotalFiles]));
    end;
    
  finally
    DirList.Free;
  end;
end;

// Поиск файлов по нескольким маскам
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
  
  // Проверяем, нужно ли искать в подпапках
  IncludeSubfolders := StrToBoolDef(FParameters.Values['IncludeSubfolders'], True);
  
  FResults.Add('Начало поиска файлов по нескольким маскам:');
  for i := 0 to Masks.Count - 1 do
  begin
    FResults.Add('- ' + Masks[i]);
  end;
  FResults.Add('Начальная папка: ' + StartPath);
  FResults.Add('Включать подпапки: ' + BoolToStr(IncludeSubfolders, True));
  
  if not DirectoryExists(StartPath) then
  begin
    FResults.Add('Ошибка: указанная папка не существует');
    Exit;
  end;
  
  // Создаем список для хранения найденных файлов и папок для обхода
  FilePaths := TStringList.Create;
  DirList := TStringList.Create;
  try
    // Добавляем начальную папку
    DirList.Add(IncludeTrailingPathDelimiter(StartPath));
    
    // Пока список папок не пуст и задача не отменена
    j := 0;
    while (j < DirList.Count) and not Terminated do
    begin
      CurrentPath := DirList[j];
      
      // Проверяем все маски в текущей папке
      for i := 0 to Masks.Count - 1 do
      begin
        if Terminated then Break;
        
        CurrentMask := Masks[i];
        if FindFirst(CurrentPath + CurrentMask, faAnyFile - faDirectory, SearchRec) = 0 then
        begin
          try
            repeat
              // Проверяем, не добавлен ли уже файл (чтобы избежать дубликатов)
              if FilePaths.IndexOf(CurrentPath + SearchRec.Name) < 0 then
              begin
                FilePaths.Add(CurrentPath + SearchRec.Name);
                Inc(TotalFilesFound);
                FResults.Add('Найден файл: ' + CurrentPath + SearchRec.Name + ' (по маске: ' + CurrentMask + ')');
                
                // Обновляем прогресс
                if (TotalFilesFound mod 10 = 0) then
                begin
                  FProgress := Min(99, FProgress + 1);
                  FResults.Add(Format('Найдено файлов: %d', [TotalFilesFound]));
                end;
              end;
              
              if Terminated then Break;
            until FindNext(SearchRec) <> 0;
          finally
            FindClose(SearchRec);
          end;
        end;
      end;
      
      // Если нужно искать в подпапках
      if IncludeSubfolders and not Terminated then
      begin
        // Находим все подпапки
        if FindFirst(CurrentPath + '*', faDirectory, SearchRec) = 0 then
        begin
          try
            repeat
              // Пропускаем . и ..
              if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') and
                 ((SearchRec.Attr and faDirectory) = faDirectory) then
              begin
                // Добавляем путь к папке в список для обработки
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
      FResults.Add('Поиск был прерван пользователем');
    end
    else
    begin
      FResults.Add(Format('Найдено всего файлов: %d', [TotalFilesFound]));
      // Добавляем все пути к файлам в результаты
      FResults.Add('');
      FResults.Add('Список найденных файлов:');
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

// Поиск одной строки в файле
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
  
  FResults.Add('Начало поиска строки в файле:');
  FResults.Add('Файл: ' + FilePath);
  FResults.Add('Искомая строка: ' + SearchStr);
  
  if not FileExists(FilePath) then
  begin
    FResults.Add('Ошибка: указанный файл не существует');
    Exit;
  end;
  
  // Преобразуем строку поиска в массив байтов
  SetLength(SearchBytes, Length(SearchStr));
  for i := 1 to Length(SearchStr) do
  begin
    SearchBytes[i-1] := Byte(SearchStr[i]);
  end;
  
  try
    AssignFile(F, FilePath);
    FileMode := fmOpenRead; // Открываем только для чтения
    Reset(F, 1);  // Открываем файл в побайтовом режиме
    
    // Получаем размер файла
    FileSize := System.FileSize(F);
    
    // Читаем файл порциями и ищем в нем строку
    while not Eof(F) and not Terminated do
    begin
      BlockRead(F, Buffer, SizeOf(Buffer), BytesRead);
      
      // Проходим по всем байтам в буфере
      i := 0;
      while (i <= BytesRead - Length(SearchBytes)) and not Terminated do
      begin
        // Проверяем, совпадает ли текущий байт с первым байтом искомой строки
        if Buffer[i] = SearchBytes[0] then
        begin
          // Если совпадает, проверяем остальные байты
          Match := True;
          for j := 1 to Length(SearchBytes) - 1 do
          begin
            if Buffer[i + j] <> SearchBytes[j] then
            begin
              Match := False;
              Break;
            end;
          end;
          
          // Если все байты совпали, увеличиваем счетчик найденных вхождений
          if Match then
          begin
            Inc(FoundCount);
            // Вычисляем позицию найденного вхождения в файле
            Position := TotalBytesRead + i;
            FResults.Add(Format('Найдено вхождение на позиции: %d', [Position]));
          end;
        end;
        Inc(i);
      end;
      
      Inc(TotalBytesRead, BytesRead);
      
      // Обновляем прогресс
      if FileSize > 0 then
      begin
        FProgress := Min(99, Trunc((TotalBytesRead / FileSize) * 100));
      end;
    end;
    
    CloseFile(F);
    
    if Terminated then
    begin
      FResults.Add('Поиск был прерван пользователем');
    end
    else
    begin
      FResults.Add(Format('Найдено вхождений: %d', [FoundCount]));
    end;
    
  except
    on E: Exception do
    begin
      FResults.Add('Ошибка при поиске: ' + E.Message);
    end;
  end;
end;

// Поиск нескольких строк в файле
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
  // Инициализируем массивы для хранения байтов поиска и счетчиков
  SetLength(SearchBytes, SearchStrings.Count);
  SetLength(FoundCounts, SearchStrings.Count);
  SetLength(Positions, SearchStrings.Count);
  
  FResults.Add('Начало поиска нескольких строк в файле:');
  FResults.Add('Файл: ' + FilePath);
  FResults.Add('Искомые строки:');
  
  // Инициализируем счетчики и преобразуем строки поиска в массивы байтов
  for i := 0 to SearchStrings.Count - 1 do
  begin
    FResults.Add('- ' + SearchStrings[i]);
    FoundCounts[i] := 0;
    
    // Преобразуем строку в массив байтов
    SetLength(SearchBytes[i], Length(SearchStrings[i]));
    for j := 1 to Length(SearchStrings[i]) do
    begin
      SearchBytes[i][j-1] := Byte(SearchStrings[i][j]);
    end;
    
    // Создаем список для хранения позиций вхождений
    Positions[i] := TStringList.Create;
  end;
  
  if not FileExists(FilePath) then
  begin
    FResults.Add('Ошибка: указанный файл не существует');
    // Освобождаем созданные списки
    for i := 0 to High(Positions) do
      Positions[i].Free;
    Exit;
  end;
  
  TotalBytesRead := 0;
  
  try
    AssignFile(F, FilePath);
    FileMode := fmOpenRead; // Открываем только для чтения
    Reset(F, 1);  // Открываем файл в побайтовом режиме
    
    // Получаем размер файла
    FileSize := System.FileSize(F);
    
    // Читаем файл порциями и ищем в нем строки
    while not Eof(F) and not Terminated do
    begin
      BlockRead(F, Buffer, SizeOf(Buffer), BytesRead);
      
      // Ищем каждую строку
      for k := 0 to High(SearchBytes) do
      begin
        // Проходим по всем байтам в буфере
        i := 0;
        while (i <= BytesRead - Length(SearchBytes[k])) and not Terminated do
        begin
          // Проверяем, совпадает ли текущий байт с первым байтом искомой строки
          if Buffer[i] = SearchBytes[k][0] then
          begin
            // Если совпадает, проверяем остальные байты
            Match := True;
            for j := 1 to Length(SearchBytes[k]) - 1 do
            begin
              if (i + j < BytesRead) and (Buffer[i + j] <> SearchBytes[k][j]) then
              begin
                Match := False;
                Break;
              end;
            end;
            
            // Если все байты совпали, увеличиваем счетчик найденных вхождений
            if Match then
            begin
              Inc(FoundCounts[k]);
              // Вычисляем позицию найденного вхождения в файле
              Position := TotalBytesRead + i;
              Positions[k].Add(IntToStr(Position));
            end;
          end;
          Inc(i);
        end;
      end;
      
      Inc(TotalBytesRead, BytesRead);
      
      // Обновляем прогресс
      if FileSize > 0 then
      begin
        FProgress := Min(99, Trunc((TotalBytesRead / FileSize) * 100));
      end;
    end;
    
    CloseFile(F);
    
    if Terminated then
    begin
      FResults.Add('Поиск был прерван пользователем');
    end
    else
    begin
      // Выводим результаты поиска
      FResults.Add('');
      FResults.Add('Результаты поиска:');
      for i := 0 to High(FoundCounts) do
      begin
        FResults.Add(Format('Строка "%s": найдено вхождений %d', [SearchStrings[i], FoundCounts[i]]));
        
        // Выводим позиции вхождений
        if FoundCounts[i] > 0 then
        begin
          FResults.Add('Позиции вхождений:');
          for j := 0 to Min(999, Positions[i].Count - 1) do // Ограничиваем количество позиций
          begin
            FResults.Add('- ' + Positions[i][j]);
          end;
          
          if Positions[i].Count > 1000 then
            FResults.Add('... (и ещё ' + IntToStr(Positions[i].Count - 1000) + ' вхождений)');
        end;
      end;
    end;
    
  except
    on E: Exception do
    begin
      FResults.Add('Ошибка при поиске: ' + E.Message);
    end;
  end;
  
  // Освобождаем созданные списки
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
    FStatus := 1; // Выполняется
    FResults.Add('Начало выполнения задачи: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
    FResults.Add('Имя задачи: ' + FTaskName);
    
    // Определяем тип задачи
    TaskType := FParameters.Values['TaskType'];
    
    if TaskType = 'FindFiles' then
    begin
      // Поиск файлов по маске
      StartPath := FParameters.Values['StartPath'];
      FileMask := FParameters.Values['FileMask'];
      
      if (StartPath <> '') and (FileMask <> '') then
      begin
        FindFiles(StartPath, FileMask);
      end
      else
      begin
        FResults.Add('Ошибка: не указаны обязательные параметры (StartPath или FileMask)');
        FStatus := 3; // Ошибка
        FErrorMessage := 'Не указаны обязательные параметры';
        FSuccess := False;
        Exit;
      end;
    end
    else if TaskType = 'FindFilesMultiMask' then
    begin
      // Поиск файлов по нескольким маскам
      StartPath := FParameters.Values['StartPath'];
      
      // Получаем список масок
      MultiMasks := TStringList.Create;
      try
        // Если параметр MultiMasks передан как строка
        MultiMasks.CommaText := FParameters.Values['MultiMasks'];
        
        if (StartPath <> '') and (MultiMasks.Count > 0) then
        begin
          FindFilesWithMultiMasks(StartPath, MultiMasks);
        end
        else
        begin
          FResults.Add('Ошибка: не указаны обязательные параметры (StartPath или MultiMasks)');
          FStatus := 3; // Ошибка
          FErrorMessage := 'Не указаны обязательные параметры';
          FSuccess := False;
        end;
      finally
        MultiMasks.Free;
      end;
    end
    else if TaskType = 'SearchInFile' then
    begin
      // Поиск строки в файле
      FilePath := FParameters.Values['FilePath'];
      SearchStr := FParameters.Values['SearchString'];
      
      if (FilePath <> '') and (SearchStr <> '') then
      begin
        SearchStringInFile(FilePath, SearchStr);
      end
      else
      begin
        FResults.Add('Ошибка: не указаны обязательные параметры (FilePath или SearchString)');
        FStatus := 3; // Ошибка
        FErrorMessage := 'Не указаны обязательные параметры';
        FSuccess := False;
        Exit;
      end;
    end
    else if TaskType = 'SearchMultipleInFile' then
    begin
      // Поиск нескольких строк в файле
      FilePath := FParameters.Values['FilePath'];
      
      // Получаем список строк для поиска
      MultiSearchStrings := TStringList.Create;
      try
        // Если строки указаны через разделитель
        MultiSearchStrings.CommaText := FParameters.Values['MultiSearchStrings'];
        
        if (FilePath <> '') and (MultiSearchStrings.Count > 0) then
        begin
          SearchMultipleStringsInFile(FilePath, MultiSearchStrings);
        end
        else
        begin
          FResults.Add('Ошибка: не указаны обязательные параметры (FilePath или MultiSearchStrings)');
          FStatus := 3; // Ошибка
          FErrorMessage := 'Не указаны обязательные параметры';
          FSuccess := False;
        end;
      finally
        MultiSearchStrings.Free;
      end;
    end
    else
    begin
      FResults.Add('Ошибка: неизвестный тип задачи - ' + TaskType);
      FStatus := 3; // Ошибка
      FErrorMessage := 'Неизвестный тип задачи';
      FSuccess := False;
      Exit;
    end;
    
    // Завершаем задачу
    if FStatus <> 3 then // Если не было ошибок
    begin
      if Terminated then
      begin
        FStatus := 3; // Ошибка
        FErrorMessage := 'Задача была принудительно остановлена';
        FSuccess := False;
        FResultMessage := 'Задача не завершена';
      end
      else
      begin
        FStatus := 2; // Завершена
        FProgress := 100; // 100%
        FSuccess := True;
        FResultMessage := 'Задача успешно выполнена';
        FResults.Add('Задача успешно завершена: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
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
  // В нашей DLL 2 типа задач
  Result := 2;
end;

// Возвращает информацию о задаче по индексу
function GetTaskInfo(Index: Integer; var TaskInfo: TTaskInfo): Boolean; stdcall;
begin
  Result := False;

  case Index of
    0: // FindFilesMultiMask
      begin
        StrPCopy(TaskInfo.Name, 'FindFilesMultiMask');
        StrPCopy(TaskInfo.Description, 'Поиск файлов по нескольким маскам в указанном каталоге');
        TaskInfo.ParamCount := 3; // StartPath, MultiMasks, IncludeSubfolders
        Result := True;
      end;
    1: // SearchMultipleInFile
      begin
        StrPCopy(TaskInfo.Name, 'SearchMultipleInFile');
        StrPCopy(TaskInfo.Description, 'Поиск вхождений нескольких строк в файле');
        TaskInfo.ParamCount := 2; // FilePath, MultiSearchStrings
        Result := True;
      end;
  end;
end;

// Возвращает информацию о параметре задачи
function GetTaskParamInfo(TaskIndex, ParamIndex: Integer; var ParamInfo: TTaskParamInfo): Boolean; stdcall;
begin
  Result := False;

  case TaskIndex of
    0: // FindFilesMultiMask
      case ParamIndex of
        0: // StartPath
          begin
            StrPCopy(ParamInfo.Name, 'StartPath');
            StrPCopy(ParamInfo.Description, 'Каталог для начала поиска');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
        1: // MultiMasks
          begin
            StrPCopy(ParamInfo.Name, 'MultiMasks');
            StrPCopy(ParamInfo.Description, 'Маски для поиска файлов, разделенные запятыми (например, *.txt,*.doc)');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
        2: // IncludeSubfolders
          begin
            StrPCopy(ParamInfo.Name, 'IncludeSubfolders');
            StrPCopy(ParamInfo.Description, 'Включать подкаталоги в поиск');
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
            StrPCopy(ParamInfo.Description, 'Путь к файлу для поиска');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
        1: // MultiSearchStrings
          begin
            StrPCopy(ParamInfo.Name, 'MultiSearchStrings');
            StrPCopy(ParamInfo.Description, 'Строки для поиска, разделенные запятыми');
            ParamInfo.ParamType := ptString;
            ParamInfo.Required := True;
            Result := True;
          end;
      end;
  end;
end;

// Запускает задачу и возвращает её ID
function StartTask(TaskName: PChar; TaskId: Integer; ParamCount: Integer; ParamNames, ParamValues: Pointer): Boolean; stdcall;
var
  ParamNamesArray: TParamArray;
  ParamValuesArray: TParamValues;
  i: Integer;
begin
  Result := False;

  // Проверяем, создан ли менеджер задач
  if TaskManager = nil then
    Exit;

  // Подготавливаем массивы для параметров
  if ParamCount > 0 then
  begin
    SetLength(ParamNamesArray, ParamCount);
    SetLength(ParamValuesArray, ParamCount);

    // Копируем указатели на имена и значения параметров
    Move(ParamNames^, ParamNamesArray[0], ParamCount * SizeOf(PChar));
    Move(ParamValues^, ParamValuesArray[0], ParamCount * SizeOf(TTaskParamValue));

    // Проверяем, что все параметры верны
    for i := 0 to ParamCount - 1 do
    begin
      if ParamNamesArray[i] = nil then
        Exit; // Если имя параметра не указано, выходим
    end;

    // Запускаем задачу с параметрами
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
    // Без параметров нельзя запустить наши задачи
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