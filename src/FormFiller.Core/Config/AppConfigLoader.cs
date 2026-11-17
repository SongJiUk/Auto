using System.Text.Json;

namespace FormFiller.Core.Config;

/// <summary>
/// config.json 을 읽어 AppConfig 객체로 반환하는 로더.
/// 파일이 없거나 역직렬화에 실패하면 CreateDefault() 로 기본값을 반환한다.
/// </summary>
public static class AppConfigLoader
{
    // JSON 역직렬화 옵션: 대소문자 구분 없이 프로퍼티 이름 매핑
    private static readonly JsonSerializerOptions ReadOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    // JSON 직렬화 옵션: 들여쓰기 적용 (사람이 읽기 좋은 형식)
    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        WriteIndented = true,
    };

    /// <summary>
    /// exe 위치(AppDomain.CurrentDomain.BaseDirectory) 기준으로 config.json 을 읽는다.
    /// 파일이 없거나 파싱 실패 시 기본값을 반환한다.
    /// </summary>
    // 마지막으로 읽은 config.json 내용. 파일 시각·크기가 같으면 디스크(구글 드라이브)에서 다시 읽지 않는다.
    // 객체가 아니라 글자를 담아 두므로, 부를 때마다 새 객체를 돌려주는 동작은 그대로다.
    private static readonly object CacheGate = new();
    private static string?  _cachedJson;
    private static DateTime _cachedWriteTime;
    private static long     _cachedLength = -1;

    private static string? ReadJsonCached(string configPath)
    {
        var fi = new FileInfo(configPath);
        if (!fi.Exists) return null;
        lock (CacheGate)
        {
            if (_cachedJson != null && fi.LastWriteTimeUtc == _cachedWriteTime && fi.Length == _cachedLength)
                return _cachedJson;
            string json = File.ReadAllText(configPath, System.Text.Encoding.UTF8);
            _cachedJson      = json;
            _cachedWriteTime = fi.LastWriteTimeUtc;
            _cachedLength    = fi.Length;
            return json;
        }
    }

    public static AppConfig Load()
    {
        string configPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "config.json");

        string? json = null;
        try { json = ReadJsonCached(configPath); }
        catch (IOException ex) { AppLog.Error("config.json 을 읽지 못했습니다", ex); }

        if (json != null)
        {
            try
            {
                var    config = JsonSerializer.Deserialize<AppConfig>(json, ReadOptions);
                if (config != null)
                    return RemapDrivePaths(config);
            }
            catch (JsonException ex)
            {
                // JSON 파싱 실패 시 기본값으로 폴백 (운영 중 설정 오타 대비)
                AppLog.Error("config.json 내용이 잘못되어 기본 설정을 씁니다", ex);
            }
        }

        // 파일 없거나 파싱 실패 → 기본값 반환
        return CreateDefault();
    }

    /// <summary>
    /// AppConfig 를 exe 위치의 config.json 으로 저장한다 (덮어쓰기).
    /// </summary>
    public static void Save(AppConfig config)
    {
        string path = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "config.json");

        string json = JsonSerializer.Serialize(config, WriteOptions);
        lock (CacheGate)
        {
            File.WriteAllText(path, json, System.Text.Encoding.UTF8);
            _cachedJson = null;   // 다음 Load 때 새로 읽어 시각·크기를 다시 기억한다
        }
    }

    /// <summary>
    /// 다른 PC 에서 고른 드라이브 폴더 경로(예: 회사 G:\내 드라이브\...)를 이 PC 경로(예: 집 H:\내 드라이브\...)로 바꾼다.
    /// 경로가 이미 있으면 그대로 둔다. 저장할 때는 바뀐 경로가 저장되지만, 다른 PC 에서 읽을 때 다시 바뀌므로 문제없다.
    /// </summary>
    private static AppConfig RemapDrivePaths(AppConfig config)
    {
        if (Path.IsPathRooted(config.OutputFolder))
            config.OutputFolder = GoogleDrive.Remap(config.OutputFolder);
        foreach (var doc in config.Documents)
            doc.PhotoFolder = GoogleDrive.Remap(doc.PhotoFolder);
        return config;
    }

    /// <summary>
    /// 고른 폴더를 설정에 저장할 모양으로 바꾼다. 프로그램 폴더 안이면 상대 경로(예: "output"), 밖이면 전체 경로.
    /// </summary>
    public static string ToStoredFolderPath(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder)) return folder;
        string baseDir = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
        string full    = Path.GetFullPath(folder);
        return full.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase)
            ? Path.GetRelativePath(baseDir, full)
            : full;
    }

    // ── 서식 경로 (두 PC 에서 같은 설정을 쓰기 위해) ───────────────────

    /// <summary>서식 파일을 모아 두는 폴더 이름 (exe 옆)</summary>
    public const string TemplateFolder = "서식";

    /// <summary>상대 경로면 exe 위치 기준 절대 경로로 바꾼다.</summary>
    public static string ResolvePath(string path) =>
        string.IsNullOrWhiteSpace(path) || Path.IsPathRooted(path)
            ? path
            : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);

    /// <summary>
    /// 설정에 저장할 서식 경로를 만든다.
    /// exe 폴더 밖의 파일이면 exe 옆 "서식" 폴더로 복사하고, 항상 exe 기준 상대 경로를 돌려준다.
    /// 그래서 프로그램 폴더째 구글 드라이브에 두면 다른 PC(드라이브 문자가 달라도)에서 그대로 열린다.
    /// </summary>
    public static string ToStoredTemplatePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return path;
        string baseDir = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
        string full = Path.GetFullPath(ResolvePath(path));

        if (!full.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase))
        {
            if (!File.Exists(full)) return path;   // 없는 파일은 그대로 둔다
            string dir = Path.Combine(baseDir, TemplateFolder);
            Directory.CreateDirectory(dir);
            string dest = Path.Combine(dir, Path.GetFileName(full));
            // 원본은 건드리지 않고 복사만 한다. 같은 이름이 있으면 새 것으로 덮어쓴다.
            File.Copy(full, dest, overwrite: true);
            full = dest;
        }
        return Path.GetRelativePath(baseDir, full);
    }

    // ── private 헬퍼 ─────────────────────────────────────────────────

    /// <summary>
    /// config.json 이 없을 때 사용할 기본 AppConfig.
    /// 공사일보_출력일보, TBM일지_사진대지 두 문서를 하드코딩한 값으로 초기화한다.
    /// </summary>
    private static AppConfig CreateDefault() => new()
    {
        OutputFolder = "output",
        Documents =
        [
            new DocumentConfig
            {
                Id            = "daily_report",
                Name          = "공사일보_출력일보",
                TemplateFile  = "samples/공사일보_출력일보_26.11.04.xlsx",
                OutputPattern = "공사일보_출력일보_{yy}.{MM}.{dd}.xlsx",
                SheetName     = "공사일보",
                Fields =
                [
                    new FieldConfig
                    {
                        Id    = "date",
                        Label = "날짜",
                        Type  = FieldTypes.DateKorean,
                        Cell  = "C6",
                    },
                    new FieldConfig
                    {
                        Id    = "weather",
                        Label = "날씨",
                        Type  = FieldTypes.Text,
                        Cell  = "L6",
                    },
                    new FieldConfig
                    {
                        Id    = "today_work",
                        Label = "금일 작업사항",
                        Type  = FieldTypes.Multiline,
                        Cells = BuildCellRange("A", 9, 17),
                    },
                    new FieldConfig
                    {
                        Id    = "tomorrow_work",
                        Label = "명일 작업사항",
                        Type  = FieldTypes.Multiline,
                        Cells = BuildCellRange("L", 9, 17),
                    },
                    new FieldConfig
                    {
                        Id         = "equipment_table",
                        Label      = "장비 투입현황",
                        Type       = FieldTypes.EquipmentTable,
                        StartRow   = 20,
                        EndRow     = 32,
                        NameCol    = "A",
                        PrevDayCol = "F",
                        TodayCol   = "H",
                    },
                    new FieldConfig
                    {
                        Id               = "personnel_table",
                        Label            = "인원 투입현황",
                        Type             = FieldTypes.PersonnelTable,
                        StartRow         = 20,
                        EndRow           = 32,
                        RoleCol          = "L",
                        PrevCumulativeCol = "N",
                        TodayCol         = "P",
                    },
                ],
            },
            new DocumentConfig
            {
                Id            = "tbm",
                Name          = "TBM일지_사진대지",
                TemplateFile  = "samples/TBM일지_사진대지_26.11.04.xlsx",
                OutputPattern = "TBM일지_사진대지_{yy}.{MM}.{dd}.xlsx",
                SheetName     = "TBM일지 서식11월",
                Fields =
                [
                    new FieldConfig
                    {
                        Id    = "date",
                        Label = "작성일자",
                        Type  = FieldTypes.DateSimple,
                        Cell  = "F4",
                    },
                    new FieldConfig
                    {
                        Id    = "personnel",
                        Label = "인원",
                        Type  = FieldTypes.TbmPersonnel,
                        Cell  = "AD6",
                    },
                    new FieldConfig
                    {
                        Id    = "equipment_names",
                        Label = "장비",
                        Type  = FieldTypes.ActiveEquipmentNames,
                        Cell  = "F7",
                    },
                    new FieldConfig
                    {
                        Id    = "planned_personnel",
                        Label = "계획인원",
                        Type  = FieldTypes.TbmPersonnel,
                        Cell  = "I16",
                    },
                    new FieldConfig
                    {
                        Id    = "actual_personnel",
                        Label = "실시인원",
                        Type  = FieldTypes.TbmPersonnel,
                        Cell  = "O16",
                    },
                    new FieldConfig
                    {
                        Id    = "attendance_rate",
                        Label = "참석율",
                        Type  = FieldTypes.AttendanceRate,
                        Cell  = "U16",
                    },
                ],
            },
        ],
    };

    /// <summary>
    /// 단일 열(col)에 대해 startRow~endRow 범위의 셀 주소 목록을 만든다.
    /// 예) BuildCellRange("A", 9, 17) → ["A9","A10",...,"A17"]
    /// </summary>
    private static List<string> BuildCellRange(string col, int startRow, int endRow)
        => Enumerable.Range(startRow, endRow - startRow + 1)
                     .Select(r => $"{col}{r}")
                     .ToList();
}
