using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Models;
using FormFiller.Core.Phone;

namespace FormFiller.App.ViewModels;

/// <summary>
/// [📱 폰 입력 불러오기]: 폰 앱이 드라이브 inbox 폴더에 올린 입력을 지금 화면에 합친다.
/// partial class: 한 클래스를 여러 파일에 나눠 쓰는 C# 기능. MainViewModel 이 커서 폰 관련 코드만 이 파일로 뺐다.
/// 읽기·검증·합치기 규칙은 Core(PhoneInbox, PhoneInputMerger, PhonePhotoMatcher)에 있고, 여기서는 화면에 넣기만 한다.
/// </summary>
public partial class MainViewModel
{
    private string _phoneStatusText = "";
    /// <summary>"폰 입력 2건 있음" 같은 안내 (없으면 빈 글자)</summary>
    public string PhoneStatusText { get => _phoneStatusText; private set => SetField(ref _phoneStatusText, value); }

    private bool _isLoadingPhone;
    public bool IsLoadingPhone
    {
        get => _isLoadingPhone;
        private set { if (SetField(ref _isLoadingPhone, value)) CommandManager.InvalidateRequerySuggested(); }
    }

    /// <summary>폰이 보낸 오늘 출근자 명단. TBM·노임대장 연동(기획 5단계)에서 쓴다. 지금은 보관만.</summary>
    public IReadOnlyList<PhoneAttendee> PhoneAttendees { get; private set; } = [];

    public ICommand LoadPhoneInputCommand { get; private set; } = null!;

    private int _phoneStatusVersion;

    private void InitPhone()
    {
        LoadPhoneInputCommand = new RelayCommand(_ => LoadPhoneInput(), _ => !IsLoadingPhone);
        RefreshPhoneStatus();
    }

    private static string InboxFolder() =>
        PhoneInbox.ResolveInboxFolder(AppConfigLoader.Load().PhoneInboxFolder, AppDomain.CurrentDomain.BaseDirectory);

    /// <summary>날짜가 바뀌면 그날 폰 입력이 있는지 확인해 안내한다. 드라이브 폴더라 뒤 스레드에서 찾는다.</summary>
    private async void RefreshPhoneStatus()
    {
        int version = ++_phoneStatusVersion;
        DateTime date = TargetDate;
        try
        {
            string inbox = InboxFolder();
            int count = await Task.Run(() => PhoneInbox.FindSubmissions(inbox, date).Count);
            if (version != _phoneStatusVersion) return;
            PhoneStatusText = count == 0 ? "" : $"📱 {date:MM/dd} 폰 입력 {count}건 있음 — [폰 입력 불러오기]를 누르세요";
        }
        catch (Exception ex)
        {
            AppLog.Error("폰 입력 확인 실패", ex);
            if (version == _phoneStatusVersion) PhoneStatusText = "";
        }
    }

    private async void LoadPhoneInput()
    {
        if (IsLoadingPhone) return;
        DateTime date  = TargetDate.Date;
        string   inbox = InboxFolder();

        IReadOnlyList<string> folders;
        PhoneReadResult read;
        IsLoadingPhone = true;
        try
        {
            (folders, read) = await Task.Run(() =>
            {
                var found = PhoneInbox.FindSubmissions(inbox, date);
                return (found, found.Count == 0 ? new PhoneReadResult(null, []) : PhoneInbox.Read(found[^1]));
            });
        }
        catch (Exception ex)
        {
            AppLog.Error($"폰 입력 읽기 실패: {inbox}", ex);
            Show($"폰 입력 폴더를 읽지 못했습니다.\n{inbox}\n\n{ex.Message}", MessageBoxImage.Warning);
            return;
        }
        finally { IsLoadingPhone = false; }

        if (folders.Count == 0)
        {
            Show($"{date:yyyy-MM-dd} 에 폰에서 보낸 입력이 없습니다.\n\n찾은 폴더: {inbox}\n" +
                 "(폰에서 보낸 직후라면 드라이브 동기화가 끝날 때까지 잠시 기다려 주세요)", MessageBoxImage.Information);
            return;
        }
        if (read.Input == null)
        {
            AppLog.Info($"폰 입력 검증 실패: {Path.GetFileName(folders[^1])} — {string.Join(" / ", read.Errors)}");
            Show($"폰 입력을 쓸 수 없습니다 ({Path.GetFileName(folders[^1])}).\n\n" +
                 string.Join("\n", read.Errors.Take(10).Select(e => "• " + e)), MessageBoxImage.Warning);
            return;
        }
        if (read.Input.Date.Date != TargetDate.Date) return;   // 읽는 사이 날짜를 바꿈

        ConfirmAndApply(read.Input, folders);
    }

    private void ConfirmAndApply(PhoneInput phone, IReadOnlyList<string> folders)
    {
        var merge   = PhoneInputMerger.Apply(BuildDailyInput(), phone);
        var targets = PhotoTargets();
        var photos  = PhonePhotoMatcher.Match(targets.Select(t => t.Label).ToList(), phone.Photos);

        if (MessageBox.Show(Summary(phone, merge, photos, targets, folders), "폰 입력 불러오기",
                            MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK)
            return;

        ApplyMerged(merge.Input, phone);
        foreach (var (index, photo) in photos.Placed) targets[index].Set(photo);
        PhoneAttendees = phone.Attendees;
        AppLog.Info($"폰 입력 적용: {Path.GetFileName(folders[^1])} (사진 {photos.Placed.Count}장, 충돌 {merge.Conflicts.Count}건)");
    }

    private static string Summary(PhoneInput phone, PhoneMergeResult merge, PhotoMatchResult photos,
                                  List<PhotoTarget> targets, IReadOnlyList<string> folders)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"{Path.GetFileName(folders[^1])} 의 입력을 화면에 넣습니다.");
        if (folders.Count > 1)
            sb.AppendLine($"(같은 날 {folders.Count}번 보내서 마지막 것을 씁니다. 앞의 것: " +
                          string.Join(", ", folders.SkipLast(1).Select(Path.GetFileName)) + ")");
        sb.AppendLine();
        if (phone.Weather.Length > 0) sb.AppendLine($"• 날씨: {phone.Weather}");
        int workLines = phone.TodayMorningWork.Count + phone.TodayAfternoonWork.Count + phone.TomorrowWork.Count;
        if (workLines > 0) sb.AppendLine($"• 작업사항 {workLines}줄 (폰이 보낸 칸만 바꿈)");
        sb.AppendLine($"• 장비 {phone.Equipment.Count}줄, 인원 {phone.Personnel.Count}줄, 수장재 {phone.SujangjaePersonnel.Count}줄의 금일 값");
        sb.AppendLine($"• 사진 {photos.Placed.Count}/{phone.Photos.Count}장: " +
                      string.Join(", ", photos.Placed.Select(p => $"{p.Photo.Label}→{targets[p.TargetIndex].Label}")));
        if (phone.Attendees.Count > 0) sb.AppendLine($"• 출근자 {phone.Attendees.Count}명 (보관만, TBM 연동은 나중에)");

        AppendList(sb, "⚠ 화면 값과 달라서 폰 값으로 바꿀 것", merge.Conflicts);
        AppendList(sb, "＋ 화면에 없어서 새 줄로 넣을 것", merge.Unmatched);
        AppendList(sb, "✖ 넣을 칸을 못 찾은 사진 (라벨이 사진 칸 이름과 달라요)",
                   photos.Unplaced.Select(p => $"{p.File} ({(p.Label.Length > 0 ? p.Label : "라벨 없음")})").ToList());
        sb.AppendLine();
        sb.Append("[확인]을 누르면 넣습니다. 넣은 뒤에도 화면에서 고칠 수 있습니다.");
        return sb.ToString();
    }

    private static void AppendList(StringBuilder sb, string title, IReadOnlyList<string> lines)
    {
        if (lines.Count == 0) return;
        sb.AppendLine().AppendLine(title);
        foreach (var l in lines.Take(8)) sb.AppendLine("   " + l);
        if (lines.Count > 8) sb.AppendLine($"   … 외 {lines.Count - 8}건");
    }

    // ── 화면에 넣기 ───────────────────────────────────────────────────

    private void ApplyMerged(DailyInput merged, PhoneInput phone)
    {
        Weather = merged.Weather;
        if (phone.TodayMorningWork.Count > 0)   ReplaceWork(TodayMorningItems, phone.TodayMorningWork);
        if (phone.TodayAfternoonWork.Count > 0) ReplaceWork(TodayAfternoonItems, phone.TodayAfternoonWork);
        if (phone.TomorrowWork.Count > 0)       ReplaceWork(TomorrowWorkItems, phone.TomorrowWork);

        // 화면 줄은 그대로 두고 금일 값만 바꾼다 (전일·[전일과 같게] 값을 지키려고). 없는 줄만 새로 넣는다.
        foreach (var e in merged.Equipment)
        {
            var vm = Equipment.FirstOrDefault(x => NormKey(x.Name) == NormKey(e.Name)) ?? AddEquipment(e.Name, e.Spec);
            vm.TodayCount = e.TodayCount;
        }
        ApplyPersonnel(Personnel, merged.Personnel);
        ApplyPersonnel(SujangjaePersonnel, merged.SujangjaePersonnel);
    }

    private void ApplyPersonnel(ObservableCollection<PersonnelEntryViewModel> rows, List<PersonnelEntry> merged)
    {
        foreach (var p in merged)
        {
            var vm = rows.FirstOrDefault(x => NormKey(x.Role) == NormKey(p.Role)) ?? AddPersonnel(rows, p.Role);
            vm.TodayCount = p.TodayCount;
        }
    }

    private static void ReplaceWork(ObservableCollection<WorkItemViewModel> items, IReadOnlyList<string> lines)
    {
        items.Clear();
        foreach (var line in lines) items.Add(CreateWorkItem(line, items));
    }

    /// <summary>폰 사진을 넣을 수 있는 칸: 보이는(오늘 만들 문서의) 사진대지 칸 + 문서별 입력의 사진 칸</summary>
    private sealed record PhotoTarget(string Label, Action<PhonePhoto> Set);

    private List<PhotoTarget> PhotoTargets()
    {
        var list = PhotoSlots.Where(s => s.IsVisible)
                             .Select(s => new PhotoTarget(s.Label, p => s.SelectedImagePath = p.FullPath))
                             .ToList();
        foreach (var card in DocInputCards.Where(c => c.IsVisible))
            foreach (var sheet in card.Sheets)
                foreach (var item in sheet.Items.Where(i => i.IsPhoto))
                {
                    // 시트가 여러 개면 "시트 이름 + 칸 이름" 으로 구분 (예: 폰 라벨 "오전초등학교 측정사진")
                    string label = card.Sheets.Count > 1 ? $"{sheet.Header} {item.Label}" : item.Label;
                    list.Add(new PhotoTarget(label, p =>
                    {
                        item.SetPhoto(p.FullPath);
                        if (p.Value.Length > 0 && item.HasValue) item.Value = p.Value;
                    }));
                }
        return list;
    }

    private static void Show(string message, MessageBoxImage icon) =>
        MessageBox.Show(message, "폰 입력 불러오기", MessageBoxButton.OK, icon);
}
