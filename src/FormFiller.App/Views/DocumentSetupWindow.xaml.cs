using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using FormFiller.App.ViewModels.Settings;
using FormFiller.Core.Config;

namespace FormFiller.App.Views;

/// <summary>
/// DocumentSetupWindow 의 코드-비하인드.
///
/// 역할:
/// 1. PreviewData 가 바뀌면 PreviewGrid 를 동적으로 재생성한다.
/// 2. 셀 클릭 이벤트를 ViewModel 의 OnCellClicked() 로 전달한다.
/// 3. 테이블 설정 버튼 클릭 시 TableConfigDialog 를 연다.
/// 4. 드래그 앤 드롭으로 xlsx 파일을 받는다.
/// </summary>
public partial class DocumentSetupWindow : Window
{
    // 현재 선택된 셀 Border (하이라이트 해제를 위해 추적)
    private Border? _selectedCellBorder;

    // 셀 선택 색상 상수
    private static readonly Color SelectedCellColor  = Color.FromRgb(0xBE, 0xE3, 0xF8); // 연파랑
    private static readonly Color HoverCellColor     = Color.FromRgb(0xEB, 0xF8, 0xFF); // 아주 연파랑
    private static readonly Color HeaderBgColor      = Color.FromRgb(0xED, 0xF2, 0xF7); // 헤더 회색
    private static readonly Color HeaderBorderColor  = Color.FromRgb(0xCB, 0xD5, 0xE0);
    private static readonly Color CellBorderColor    = Color.FromRgb(0xED, 0xF2, 0xF7);

    public DocumentSetupWindow()
    {
        InitializeComponent();

        var vm = (DocumentSetupViewModel)DataContext;

        // PreviewData 가 바뀔 때마다 그리드 재생성
        vm.PropertyChanged += OnViewModelPropertyChanged;
    }

    // ── PropertyChanged 핸들러 ────────────────────────────────────────

    /// <summary>
    /// ViewModel 프로퍼티 변경 감지.
    /// PreviewData 가 null 아닌 값으로 바뀌면 그리드를 재생성한다.
    /// </summary>
    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(DocumentSetupViewModel.PreviewData))
            BuildPreviewGrid();
    }

    // ── 엑셀 미리보기 그리드 동적 생성 ──────────────────────────────

    /// <summary>
    /// PreviewData 를 읽어 PreviewGrid 를 동적으로 생성한다.
    ///
    /// 구조:
    /// - 행 0:        열 헤더 (A, B, C …)
    /// - 행 1~N:      행 번호(1, 2, 3 …) + 셀 값
    /// - 열 0:        행 번호 헤더
    /// - 열 1~M:      셀 데이터
    /// </summary>
    private void BuildPreviewGrid()
    {
        // 기존 셀 초기화
        PreviewGrid.Children.Clear();
        PreviewGrid.RowDefinitions.Clear();
        PreviewGrid.ColumnDefinitions.Clear();
        _selectedCellBorder = null;

        var vm = (DocumentSetupViewModel)DataContext;
        if (vm.PreviewData == null) return;

        int rows = vm.PreviewRowCount;
        int cols = vm.PreviewColCount;
        if (rows <= 0 || cols <= 0) return;

        // ── 행/열 정의 ──
        // 행 0: 열 헤더 (고정 높이)
        PreviewGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(22) });
        for (int r = 0; r < rows; r++)
            PreviewGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(22) });

        // 열 0: 행 번호 헤더 (고정 너비)
        PreviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(36) });
        for (int c = 0; c < cols; c++)
            PreviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });

        // ── 열 헤더 (A, B, C ...) ──
        for (int c = 0; c < cols; c++)
        {
            var header = MakeHeaderCell(ExcelAnalyzer_ColIndexToLetter(c + 1));
            Grid.SetRow(header, 0);
            Grid.SetColumn(header, c + 1);
            PreviewGrid.Children.Add(header);
        }

        // ── 행 번호 + 셀 데이터 ──
        for (int r = 0; r < rows; r++)
        {
            // 행 번호 헤더
            var rowHeader = MakeHeaderCell((r + 1).ToString());
            Grid.SetRow(rowHeader, r + 1);
            Grid.SetColumn(rowHeader, 0);
            PreviewGrid.Children.Add(rowHeader);

            // 셀 데이터
            for (int c = 0; c < cols; c++)
            {
                // 셀 주소 계산: 1-based 열/행 → "A1", "B3" 형식
                string addr  = ExcelAnalyzer_ColIndexToLetter(c + 1) + (r + 1);
                string value = vm.PreviewData[r, c] ?? "";

                var cell = MakeDataCell(addr, value);
                Grid.SetRow(cell, r + 1);
                Grid.SetColumn(cell, c + 1);
                PreviewGrid.Children.Add(cell);
            }
        }
    }

    /// <summary>헤더 셀(배경 회색, 볼드) 을 생성한다.</summary>
    private static Border MakeHeaderCell(string text)
    {
        return new Border
        {
            Background      = new SolidColorBrush(HeaderBgColor),
            BorderBrush     = new SolidColorBrush(HeaderBorderColor),
            BorderThickness = new Thickness(0, 0, 1, 1),
            Child = new TextBlock
            {
                Text                = text,
                FontSize            = 11,
                FontWeight          = FontWeights.SemiBold,
                Foreground          = new SolidColorBrush(Color.FromRgb(0x4A, 0x55, 0x68)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            }
        };
    }

    /// <summary>
    /// 클릭 가능한 데이터 셀 Border 를 생성한다.
    /// Tag 에 셀 주소를 저장해 클릭 핸들러에서 사용한다.
    /// </summary>
    private Border MakeDataCell(string address, string value)
    {
        var tb = new TextBlock
        {
            Text            = value,
            FontSize        = 11,
            Padding         = new Thickness(3, 0, 3, 0),
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming    = TextTrimming.CharacterEllipsis,
            // 긴 값은 ToolTip 으로 전체 내용 표시
            ToolTip         = string.IsNullOrEmpty(value) ? null : (object)value,
        };

        var border = new Border
        {
            Tag             = address, // 셀 주소를 태그에 보관
            Background      = Brushes.White,
            BorderBrush     = new SolidColorBrush(CellBorderColor),
            BorderThickness = new Thickness(0, 0, 1, 1),
            Cursor          = Cursors.Hand,
            Child           = tb,
        };

        // 이벤트 연결
        border.MouseLeftButtonDown += DataCell_Click;
        border.MouseEnter          += DataCell_MouseEnter;
        border.MouseLeave          += DataCell_MouseLeave;

        return border;
    }

    // ── 셀 이벤트 핸들러 ─────────────────────────────────────────────

    /// <summary>
    /// 데이터 셀 클릭 시:
    /// 1. 이전 선택 셀 하이라이트 해제
    /// 2. 새 선택 셀 하이라이트
    /// 3. ViewModel 에 주소/값 전달 (피커 모드이면 필드에 자동 입력됨)
    /// </summary>
    private void DataCell_Click(object sender, MouseButtonEventArgs e)
    {
        var border = (Border)sender;

        // 이전 선택 해제
        if (_selectedCellBorder != null)
            _selectedCellBorder.Background = Brushes.White;

        // 새 선택 하이라이트
        border.Background = new SolidColorBrush(SelectedCellColor);
        _selectedCellBorder = border;

        string address = (string)border.Tag;
        string value   = (border.Child is TextBlock tb) ? tb.Text : "";

        // ViewModel 에 클릭 전달
        var vm = (DocumentSetupViewModel)DataContext;
        vm.OnCellClicked(address, value);
    }

    /// <summary>마우스 진입 시 hover 색상 (선택된 셀이면 무시)</summary>
    private void DataCell_MouseEnter(object sender, MouseEventArgs e)
    {
        var border = (Border)sender;
        // 선택된 셀은 hover 색상으로 덮지 않는다
        if (ReferenceEquals(border, _selectedCellBorder)) return;
        border.Background = new SolidColorBrush(HoverCellColor);
    }

    /// <summary>마우스 이탈 시 원래 흰색으로 복원 (선택된 셀이면 무시)</summary>
    private void DataCell_MouseLeave(object sender, MouseEventArgs e)
    {
        var border = (Border)sender;
        if (ReferenceEquals(border, _selectedCellBorder)) return;
        border.Background = Brushes.White;
    }

    // ── 테이블 설정 버튼 ─────────────────────────────────────────────

    /// <summary>
    /// 테이블 필드의 "설정" 버튼 클릭 시 간이 설정 다이얼로그를 표시한다.
    /// Button.Tag 에 FieldAssignRow 가 바인딩되어 있다.
    /// </summary>
    private void TableSetup_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        if (btn.Tag is not FieldAssignRow field) return;

        // 테이블 타입에 따라 입력 항목 구성
        bool isEquip = field.Type == FieldTypes.EquipmentTable;

        string msg = isEquip
            ? $"장비 테이블 설정\n\n시작행: {field.StartRow}\n끝행: {field.EndRow}\n장비명열: {field.NameCol}\n전일열: {field.PrevDayCol}\n금일열: {field.TodayCol}"
            : $"인원 테이블 설정\n\n시작행: {field.StartRow}\n끝행: {field.EndRow}\n직종열: {field.RoleCol}\n전일누계열: {field.PrevCumulativeCol}\n금일열: {field.TodayCol}";

        // 현재 값을 보여 주고 편집 안내
        MessageBox.Show(
            msg + "\n\n※ 테이블 범위는 오른쪽 패널에서 직접 수정하세요.",
            "테이블 설정 확인",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    // ── 드래그 앤 드롭 ───────────────────────────────────────────────

    /// <summary>파일 드래그 중: 드롭 가능 여부를 커서로 알린다.</summary>
    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    /// <summary>파일 드롭: ViewModel 의 HandleDroppedFile 에 전달한다.</summary>
    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
        {
            var vm = (DocumentSetupViewModel)DataContext;
            vm.HandleDroppedFile(files);
        }
    }

    // ── 유틸리티 ─────────────────────────────────────────────────────

    /// <summary>
    /// 1-based 열 인덱스를 열 문자열로 변환한다.
    /// 예) 1→"A", 26→"Z", 27→"AA", 28→"AB"
    /// </summary>
    private static string ExcelAnalyzer_ColIndexToLetter(int colIndex)
    {
        string result = "";
        while (colIndex > 0)
        {
            int rem = (colIndex - 1) % 26;
            result   = (char)('A' + rem) + result;
            colIndex = (colIndex - 1) / 26;
        }
        return result;
    }
}
