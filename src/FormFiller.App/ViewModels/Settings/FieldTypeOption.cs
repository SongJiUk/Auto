using FormFiller.Core.Config;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 필드 타입 선택 콤보박스에 표시되는 항목.
/// Label 은 사용자에게 보이는 한국어 이름, TypeId 는 FieldTypes 상수.
/// </summary>
public class FieldTypeOption
{
    public string Label  { get; }
    public string TypeId { get; }

    public FieldTypeOption(string label, string typeId)
    {
        Label  = label;
        TypeId = typeId;
    }

    public override string ToString() => Label;

    /// <summary>전체 선택 가능한 타입 목록 (콤보박스 아이템 소스)</summary>
    public static readonly IReadOnlyList<FieldTypeOption> All =
    [
        new("날짜 (한국식)  예) 2026년 11월 05일 목요일", FieldTypes.DateKorean),
        new("날짜 (숫자식)  예) 2026.11.05",              FieldTypes.DateSimple),
        new("텍스트",                                      FieldTypes.Text),
        new("숫자",                                        FieldTypes.Number),
        new("여러줄 텍스트 (작업사항 등)",                FieldTypes.Multiline),
        new("장비 투입현황 표",                            FieldTypes.EquipmentTable),
        new("인원 투입현황 표",                            FieldTypes.PersonnelTable),
        new("투입 장비명 (금일>0 자동 연결)",              FieldTypes.ActiveEquipmentNames),
        new("TBM 인원",                                    FieldTypes.TbmPersonnel),
        new("참석율 (항상 100)",                           FieldTypes.AttendanceRate),
    ];
}
