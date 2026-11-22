using FluentAssertions;
using FormFiller.Core.Config;

namespace FormFiller.Tests.Core;

/// <summary>
/// 2026-11-23 화면 정리 3단계: 오늘 작성 화면에서 예전 고정 입력(작업사항·장비·인원·날씨)을 뺐다.
/// 그 입력이 필요한 예전 항목이 남은 문서는 화면에서 알려 줘야 한다.
/// </summary>
public class LegacyInputTests
{
    [Theory]
    [InlineData(FieldTypes.Multiline)]
    [InlineData(FieldTypes.EquipmentTable)]
    [InlineData(FieldTypes.PersonnelTable)]
    [InlineData(FieldTypes.TbmPersonnel)]
    [InlineData(FieldTypes.ActiveEquipmentNames)]
    [InlineData(FieldTypes.Text)]
    public void 예전_입력_화면이_필요한_종류(string type) =>
        FieldTypes.NeedsRemovedInput(type).Should().BeTrue();

    [Theory]
    [InlineData(BlockTypes.Lines)]
    [InlineData(BlockTypes.Table)]
    [InlineData(BlockTypes.Value)]
    [InlineData(FieldTypes.Photo)]          // 사진·측정값 단계는 그대로 있다
    [InlineData(FieldTypes.DateKorean)]     // 날짜는 자동
    [InlineData(FieldTypes.AttendanceRate)] // 자동 100
    public void 블록_사진_자동값은_아니다(string type) =>
        FieldTypes.NeedsRemovedInput(type).Should().BeFalse();

    [Fact]
    public void 문서에_하나라도_있으면_true()
    {
        var doc = new DocumentConfig
        {
            Fields = [new FieldConfig { Type = BlockTypes.Date }, new FieldConfig { Type = FieldTypes.EquipmentTable }],
        };
        doc.HasRemovedInputFields().Should().BeTrue();
        new DocumentConfig { Fields = [new FieldConfig { Type = BlockTypes.Date }] }.HasRemovedInputFields().Should().BeFalse();
    }
}
