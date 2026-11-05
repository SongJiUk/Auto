// 이 파일은 AppConfigLoader.cs 로 대체되었습니다.
// 빌드 참조 오류 방지를 위해 namespace 와 빈 클래스 선언만 남깁니다.

namespace FormFiller.Core.Config;

#pragma warning disable CS1591
[Obsolete("AppConfigLoader 로 대체되었습니다.")]
public static class ConfigLoader
{
    [Obsolete("AppConfigLoader.Load() 를 사용하세요.")]
    public static FormConfig LoadFromBaseDirectory() => new();

    [Obsolete("AppConfigLoader.Load() 를 사용하세요.")]
    public static FormConfig Load(string configPath) => new();

    [Obsolete("AppConfigLoader.Save() 를 사용하세요.")]
    public static void Save(FormConfig config) { }
}
#pragma warning restore CS1591
