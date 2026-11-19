using System.Globalization;
using System.Text.RegularExpressions;
using FormFiller.Core.Config;

namespace FormFiller.Core.Excel;

/// <summary>
/// 식 블록의 계산기. 사용자가 설정 화면에서 만든 식을 읽어 값(숫자 또는 글자)을 만든다.
///
/// [쓸 수 있는 것]
///   [어제 H20]   어제 파일의 H20 칸
///   [H20]        오늘 파일의 H20 칸 (다른 항목을 다 쓴 뒤의 값)
///   [입력 장비]  오늘 입력한 '장비' 값 (표 = 금일 합계, 명단 = 인원 수)
///   12, 3.5      숫자          "글자"   글자
///   + - * / ( )  사칙연산      &        글자 이어붙이기 (예: "장비 " &amp; [입력 장비] &amp; "대")
///
/// 엑셀 수식처럼 계산하지만 엑셀 파일에는 결과 값만 쓴다 (서식의 다른 칸을 건드리지 않게).
/// </summary>
public static class BlockFormula
{
    public const string KindYesterday = "어제", KindInput = "입력", KindCell = "칸";

    /// <summary>식 안의 참조 하나. Kind = 어제 / 입력 / 칸, Arg = 칸 주소(대문자) 또는 입력 이름</summary>
    public record Ref(string Kind, string Arg);

    private static readonly Regex RefPattern = new(@"\[([^\]]*)\]", RegexOptions.Compiled);
    private static readonly Regex CellPattern = new(@"^[A-Z]{1,3}[0-9]{1,7}$", RegexOptions.Compiled);

    /// <summary>식에 들어 있는 참조들 (어제 파일에서 미리 읽을 칸을 고르는 데 쓴다)</summary>
    public static List<Ref> Refs(string? formula) =>
        string.IsNullOrEmpty(formula) ? []
        : RefPattern.Matches(formula).Select(m => TryParseRef(m.Groups[1].Value, out var r) ? r : null)
                    .Where(r => r != null).Select(r => r!).ToList();

    /// <summary>식이 맞으면 null, 틀리면 이유 (설정 화면에서 바로 보여 준다)</summary>
    public static string? Check(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula)) return "식이 비어 있습니다";
        try { Evaluate(formula, _ => null); return null; }
        catch (FormatException ex) { return ex.Message; }
    }

    /// <summary>
    /// 식을 계산한다. 결과는 decimal(숫자) 또는 string(글자).
    /// resolve: 참조 → 값 (decimal / string / null=빈 칸). 틀린 식이면 FormatException (메시지는 사용자용).
    /// </summary>
    public static object Evaluate(string formula, Func<Ref, object?> resolve) =>
        new Parser(formula, resolve).ParseAll();

    /// <summary>값 → 숫자 (글자면 숫자로 읽어 보고, 안 되면 0)</summary>
    public static decimal ToNumber(object? v) => v switch
    {
        decimal d => d,
        string s  => BlockTypes.ParseNumber(s) ?? 0,
        _         => 0,
    };

    /// <summary>값 → 글자 (숫자는 쓸데없는 0 없이)</summary>
    public static string ToText(object? v) => v switch
    {
        decimal d => d.ToString("0.############", CultureInfo.InvariantCulture),
        string s  => s,
        _         => "",
    };

    private static bool TryParseRef(string inner, out Ref r)
    {
        string s = inner.Trim();
        r = new Ref(KindCell, "");
        if (s.StartsWith(KindYesterday))
        {
            string cell = s[KindYesterday.Length..].Trim().ToUpperInvariant();
            if (!CellPattern.IsMatch(cell)) return false;
            r = new Ref(KindYesterday, cell);
            return true;
        }
        if (s.StartsWith(KindInput))
        {
            string name = s[KindInput.Length..].Trim();
            if (name.Length == 0) return false;
            r = new Ref(KindInput, name);
            return true;
        }
        string c = s.ToUpperInvariant();
        if (!CellPattern.IsMatch(c)) return false;
        r = new Ref(KindCell, c);
        return true;
    }

    // ── 계산기 (재귀 하강) ──────────────────────────────────────────
    //   식   = 합 ( & 합 )*
    //   합   = 곱 ( (+|-) 곱 )*
    //   곱   = 단항 ( (*|/) 단항 )*
    //   단항 = - 단항 | 숫자 | "글자" | [참조] | ( 식 )

    private sealed class Parser(string text, Func<Ref, object?> resolve)
    {
        private int _pos;

        public object ParseAll()
        {
            var v = Concat();
            SkipSpace();
            if (_pos < text.Length) throw Error($"'{text[_pos]}' 은(는) 식에 쓸 수 없습니다");
            return v ?? 0m;
        }

        private object Concat()
        {
            var v = Sum();
            while (Eat('&')) v = ToText(v) + ToText(Sum());
            return v;
        }

        private object Sum()
        {
            var v = Product();
            while (true)
            {
                if (Eat('+')) v = ToNumber(v) + ToNumber(Product());
                else if (Eat('-')) v = ToNumber(v) - ToNumber(Product());
                else return v;
            }
        }

        private object Product()
        {
            var v = Unary();
            while (true)
            {
                if (Eat('*')) v = ToNumber(v) * ToNumber(Unary());
                else if (Eat('/'))
                {
                    decimal d = ToNumber(Unary());
                    v = d == 0 ? 0m : ToNumber(v) / d;   // 0 으로 나누면 0 (빈 칸으로 나누는 날도 있어서)
                }
                else return v;
            }
        }

        private object Unary()
        {
            SkipSpace();
            if (_pos >= text.Length) throw Error("식이 끝났습니다. 뒤에 숫자나 칸이 더 있어야 합니다");
            char c = text[_pos];
            if (c == '-') { _pos++; return -ToNumber(Unary()); }
            if (c == '(')
            {
                _pos++;
                var v = Concat();
                if (!Eat(')')) throw Error("괄호 ')' 가 닫히지 않았습니다");
                return v;
            }
            if (c == '"')
            {
                int end = text.IndexOf('"', _pos + 1);
                if (end < 0) throw Error("글자를 닫는 \" 가 없습니다");
                string s = text[(_pos + 1)..end];
                _pos = end + 1;
                return s;
            }
            if (c == '[')
            {
                int end = text.IndexOf(']', _pos + 1);
                if (end < 0) throw Error("칸을 닫는 ] 가 없습니다");
                string inner = text[(_pos + 1)..end];
                _pos = end + 1;
                if (!TryParseRef(inner, out var r))
                    throw Error($"[{inner}] 을(를) 알 수 없습니다. [어제 H20] / [H20] / [입력 이름] 처럼 칸이나 입력 이름을 넣으세요");
                return resolve(r) ?? "";
            }
            if (char.IsDigit(c) || c == '.')
            {
                int start = _pos;
                while (_pos < text.Length && (char.IsDigit(text[_pos]) || text[_pos] is '.' or ',')) _pos++;
                string num = text[start.._pos].Replace(",", "");
                if (!decimal.TryParse(num, NumberStyles.Number, CultureInfo.InvariantCulture, out var d))
                    throw Error($"'{num}' 은(는) 숫자가 아닙니다");
                return d;
            }
            throw Error($"'{c}' 은(는) 식에 쓸 수 없습니다");
        }

        private bool Eat(char c)
        {
            SkipSpace();
            if (_pos < text.Length && text[_pos] == c) { _pos++; return true; }
            return false;
        }

        private void SkipSpace()
        {
            while (_pos < text.Length && char.IsWhiteSpace(text[_pos])) _pos++;
        }

        private FormatException Error(string message) => new(message);
    }
}
