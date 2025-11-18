(DateText as text, optional locale as nullable text) as nullable date =>
let
    // 0. ตัด ordinal suffixes st/nd/rd/th ที่มาต่อท้ายตัวเลข (กรณีตามด้วย space)
    OrdinalDelims = {"st ", "nd ", "rd ", "th "},
    RemoveOrdinal = List.Accumulate(
        OrdinalDelims,
        DateText,
        (state, s) => Text.Replace(state, s, " ")
    ),

    // 1. ตาราง mapping ไทย→อังกฤษ (full name + abbr. มี/ไม่มีจุด)
    MonthTH2EN = {
        {"มกราคม", "Jan"}, {"ม.ค.", "Jan"}, {"ม.ค", "Jan"}, {"มค", "Jan"},
        {"กุมภาพันธ์", "Feb"}, {"ก.พ.", "Feb"}, {"ก.พ", "Feb"}, {"กพ", "Feb"},
        {"มีนาคม", "Mar"},    {"มี.ค.", "Mar"}, {"มี.ค", "Mar"}, {"มีค", "Mar"},
        {"เมษายน", "Apr"},    {"เม.ย.", "Apr"}, {"เม.ย", "Apr"}, {"เมย", "Apr"},
        {"พฤษภาคม", "May"},  {"พ.ค.", "May"}, {"พ.ค", "May"}, {"พค", "May"},
        {"มิถุนายน", "Jun"},  {"มิ.ย.", "Jun"}, {"มิ.ย", "Jun"}, {"มิย", "Jun"},
        {"กรกฎาคม", "Jul"},  {"ก.ค.", "Jul"}, {"ก.ค", "Jul"}, {"กค", "Jul"},
        {"สิงหาคม", "Aug"},   {"ส.ค.", "Aug"}, {"ส.ค", "Aug"}, {"สค", "Aug"},
        {"กันยายน", "Sep"},   {"ก.ย.", "Sep"}, {"ก.ย", "Sep"}, {"กย", "Sep"},
        {"ตุลาคม", "Oct"},    {"ต.ค.", "Oct"}, {"ต.ค", "Oct"}, {"ตค", "Oct"},
        {"พฤศจิกายน", "Nov"}, {"พ.ย.", "Nov"}, {"พ.ย", "Nov"}, {"พย", "Nov"},
        {"ธันวาคม", "Dec"},   {"ธ.ค.", "Dec"}, {"ธ.ค", "Dec"}, {"ธค", "Dec"}
    },

    // 2. แปลงชื่อเดือนไทย → อังกฤษ (ใช้ RemoveOrdinal แทน DateText)
    ReplaceMonthTH = List.Accumulate(
        MonthTH2EN,
        RemoveOrdinal,
        (state, m) => Text.Replace(state, m{0}, m{1})
    ),

    // 3. แทนตัวคั่นทุกชนิด (รวม . , space, -, _ , /) เป็น "/"
    delimiters = {"-", ",", "_", " ", "/", "."},
    NormalizeDelimiter = List.Accumulate(
        delimiters,
        ReplaceMonthTH,
        (state, d) => Text.Replace(state, d, "/")
    ),

    // 3.1 รายชื่อวัน (ไทย/อังกฤษ) ที่จะตัดทิ้ง (ทั้งหมดเป็นตัวพิมพ์เล็ก เพราะเดี๋ยวเราจะ Text.Lower ก่อน)
    WeekdayTokens = {
        // ไทยเต็ม
        "วันอาทิตย์", "วันจันทร์", "วันอังคาร", "วันพุธ",
        "วันพฤหัส", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์",
        "อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัส", "พฤหัสบดี", "ศุกร์", "เสาร์",

        // ไทยย่อ (มี/ไม่มีจุด)
        "อา", "จ", "อ", "พ", "พฤ", "ศ", "ส",
        "อา.", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.",

        // อังกฤษเต็ม
        "monday", "tuesday", "wednesday", "thursday",
        "friday", "saturday", "sunday",

        // อังกฤษย่อ (มี/ไม่มีจุด)
        "mon", "tue", "tues", "wed", "thu", "thur", "thurs", "fri", "sat", "sun",
        "mon.", "tue.", "tues.", "wed.", "thu.", "thur.", "thurs.", "fri.", "sat.", "sun."
    },

    // 4. split แล้วเอาค่าว่าง/null + ชื่อวัน ออก
    RawTokens =
        List.Transform(
            Text.Split(Text.Lower(NormalizeDelimiter), "/"),
            Text.Trim
        ),

    TokenList =
        List.Select(
            RawTokens,
            each _ <> "" and _ <> null and not List.Contains(WeekdayTokens, _)
        ),

    PreprocessedText = Text.Combine(TokenList, "/"),

    // 5. ลอง parse ด้วย locale ต่างๆ
    TryTH = try Date.FromText(PreprocessedText, "th-TH") otherwise null,
    TryUK = try Date.FromText(PreprocessedText, "en-GB") otherwise null,
    TryUS = try Date.FromText(PreprocessedText, "en-US") otherwise null,

    // 6. สร้าง MakeSenseDate: เลือกวันที่ห่างจากวันนี้น้อยสุด
    CurrentDate = DateTime.Date(DateTime.LocalNow()),
    CandidateDates = List.RemoveNulls({TryTH, TryUK, TryUS}),
    DateWithDiff = List.Transform(
        CandidateDates,
        each {_, Number.Abs(Duration.Days(_ - CurrentDate))}
    ),
    SortedByDiff = List.Sort(
        DateWithDiff,
        (a, b) => if a{1} < b{1} then -1 else if a{1} > b{1} then 1 else 0
    ),
    MakeSenseDate = if List.Count(SortedByDiff) > 0 then SortedByDiff{0}{0} else null,

    // 7. เลือกผลลัพธ์สุดท้ายตาม localeValue หรือ fallback เป็น MakeSenseDate
    localeValue = if locale = null then "auto" else Text.Lower(locale),
    Result =
        if localeValue = "th" then TryTH
        else if localeValue = "uk" then TryUK
        else if localeValue = "us" then TryUS
        else MakeSenseDate
in
    Result
