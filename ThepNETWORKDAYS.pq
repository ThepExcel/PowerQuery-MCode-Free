// ThepNETWORKDAYS: นับวันทำการระหว่าง startDate กับ endDate
// Parameters:
//   startDate as date                   – วันที่เริ่ม
//   endDate as date                     – วันที่สิ้นสุด
//   optional weekendPattern as text     – สตริง 7 ตัว 0/1 แทนวันหยุดสุดสัปดาห์ (Mon→Sun) เช่น "0000011"
//                                         ถ้าไม่ใส่จะใช้ค่า default = "0000011" (Sat,Sun)
//   optional holidays as nullable list  – (ไม่บังคับ) list ของวันที่หยุดพัก (date values)
// Returns:
//   จำนวนวันทำการ (number)
(startDate as date, endDate as date, optional weekendPattern as text, optional holidays as nullable list) as number =>
let
    // สลับถ้าวันสิ้นสุดน้อยกว่าวันเริ่ม
    RealStart      = if endDate < startDate then endDate else startDate,
    RealEnd        = if endDate < startDate then startDate else endDate,

    // Default weekendPattern เป็น Sat,Sun
    weekend        = if weekendPattern <> null and Text.Length(weekendPattern)=7 then weekendPattern else "0000011",

    // สร้าง list ของวันที่จาก RealStart ถึง RealEnd
    AllDates       = List.Dates(RealStart, Duration.Days(RealEnd - RealStart) + 1, #duration(1,0,0,0)),

    // กรองเอาเฉพาะวันซึ่งในสตริง weekend คือ “0” (ไม่ใช่ weekend)
    WorkdayFilter  = List.Select(
                        AllDates,
                        each
                            let
                                w = Date.DayOfWeek(_, Day.Monday),                // 0=Mon … 6=Sun
                                c = Text.Middle(weekend, w, 1)                    // เลือก char ตรงตำแหน่ง
                            in
                                c = "0"
                    ),

    // ตัดวันหยุดราชการหรือวันหยุดพิเศษถ้ามี
    Cleaned        = if holidays <> null then List.RemoveItems(WorkdayFilter, holidays) else WorkdayFilter,

    // นับจำนวนวันที่เหลือ
    Result         = List.Count(Cleaned)
in
    Result
