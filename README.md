# ThepExcel Custom M Functions (Free)

Custom M Functions สำหรับ Power Query ใน Excel และ Power BI โดย [ThepExcel.com](https://www.thepexcel.com/thepexcel-mfx/)

## วิธีใช้งาน

1. เปิด Power Query Editor ใน Excel หรือ Power BI
2. สร้าง Blank Query ใหม่
3. เปิด Advanced Editor แล้ววางโค้ดจากไฟล์ `.pq` ที่ต้องการ
4. ตั้งชื่อ Query ตามชื่อฟังก์ชัน (เช่น `ThepTrim`)
5. เรียกใช้ฟังก์ชันจาก Query อื่นได้เลย

---

## รายการฟังก์ชัน

### Text Processing (จัดการข้อความ)

#### ThepTrim

ตัดช่องว่าง (หรืออักขระที่กำหนด) ที่ซ้ำกันออก เหลือแค่ตัวเดียว ทำงานได้ดีกว่า `Text.Trim` ตรงที่จัดการช่องว่างตรงกลางข้อความด้วย

```
ThepTrim(OriginalText as text, optional TrimChar as text) as text
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `TrimChar` | text (optional) | อักขระที่ต้องการตัด (default: `" "` ช่องว่าง) |

**ตัวอย่าง:**
- `ThepTrim("Hello    World")` → `"Hello World"`
- `ThepTrim("A--B---C", "-")` → `"A-B-C"`

---

#### ThepExtractNumber

ดึงเฉพาะตัวเลขออกจากข้อความ สามารถเลือกเก็บจุดทศนิยมและช่องว่างได้

```
ThepExtractNumber(OriginalText as text, optional Keepdot as logical, optional Keepspace as logical) as text
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `Keepdot` | logical (optional) | เก็บจุดทศนิยมไว้ด้วยหรือไม่ (default: `true`) |
| `Keepspace` | logical (optional) | เก็บช่องว่างไว้ด้วยหรือไม่ (default: `true`) |

**ตัวอย่าง:**
- `ThepExtractNumber("ราคา 1,234.50 บาท")` → `"1234.50"`
- `ThepExtractNumber("โทร 081-234-5678", false, false)` → `"0812345678"`

---

#### ThepThaiNumbertoArabic

แปลงตัวเลขไทย (๐-๙) เป็นตัวเลขอารบิก (0-9)

```
ThepThaiNumbertoArabic(OriginalText as text) as text
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความที่มีตัวเลขไทย |

**ตัวอย่าง:**
- `ThepThaiNumbertoArabic("แมว ๓๘ ตัว ราคา ๕๔๓๖ บาท")` → `"แมว 38 ตัว ราคา 5436 บาท"`

---

### Regex (Regular Expression)

ฟังก์ชัน Regex ทั้งหมดใช้ JavaScript ผ่าน `Web.Page()` เนื่องจาก Power Query ไม่มี Regex ในตัว

#### ThepRegExExtract

ดึงข้อความที่ตรงกับ Regex pattern (ผลลัพธ์ตัวแรกที่เจอ)

```
ThepRegExExtract(OriginalText as text, RegExPattern as text, optional RegExMode as text) as any
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `RegExPattern` | text | Regex pattern |
| `RegExMode` | text (optional) | โหมด เช่น `"g"`, `"i"`, `"gi"` (default: single match, case sensitive) |

**ตัวอย่าง:**
- `ThepRegExExtract("Hello 123 World 456", "\d+")` → `"123"`

---

#### ThepRegExMatchCount

นับจำนวนข้อความที่ match กับ Regex pattern

```
ThepRegExMatchCount(OriginalText as text, RegExPattern as text, optional RegExMode as text) as number
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `RegExPattern` | text | Regex pattern |
| `RegExMode` | text (optional) | โหมดเพิ่มเติม เช่น `"i"` (default: case sensitive, นับทั้งหมดเสมอ) |

**ตัวอย่าง:**
- `ThepRegExMatchCount("abc 123 def 456 ghi 789", "\d+")` → `3`

---

#### ThepRegExReplace

แทนที่ข้อความด้วย Regex pattern

```
ThepRegExReplace(OriginalText as text, RegExPattern as text, NewText as text, optional RegExMode as text) as any
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `RegExPattern` | text | Regex pattern ที่ต้องการหา |
| `NewText` | text | ข้อความที่ใช้แทนที่ |
| `RegExMode` | text (optional) | โหมด เช่น `"g"` แทนที่ทั้งหมด, `"gi"` แทนที่ทั้งหมดไม่สน case |

**ตัวอย่าง:**
- `ThepRegExReplace("Hello 123", "\d+", "XXX")` → `"Hello XXX"`

---

### Date (จัดการวันที่)

#### ThepDatefromText

แปลงข้อความวันที่ (ตัวเลขล้วน) เป็น date โดยระบุ format

```
ThepDatefromText(DateText as text, DateFormat as text, optional offsetYear as number) as date
```

| Parameter | Type | Description |
|---|---|---|
| `DateText` | text | ข้อความวันที่ (เช่น `"20240115"`, `"15/01/2024"`) |
| `DateFormat` | text | รูปแบบ: `"yyyymmdd"`, `"ddmmyyyy"`, หรือ `"mmddyyyy"` |
| `offsetYear` | number (optional) | ค่าชดเชยปี เช่น `-543` สำหรับแปลง พ.ศ. เป็น ค.ศ. (default: `0`) |

**ตัวอย่าง:**
- `ThepDatefromText("20240115", "yyyymmdd")` → `#date(2024, 1, 15)`
- `ThepDatefromText("15012567", "ddmmyyyy", -543)` → `#date(2024, 1, 15)`

---

#### ThepGenDateTableFromDate

สร้างตาราง Date Table จากวันที่เริ่มต้นถึงวันที่สิ้นสุด

```
ThepGenDateTableFromDate(StartDate as date, EndDate as date) as table
```

| Parameter | Type | Description |
|---|---|---|
| `StartDate` | date | วันที่เริ่มต้น |
| `EndDate` | date | วันที่สิ้นสุด |

**ผลลัพธ์:** ตารางที่มีคอลัมน์ `Date` แต่ละแถวเป็นวันที่ต่อเนื่องกัน

---

#### ThepGenDateTableFromText

สร้างตาราง Date Table จากข้อความวันที่ format `yyyymmdd`

```
ThepGenDateTableFromText(StartDateText as text, EndDateText as text) as table
```

| Parameter | Type | Description |
|---|---|---|
| `StartDateText` | text | วันที่เริ่มต้น format `yyyymmdd` (เช่น `"20240101"`) |
| `EndDateText` | text | วันที่สิ้นสุด format `yyyymmdd` (เช่น `"20241231"`) |

**ผลลัพธ์:** ตารางที่มีคอลัมน์ `Date` แต่ละแถวเป็นวันที่ต่อเนื่องกัน

---

#### ThepNETWORKDAYS

นับจำนวนวันทำการ (เหมือน NETWORKDAYS ใน Excel) รองรับกำหนดวันหยุดสุดสัปดาห์เองและวันหยุดพิเศษ

```
ThepNETWORKDAYS(startDate as date, endDate as date, optional weekendPattern as text, optional holidays as nullable list) as number
```

| Parameter | Type | Description |
|---|---|---|
| `startDate` | date | วันที่เริ่มต้น |
| `endDate` | date | วันที่สิ้นสุด |
| `weekendPattern` | text (optional) | สตริง 7 ตัว (Mon→Sun) `"0"` = วันทำงาน, `"1"` = วันหยุด (default: `"0000011"` คือ Sat-Sun) |
| `holidays` | list (optional) | list ของวันหยุดพิเศษ (date values) |

**ตัวอย่าง:**
- `ThepNETWORKDAYS(#date(2024,1,1), #date(2024,1,31))` → นับวันจันทร์-ศุกร์
- `ThepNETWORKDAYS(#date(2024,1,1), #date(2024,1,31), "1000001")` → หยุดวันจันทร์และอาทิตย์

---

### Table Operations (จัดการตาราง)

#### ThepGetColumnName

ดึงชื่อคอลัมน์จากลำดับที่ (เริ่มจาก 1)

```
ThepGetColumnName(TableName as table, ColNumber as number) as text
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตารางต้นฉบับ |
| `ColNumber` | number | ลำดับคอลัมน์ (เริ่มจาก 1) |

**ตัวอย่าง:**
- `ThepGetColumnName(MyTable, 3)` → ชื่อคอลัมน์ที่ 3

---

#### ThepGetMultipleListItem

ดึงหลาย item จาก list ด้วย index (รองรับ index ติดลบนับจากท้าย)

```
ThepGetMultipleListItem(OriginalList as list, PosIndex as list) as list
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalList` | list | list ต้นฉบับ |
| `PosIndex` | list | list ของ index ที่ต้องการ (เลขติดลบ = นับจากท้าย เช่น `-1` = ตัวสุดท้าย) |

**ตัวอย่าง:**
- `ThepGetMultipleListItem({"A","B","C","D","E"}, {0, 2, -1})` → `{"A", "C", "E"}`

---

#### ThepRenameColumn

เปลี่ยนชื่อคอลัมน์โดยใช้ลำดับที่ (ไม่ต้องรู้ชื่อเดิม) รองรับเปลี่ยนทีละหลายคอลัมน์

```
ThepRenameColumn(TableName as table, ColNumber as any, NewName as any) as table
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตารางต้นฉบับ |
| `ColNumber` | number หรือ list of number | ลำดับคอลัมน์ที่ต้องการเปลี่ยนชื่อ (เริ่มจาก 1, ติดลบ = นับจากท้าย) |
| `NewName` | text หรือ list of text | ชื่อใหม่ |

**ตัวอย่าง:**
- `ThepRenameColumn(MyTable, 1, "ID")` → เปลี่ยนชื่อคอลัมน์แรกเป็น "ID"
- `ThepRenameColumn(MyTable, {1, 3}, {"ID", "Name"})` → เปลี่ยนชื่อคอลัมน์ที่ 1 และ 3

---

#### ThepReplaceAllError

แทนที่ค่า Error ทุกตัวในทุกคอลัมน์ของตารางด้วยค่าที่กำหนด

```
ThepReplaceAllError(TableName as table, optional ReplaceWith as text) as table
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตารางต้นฉบับ |
| `ReplaceWith` | text (optional) | ค่าที่ใช้แทน Error (default: `null`) |

**ตัวอย่าง:**
- `ThepReplaceAllError(MyTable)` → แทนที่ Error ทั้งหมดด้วย null
- `ThepReplaceAllError(MyTable, "N/A")` → แทนที่ Error ทั้งหมดด้วย "N/A"

---

#### ThepOneHot

ทำ One-Hot Encoding แปลงค่าในคอลัมน์ categorical เป็นคอลัมน์ binary (0/1) แยกตามค่า

```
ThepOneHot(TableName as table, TargetColumnName as text) as table
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตารางต้นฉบับ |
| `TargetColumnName` | text | ชื่อคอลัมน์ที่ต้องการทำ One-Hot |

**ตัวอย่าง:** ถ้าคอลัมน์ "Color" มีค่า "Red", "Blue", "Green" จะได้คอลัมน์ใหม่ "Red", "Blue", "Green" ที่มีค่า 1 หรือ 0

---

### API Integration

#### ThepOpenAI

เรียกใช้ OpenAI Chat Completion API จาก Power Query

```
ThepOpenAI(apiKey as text, userPrompt as text, optional systemPrompt as nullable text, optional model as nullable text) as text
```

| Parameter | Type | Description |
|---|---|---|
| `apiKey` | text | OpenAI API Key |
| `userPrompt` | text | ข้อความที่ต้องการส่งถาม |
| `systemPrompt` | text (optional) | System prompt (default: `"you are helpful assistance"`) |
| `model` | text (optional) | ชื่อโมเดล (default: `"gpt-4o-mini"`) |

**ตัวอย่าง:**
- `ThepOpenAI("sk-xxx", "สรุปข้อความนี้ให้หน่อย: ...")` → คำตอบจาก GPT

---

## More Details

https://www.thepexcel.com/thepexcel-mfx/
