(apiKey as text, userPrompt as text, optional systemPrompt as nullable text, optional model as nullable text ) as text =>
let
    // กำหนดค่า default ถ้าไม่ได้ส่งมา
    modelName   = if model <> null then model else "gpt-4o-mini",
    systemMsg   = if systemPrompt <> null then systemPrompt else "you are helpful assistance",

    url         = "https://api.openai.com/v1/chat/completions",

    // สร้าง body ตาม spec ของ API
    bodyRecord  = [
        model    = modelName,
        messages = {
            [ role = "system", content = systemMsg ],
            [ role = "user",   content = userPrompt ]
        }
    ],

    // แปลง record เป็น JSON binary (ไม่ต้องใช้ Text.ToBinary)
    jsonBinary  = Json.FromValue(bodyRecord),

    // เรียก API พร้อมส่ง JSON binary ลงใน Content
    responseBin = Web.Contents(
        url,
        [
            Headers = [
                #"Content-Type"  = "application/json",
                #"Authorization" = "Bearer " & apiKey
            ],
            Content = jsonBinary
        ]
    ),

    // แปลงผลลัพธ์ JSON binary เป็น record
    result      = Json.Document(responseBin),
    msgOutput   = result[choices]{0}[message][content]
in
    msgOutput
