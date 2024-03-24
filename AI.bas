Attribute VB_Name = "AI"
' Excel LLM add-in by Abram Jackson
' https://github.com/abrakjamson/ExcelLLM
' MIT license

' Add this function to the workbook start to add usage tooltips
Function GenerateDescriptions()
    Call AddAICorrectDescription
    Call AddAITemplateDescription
    Call AddAIDescription
End Function

Sub AddAICorrectDescription()
    Application.MacroOptions _
        Macro:="AICorrect", _
        Description:="Correct typos and formatting errors in the input data. Value is the cell with data to fix, and category is the expected type of data, such as ""email address"" or ""phone number"""
End Sub

Function AICorrect(value As String, category As String)
Attribute AICorrect.VB_Description = "Correct typos and formatting errors in the input data. Value is the cell with data to fix, and category is the expected type of data, such as ""email address"" or ""phone number"""
    Dim userInputs() As String
    Dim assistantOutputs() As String
    
    ' Add first shot example
    ReDim userInputs(0)
    userInputs(0) = "Data: michael_nguyen6@@hoymail com\nType: email address\n"
    ReDim assistantOutputs(0)
    assistantOutputs(0) = "michael_nguyen6@hotmail.com"
    
    ' Add second shot example
    ReDim Preserve userInputs(1)
    userInputs(1) = "Data: ((218-) 983--0013\nType: phone number\n"
    ReDim Preserve assistantOutputs(1)
    assistantOutputs(1) = "(218) 983-0013"
   
    Dim data As String
    data = "Data: " & value & "\n" & "Type: " & category & "\n"
    
    AICorrect = AIInternalCall(data, 20, "Correct the data to the type.", userInputs(), assistantOutputs())
    
End Function

Sub AddAITemplateDescription()
    Application.MacroOptions _
        Macro:="AITemplate", _
        Description:="Calls the language model with multiple placeholders provided. For example, ""Convert %1 to type %2"", A1, ""French""."
End Sub

Function AITemplate(template As String, ParamArray args() As Variant)
Attribute AITemplate.VB_Description = "Calls the language model with multiple placeholders provided. For example, ""Convert %1 to type %2"", A1, ""French""."
    ' Replace placeholders in the template with values from the args array
    Dim i As Integer
    Dim prompt As String
    prompt = template
    For i = LBound(args) To UBound(args)
        prompt = Replace(prompt, "%" & i + 1, args(i))
    Next i
    
    Dim userInputs() As String
    Dim assistantOutputs() As String
    
    ' Add first shot example
    ReDim userInputs(0)
    userInputs(0) = "Alphabatize these words: Giraffe, Monkey, Aardvark"
    ReDim assistantOutputs(0)
    assistantOutputs(0) = "Aardvark, Giraffe, Monkey"
    
    ' Add second shot example
    ReDim Preserve userInputs(1)
    userInputs(1) = "Give me the keywords of 'I like bananas and pineapples.'"
    ReDim Preserve assistantOutputs(1)
    assistantOutputs(1) = "bananas, pineapples"
    
    ' Add third shot example
    ReDim Preserve userInputs(1)
    userInputs(1) = "What is the the capital and continent of France?"
    ReDim Preserve assistantOutputs(1)
    assistantOutputs(1) = "Paris, Europe"
        
    AITemplate = AIInternalCall(prompt, 150, "Complete the task with as few words as possible. Return only the answer.", userInputs(), assistantOutputs())
End Function

Sub AddAIDescription()
    Application.MacroOptions _
        Macro:="AI", _
        Description:="Call a language model to get a response. The userExamples and assistantExamples are for few-shot prompting.", _
        category:="Custom Functions", _
        ArgumentDescriptions:=Array( _
            "userPrompt: The input to the language model. This can be an entire instruction or used with a system prompt and few-shot examples", _
            "systemPrompt: The optional initial system message to set the behavior of the assistant.", _
            "userInputs: An optional array of the user portion of few-shot examples.", _
            "assistantOutputs: An optional array of the assistant portion of few-shot examples.", _
            "maxTokens: An optional parameter to limit the length of the generated message. If not provided, a default value of 150 is used.")
End Sub

Function AIAdvanced(userPrompt As String, maxTokens As Integer, systemPrompt As String, userExamples As Range, assistantExamples As Range)
    Dim users() As String
    Dim assistants() As String
    
    If Not IsMissing(userExamples) And Not IsMissing(assistantExamples) Then
        Dim i As Integer
        ReDim users(userExamples.Rows.Count - 1)
        ReDim assistants(assistantExamples.Rows.Count - 1)
        For i = 0 To userExamples.Rows.Count - 1
 '           ReDim users(i)
            users(i) = userExamples.Cells(i + 1, 1)
 '           ReDim assistants(i)
            assistants(i) = assistantExamples(i + 1, 1)
        Next i
    End If
    
    AIAdvanced = AIInternalCall(userPrompt, maxTokens, systemPrompt, users, assistants)
End Function

Function AIInternalCall(userPrompt As String, maxTokens As Integer, systemPrompt As String, userExamples() As String, assistantExamples() As String)
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' If you are using a local model host, the API key can be anything
    ' If you want to point this at OpenAI, you'll need a real API key
    Const apiKey As String = "sk-000000"
    
    ' The URL for the OpenAI Completions API
    Const apiUrl As String = "http://localhost:1234/v1/chat/completions"
        
    ' Setting the Max Tokens prevents problems if the LLM gets into a loop
    If IsNull(maxTokens) Then
        maxTokens = 150
    End If
    
    ' Prepare the JSON payload
    Dim jsonBody As String
    jsonBody = "{""model"":""text-davinci-003"",""max_tokens"":" & maxTokens & ",""messages"":["
    
    If Not IsNull(systemPrompt) Then
        jsonBody = jsonBody & "{""role"": ""system"", ""content"": """ & systemPrompt & """},"
    End If
    
    ' Add the few shot examples, if any
    If Not IsMissing(userExamples) Then
        Dim i As Integer
        For i = LBound(userExamples) To UBound(userExamples)
            jsonBody = jsonBody & "{""role"": ""user"", ""content"": """ & userExamples(i) & """},"
            jsonBody = jsonBody & "{""role"": ""assistant"", ""content"": """ & assistantExamples(i) & """},"
        Next i
    End If
    
    ' Add the actual prompt
    jsonBody = jsonBody & "{""role"": ""user"", ""content"": """ & userPrompt & """}"
    
    jsonBody = jsonBody & "]}"
    
    
    ' Open the HTTP request
    xmlHttp.Open "POST", apiUrl, False
    xmlHttp.setRequestHeader "Authorization", "Bearer " & apiKey
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    AIChatCompletions = jsonBody
    ' Send the request with the JSON payload
    xmlHttp.send jsonBody

    ' Debug JSON request
    AIInternalCall = jsonBody

    ' Check the response
    If xmlHttp.Status = 200 Then
        ' Parse the JSON response
        Dim Parsed As Object
        Dim vJSON
        Dim sState As String
        JSON.Parse xmlHttp.responseText, vJSON, sState

        ' Extract the response text from the first (only) choice
        Dim Output
        Output = vJSON("choices")(0)("message")("content")
        AIInternalCall = Output

    Else
        AIInternalCall = "Error " & xmlHttp.Status & ": " & xmlHttp.statusText
    End If
End Function
