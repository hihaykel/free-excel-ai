Function FREE_AI(user_question As String) As String
    Dim http As Object
    Dim url As String
    Dim payload As String
    Dim response As String
    Dim system_prompt As String
    Dim clean_question As String
    Dim start_pos As Long, end_pos As Long
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://text.pollinations.ai/openai"
    
    ' 1. Security: Double up or remove quotes to prevent JSON crashes
    clean_question = Replace(user_question, """", " ")
    
    ' Strict instructions for the AI
    system_prompt = "Provide only the raw result or the Excel formula requested. No explanations, no greeting, no markdown, no extra text. Be as concise as possible."
    
    ' JSON payload configuration
    payload = "{""messages"":[" & _
              "{""role"":""system"",""content"":""" & system_prompt & """}," & _
              "{""role"":""user"",""content"":""" & clean_question & """}],""model"":""openai""}"
    
    On Error GoTo ErrorHandler
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send payload
    
    response = http.responseText
    
    ' If the server returned an HTML error page, catch it here
    If InStr(response, "<!DOCTYPE") > 0 Or InStr(response, "<html") > 0 Then
        FREE_AI = "Text Error (Check Quotes)"
        Exit Function
    End If
    
    ' Target and extract ONLY the text inside the "content":"..." block
    start_pos = InStr(response, """content"":""")
    If start_pos > 0 Then
        start_pos = start_pos + 11
        end_pos = InStr(start_pos, response, """")
        If end_pos > 0 Then
            response = Mid(response, start_pos, end_pos - start_pos)
        End If
    End If
    
    ' Clean up any remaining formatting escape characters
    response = Replace(response, "```excel", "")
    response = Replace(response, "```", "")
    response = Replace(response, "\n", " ")
    
    FREE_AI = Trim(response)
    Exit Function

ErrorHandler:
    FREE_AI = "Error"
End Function

