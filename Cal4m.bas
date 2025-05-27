Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If



'=====================================================================================
'  CAL4M(prompt)
'     – queries a local Ollama server
'     – sizes the answer to fit the width of the caller cell
'     – caches answers so the same prompt isn’t re-fetched during this Excel session
'
'  Usage in Excel:  =CAL4M(A2)   or   =CAL4M("Tell me a joke about spreadsheets")
'=====================================================================================
Public Function CAL4M(prompt As String, _
                      Optional resultType As String = "string") As Variant

    On Error GoTo ErrorHandler
    
    '----------------------------- 1) quick in-memory cache --------------------------
    Static PromptCache As Object
    If PromptCache Is Nothing Then Set PromptCache = CreateObject("Scripting.Dictionary")
    
    Dim cacheKey As String
    cacheKey = prompt & "¶" & LCase(resultType)       '¶ unlikely to appear in prompt

    If PromptCache.Exists(cacheKey) Then
        CAL4M = PromptCache(cacheKey)
        Exit Function
    End If

    
    
    '----------------------------- 2) estimate visible width -------------------------
    Dim pxWidth As Double
    pxWidth = Application.Caller.Width * 1.3333          'points ? pixels
    
    Const AVG_PX_PER_CHAR As Double = 7.2               'Calibri 11
    Dim charBudget As Long
    charBudget = Int(pxWidth / AVG_PX_PER_CHAR) - 2      '-2 for padding
    If charBudget < 8 Then charBudget = 8
    
    
    Dim maxTokens As Long
    maxTokens = charBudget \ 2                           'more generous than 4-to-1
    If maxTokens < 10 Then maxTokens = 10
        
    'restrict if we have a specific request here
    If resultType = "boolean" Then maxTokens = 4
    If resultType = "number" Then maxTokens = 10
    If resultType = "word" Then maxTokens = 10
    If resultType = "date" Then maxTokens = 10
        
    
    '----------------------------- 3) wait for server up to 30 s ----------------------
    Const GET_URL  As String = "http://localhost:11434/api/tags"
    Const POST_URL As String = "http://localhost:11434/v1/chat/completions"
    Const KEY      As String = """content"":"""
    
    Static serverReady As Boolean
    
    If Not serverReady Then
        
        Dim http As Object, t0 As Single, maxWait As Single
        maxWait = 30: t0 = Timer
        
        Do
            Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            http.Open "GET", GET_URL, False
            http.setRequestHeader "Content-Type", "application/json"
            http.Send
            If http.Status = 200 Then Exit Do       'server
        
            If Timer - t0 > maxWait Then
                CAL4M = "Error: Ollama server not responding."
                Exit Function
            End If
            Sleep 500                               'wait 0.5 s before retry
        Loop
        
        serverReady = True
        
    End If
    
    
    '----------------------------- 4) build JSON payload -----------------------------
    Dim userMsg As String
    userMsg = Replace(prompt, """", "\""")       'escape embedded quotes
        
        
    Dim typeRule As String
    
    Select Case LCase(resultType)
        Case "number", "numeric"
            typeRule = "Return only a valid number (no text, no commas)."
        Case "boolean"
            typeRule = "Return only TRUE or FALSE. Nothing else."
        Case "word"
            typeRule = "Return exactly one word—no spaces, no punctuation."
        Case Else 'string
            typeRule = "Return only " & resultType
    End Select
    
    Dim sysPrompt As String
    
    sysPrompt = "You are an Excel-embedded function. " & _
                typeRule & " Never output back-slashes, pipes, or line-breaks."
        
    
    Dim ctxBudget As Long
    
    'Rough token estimate: chars / 4
    ctxBudget = (Len(sysPrompt) + Len(userMsg)) \ 4 _
                + maxTokens + 16            '16-token slack
    
    'Clamp into sensible bounds
    If ctxBudget < 512 Then ctxBudget = 512
    If ctxBudget > 2048 Then ctxBudget = 2048   'no need for more in this task

    
    '--- build JSON body ---
    Dim body As String
    body = "{" & _
              """model"":""phi3.5:3.8b""," & _
              """max_tokens"":" & CStr(maxTokens) & "," & _
              """temperature"":0.1," & _
              """stream"":false," & _
              """options"":{""num_ctx"":" & ctxBudget & ",""num_batch"":16}," & _
              """messages"":[{" & _
                  """role"":""system""," & _
                  """content"":""" & Replace(sysPrompt, """", "\""") & """" & _
                "},{" & _
                  """role"":""user""," & _
                  """content"":""" & Replace(userMsg, """", "\""") & """" & _
              "}]" & _
           "}"

    '----------------------------- 5) send chat-completion ---------------------------
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", POST_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send body
    
    If http.Status <> 200 Then
        CAL4M = "HTTP " & http.Status & " error"
        Exit Function
    End If
    
    '----------------------------- 6) extract the "content" field --------------------
    Dim resp As String, startPos As Long, endPos As Long, contentStart As Long
    resp = http.responseText
    
    startPos = InStr(resp, KEY)
    If startPos = 0 Then
        CAL4M = "Error: no content field in response."
        Exit Function
    End If
    
    contentStart = startPos + Len(KEY)
    endPos = InStr(contentStart, resp, """")
    If endPos = 0 Then
        CAL4M = "Error: malformed JSON."
        Exit Function
    End If
    
    '----------------------------- 7) final sanitising -------------------------------
    
    Dim answer As String
    answer = Mid$(resp, contentStart, endPos - contentStart)
    
    answer = Replace$(answer, "\n", "|")   '? do this FIRST
    answer = Replace$(answer, "\", "")      'drop lone back-slashes
    answer = Replace$(answer, "|", " ")     'pipe ? space
    answer = Trim$(answer)                  'tidy whitespace
    
    Select Case Left$(answer, 1)                'prevent formula injection
        Case "=", "+", "-", "@": answer = "'" & answer
    End Select
    
    Select Case LCase(resultType)
    
        Case "number", "numeric"
            If IsNumeric(answer) Then
                CAL4M = CDbl(answer)
            Else
                CAL4M = CVErr(xlErrNA)         'signals bad numeric
            End If
            
        Case "boolean"
            Select Case LCase(answer)
                Case "true", "yes", "1": CAL4M = True
                Case "false", "no", "0": CAL4M = False
                Case Else: CAL4M = CVErr(xlErrValue)
            End Select
            
        Case "word"
            answer = Split(answer, " ")(0)        'take first token
            answer = Application.WorksheetFunction.Proper(answer)
            CAL4M = answer
    
        Case Else
            CAL4M = answer              'helper strips formula chars
    
    End Select
    

     '----------------------------- 8) cache & return ---------------------------------
    On Error Resume Next
    If Not PromptCache.Exists(cacheKey) Then
        PromptCache.Add cacheKey, CAL4M
    End If
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    CAL4M = "Error: " & Err.Description & " (Line " & Erl & ")"
    
End Function


