; ====================================================================================
; === AQUA BRIEFING - OUTLOOK SYNC AGENT
; ===
; === Purpose: 
; === 1. Fetches Unread Emails (Last 24h) using your CUSTOM FORMAT (HTML Ready).
; === 2. Fetches Upcoming Meetings (Today-Friday) using your CUSTOM FORMAT (HTML Ready).
; === 3. Pushes data to a SEPARATE Firebase document (User_outlook) to avoid overwriting tasks.
; ====================================================================================

#SingleInstance, Force
#Warn
SendMode Input
SetWorkingDir, %A_ScriptDir%

; --- CONFIGURATION ---
global g_ReportID
global g_Prompt
global g_RunAtLogin

; --- Define the config file path ---
configFile := "G:\dailybriefingconfig.txt"
configSection := "Settings"

; --- Read values from INI ---
IniRead, g_ReportID, %configFile%, %configSection%, ReportID, default
IniRead, g_Prompt, %configFile%, %configSection%, Prompt, Yes
IniRead, g_RunAtLogin, %configFile%, %configSection%, RunAtLogin, Yes

; --- Validate ID ---
if (g_ReportID = "default" || g_ReportID = "")
{
    MsgBox, 0x10, Config Error, Please open G:\dailybriefingconfig.txt and set ReportID to your user ID (e.g. jdoe).
    ExitApp
}

; ====================================================================================
; --- MAIN EXECUTION ---
; ====================================================================================

Main()
ExitApp

Main() {
    ; 0. Manage Startup Shortcut based on config
    StartUp() 

    ; 1. Get Outlook Data using your custom formatting functions
    unreadEmailsObj  := GetUnreadEmails()
    todayMeetingsObj := GetTodayMeetings()

    ; 2. Push to Firebase (Safe Mode)
    PushOutlookData(unreadEmailsObj, todayMeetingsObj)
}



; ====================================================================================
; --- 1. Get Outlook Emails (HTML Formatted) ---
; ====================================================================================

GetUnreadEmails() {
    emails := ""
    try {
        ol := GetOutlook()
        if !IsObject(ol)
            return {content: ol, count: 0} 
        ns := ol.GetNamespace("MAPI")
        inbox := ns.GetDefaultFolder(6)
        yesterday := A_Now
        EnvAdd, yesterday, -1, Days
        FormatTime, filterTime, %yesterday%, MM/dd/yyyy hh:mm tt
        filter := "[Unread] = true AND [ReceivedTime] >= '" . filterTime . "'"
        restricted := inbox.Items.Restrict(filter)
        restricted.Sort("[ReceivedTime]", True)
        if (restricted.Count = 0)
            return {content: "No unread emails in the last 24 hours.", count: 0}
        
        emails .= "&nbsp;&nbsp;&nbsp;&nbsp;<small style='color:#6c7475; font-size:0.9em;'>Unread emails (last 24h): Sent date\time, sender, subject, and email preview...</small><br><br>"
        counter := 1 
        
        for item in restricted {
            ; --- Clean the item body to plain text ---
            CleanBody := item.Body
            CleanSubject := item.Subject

            ; Merge real paragraph boundaries
            CleanBody := RegExReplace(CleanBody, "(?i)</p>\s*<p[^>]*>", " ")
            CleanBody := RegExReplace(CleanBody, "(?i)&lt;/p&gt;\s*&lt;p[^&]*&gt;", " ")

            ; Remove all real HTML tags
            CleanBody := RegExReplace(CleanBody, "<[^>]*>", "")
            CleanBody := RegExReplace(CleanBody, "(?i)&lt;/?p[^&]*&gt;", "")
            CleanSubject := RegExReplace(CleanSubject, "<[^>]*>", "")

            ; Normalize whitespace
            CleanBody := StrReplace(StrReplace(CleanBody, "`r", " "), "`n", " ")
            CleanBody := RegExReplace(CleanBody, "\s+", " ")
            CleanBody := Trim(CleanBody)
            CleanSubject := StrReplace(StrReplace(CleanSubject, "`r", " "), "`n", " ")
            CleanSubject := RegExReplace(CleanSubject, "\s+", " ")
            CleanSubject := Trim(CleanSubject)

            ; Truncate
            CleanBody := (StrLen(CleanBody) > 75) ? SubStr(CleanBody, 1, 75) . "..." : CleanBody
            CleanSubject := (StrLen(CleanSubject) > 50) ? SubStr(CleanSubject, 1, 50) . "..." : CleanSubject
            
            ; Format Sent Time (using simple property access if possible, otherwise could apply Ole helper here too)
            try {
                sentTime := item.SentOn
            } catch {
                sentTime := "Unknown"
            }

            emails .= "" . counter . "<b>" . ".</b> [" . sentTime . "]: " . item.SenderName . ": " . CleanSubject . "<br>"
            emails .= "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<small style='color:#666; font-size:0.8em;'><i>EMAIL PREVIEW:</i>&nbsp;" . CleanBody . "</small><br><br>"
            
            counter++ 
        }
    } catch e {
        return {content: "Error: Failed to get Outlook emails. " . e.message, count: 0}
    }
    
    return {content: emails, count: counter - 1}
}

; ====================================================================================
; --- 2. Date Helper Functions (YOUR EXACT LOGIC) ---
; ====================================================================================

OleDateToYmdHMS(oleDate) {
    ; This logic forces the COM Object into a Double calculation to avoid 1899 errors
    dateVal := ComObjValue(oleDate) + 0.0
    VarSetCapacity(st, 16, 0)
    ok := DllCall("OleAut32\VariantTimeToSystemTime", "double", dateVal, "ptr", &st, "int")
    if (!ok)
        return ""
    y := NumGet(st, 0, "UShort"), m := NumGet(st, 2, "UShort"), d := NumGet(st, 6, "UShort")
    hh := NumGet(st, 8, "UShort"), mi := NumGet(st, 10, "UShort"), ss := NumGet(st, 12, "UShort")
    return Format("{:04}{:02}{:02}{:02}{:02}{:02}", y,m,d,hh,mi,ss)
}

FormatOleDate(oleDate, ByRef outDate, ByRef outTime) {
    ts := OleDateToYmdHMS(oleDate)
    if (ts = "") { 
        outDate := "", outTime := ""
        return false
    }
    ; Format: MM/dd/yyyy (Date) and h:mm tt (Time)
    FormatTime, outDate, %ts%, MM/dd/yyyy
    FormatTime, outTime, %ts%, h:mm tt
    return true
}

; ====================================================================================
; --- 3. Get Outlook Meetings (Using Your Date Logic) ---
; ====================================================================================

GetTodayMeetings() {
    olFolderCalendar := 9
    Today_WDay := A_WDay
    if (Today_WDay = 7)
        return {content: "Today is Saturday - no meetings remaining this week.", count: 0}
    DaysToAdd := 6 - Today_WDay 
    FormatTime, StartTime, %A_Now%, MMMM d, yyyy 00:00
    EndTime := A_Now
    EndTime += %DaysToAdd%, Days
    FormatTime, EndTime, %EndTime%, MMMM d, yyyy 23:59
    Filter := "[Start] >= '" . StartTime . "' AND [Start] <= '" . EndTime . "'" 

    try {
        olApp := ComObjActive("Outlook.Application")
    } catch {
        try olApp := ComObjCreate("Outlook.Application")
        catch e
            return {content: "Error: Could not start or connect to Outlook.", count: 0} 
    }

    try {
        ns := olApp.GetNamespace("MAPI")
        cal := ns.GetDefaultFolder(olFolderCalendar)
        items := cal.Items
        items.IncludeRecurrences := 1
        items.Sort("[Start]")
        filtered := items.Restrict(Filter)
        if (filtered.Count = 0)
            return {content: "No meetings found from today through Friday.", count: 0}
        
        out := "<b>Your Meetings (Today - Friday):</b><br>===============================<br><br>"
        counter := 1 

        for item in filtered {
            try {
                 if !FormatOleDate(item.Start, meetDate, startTime)
                    continue
                FormatOleDate(item.End, dummyDate, endTime)
            
 
                subject  := item.Subject ? item.Subject : "(No Subject)"
                location := item.Location ? item.Location : "(No Location)"
                meetDate := item.Start ? item.Start : "(No Date)"
                endDate  := item.End ? item.End : "(No End Date)"
                body := item.Body
                if (body != "") {
                    body := RegExReplace(body, "\r|\n", " ")
                    body := SubStr(body, 1, 200)
         
                    if (StrLen(item.Body) > 200)
                        body .= "..."
                } else
                    body := "(No Notes)"

                ; --- MODIFIED to include counter ---
                out .= "" . counter . "<b>" . ". Date/Time:</b>&nbsp;" . meetDate . " - " . endTime . "<br>"
                out .= "<small style='color:#666; font-size:0.8em;'>&nbsp;&nbsp;&nbsp;&nbsp;<b>Subject  :</b>&nbsp;" . subject . "</small><br>"
                out .= "<small style='color:#666; font-size:0.8em;'>&nbsp;&nbsp;&nbsp;&nbsp;<i>Location :</i>&nbsp;" . location . "</small><br>"
                out .= "<small style='color:#666; font-size:0.8em;'>&nbsp;&nbsp;&nbsp;&nbsp;<i>Notes    :</i>&nbsp;" . body . "</small><br>"
                out .= "------------------------------------<br>"
                
                counter++ ; <-- ADDED

            } catch err {
                continue
            }
        }
        
        return {content: out, count: counter - 1}
        
    } catch e {
        return {content: "Error: Unable to fetch Outlook meetings. " . e.Message, count: 0}
    }
}

; ====================================================================================
; --- 4. STARTUP SHORTCUT MANAGEMENT ---
; ====================================================================================

StartUp() {
    global g_RunAtLogin
    
    ; Define paths (used for both creation and deletion)
    StartupFolder := A_StartMenu . "\Programs\Startup"
    TargetExe := "G:\AquaDailyBriefing.exe"
    ShortcutPath := StartupFolder . "\AquaDailyBriefing.lnk"
    
    ; Normalize the config value for reliable comparison
    CleanRunAtLogin := RegExReplace(g_RunAtLogin, "[\s""]", "")
    
    if (CleanRunAtLogin = "Yes") 
    {
        ; --- Action: Create Shortcut (if needed) ---
        ; Check if the shortcut already exists
        if (!FileExist(ShortcutPath))
        {
            ; Create the shortcut
            FileCreateShortcut, %TargetExe%, %ShortcutPath%
            
            ; Error check is optional here, as before
            ; if (ErrorLevel != 0) { ... }
        }
    } 
    else if (CleanRunAtLogin = "No") 
    {
        ; --- Action: Delete Shortcut (if it exists) ---
        ; Check if the shortcut exists
        if (FileExist(ShortcutPath))
        {
            ; Delete the shortcut file
            FileDelete, %ShortcutPath%
            
            ; Optional: Error check
            if (ErrorLevel != 0)
            {
                ; MsgBox, 0x10, Shortcut Error, Failed to delete the startup shortcut: %ShortcutPath%
            }
        }
    }
}

; ====================================================================================
; --- UTILITIES ---
; ====================================================================================

GetOutlook() {
    try {
        return ComObjActive("Outlook.Application")
    } catch {
        try {
            return ComObjCreate("Outlook.Application")
        } catch {
            return "Outlook not running"
        }
    }
}

PushOutlookData(emailsObj, meetingsObj) {
    global g_ReportID, g_Prompt
    
    
 functionURL := "https://us-central1-dailybriefing-fe7df.cloudfunctions.net/updateBriefing"

    ; --- SAFE MODE: Append _outlook to ID ---
    safeUploadID := EscapeJSON(g_ReportID . "_outlook")
    
    safeEmails    := EscapeJSON(emailsObj.content)
    safeMeetings  := EscapeJSON(meetingsObj.content)
    
    FormatTime, dateStr, %A_Now%, dddd, MMMM d, yyyy
    
    json_payload := "{"
    ; Corrected line continuation for safer compilation
    json_payload .= """reportId"": """    . safeUploadID . """, "
    json_payload .= """dateString"": """  . dateStr      . """, "
    
    json_payload .= """meetings"": """    . safeMeetings . """, "
    json_payload .= """meetings_count"": " . meetingsObj.count . ", "
    
    json_payload .= """emails"": """       . safeEmails   . """, "
    json_payload .= """emails_count"": "   . emailsObj.count . ", "
    
    json_payload .= """tasks"": """", "
    json_payload .= """projects"": """", "
    json_payload .= """activeTasks"": """", "
    json_payload .= """tasks_count"": 0, "
    json_payload .= """projects_count"": 0, "
    json_payload .= """activeTasks_count"": 0"
    
    json_payload .= "}"

    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    whr.Open("POST", functionURL, false)
    whr.SetRequestHeader("Content-Type", "application/json")
    
    try {
        whr.Send(json_payload)
    
     CleanPrompt := RegExReplace(g_Prompt, "[\s""]", "")

        if (whr.Status == 200) {
            if (CleanPrompt = "Yes") {
                
                ; --- DETERMINE GREETING BASED ON TIME OF DAY ---
                currentHour := A_Hour
                if (currentHour >= 5 && currentHour < 12)
                    greeting := "Good Morning!"
                else if (currentHour >= 12 && currentHour < 17)
                    greeting := "Good Afternoon!"
                else
                    greeting := "Good Evening!"
                
                ; --- MESSAGE COMPOSITION ---
                
                emailWord := (emailsObj.count = 1) ? "email" : "emails"
                meetingWord := (meetingsObj.count = 1) ? "meeting" : "meetings"
                
                meetingsSummary := ""
                if (meetingsObj.count > 0)
                    meetingsSummary := "You have (" . meetingsObj.count . ") scheduled " . meetingWord . " through Friday."
                else
                    meetingsSummary := "No meetings scheduled through Friday. Enjoy the open calendar!"

                emailsSummary := ""
                if (emailsObj.count > 0)
                    emailsSummary := "You have (" . emailsObj.count . ") unread " . emailWord . " to catch up on (last 24h)."
                else
                    emailsSummary := "Your inbox is clear! (No unread emails in the last 24h)."
                
               ; Final Message
                
                msg := greeting . "`n`n"
                msg .= "- " . meetingsSummary . "`n"
                msg .= "- " . emailsSummary . "`n`n"
                msg .= "View your full Aqua Daily Briefing Report?"
                
                MsgBox, 68, Aqua Daily Briefing, % msg
                IfMsgBox, Yes
                    ; Corrected line: Ensure the Run command expression is on a single line
                    Run, % "https://dailyaquabriefing.github.io/?daily=" . g_ReportID
            }
        } else {
            MsgBox, 0x10, Sync Error, % "Upload failed.`nStatus: " . whr.Status . "`nResponse: " . whr.ResponseText
        }
    } catch e {
        MsgBox, 0x10, Error, % "Upload failed: " . e.Message
    }
}

EscapeJSON(str) {
  ; This must be a standalone function
    str := StrReplace(str, "\", "\\")
    str := StrReplace(str, """", "\""")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`n", "\n")
    return str

}