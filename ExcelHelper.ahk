#Requires AutoHotkey v2.0
#Include <Sequence>
#Include <StrConcat>
#Include <TryParseInteger>
#Include <Array>
#Include <Result>

id := 0
window := unset
DEBUG := !A_IsCompiled

#HotIf DEBUG
F5::Reload

Render()
{
    global window

    window := Gui("+ToolWindow", "Excel Window")
    window.OnEvent("Escape", (*) => HideWindow())

    publishButton := window.AddButton("Default", "P&ublish (Ctrl + Shift + F12)")
    publishButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => Publish()]))
    
    refreshButton := window.AddButton("", "&Refresh")
    refreshButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => Refresh()]))

    getWorkItemsButton := window.AddButton("", "&Get Work Items (Ctrl + F12)")
    getWorkItemsButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => GetWorkItems()]))

    addTreeLevelButton := window.AddButton("", "Add &Tree Level")
    addTreeLevelButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => AddTreeLevel()]))

    fixRowsButton := window.AddButton("", "&Fix Rows (F12)")
    fixRowsButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => FixRows()]))

    addLinkButton := window.AddButton("", "Add Re&lated")
    addLinkButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => AddRelated()]))

    manageLinksButton := window.AddButton("", "&Manage Links")
    manageLinksButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => ManageLinks()]))

    addPredecessorButton := window.AddButton("", "Add &Predecessor")
    addPredecessorButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => AddPredecessor()]))

    addSuccessorButton := window.AddButton("", "Add &Successor")
    addSuccessorButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => AddSuccessor()]))

    openInWebButton := window.AddButton("", "&Open in Web (Alt-Gr + F12)")
    openInWebButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => OpenInWeb()]))

    reparentButton := window.AddButton("", "R&eparent")
    reparentButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => ReparentWorkItem()]))

    window.Show()
}

HideWindow()
{
    global window

    window.Destroy()
    window := unset
}

#HotIf WinActive("Create Work Item.xlsx ahk_exe EXCEL.EXE ahk_class XLMAIN")
F12::
{
    global id := WinGetID("A")

    global selectParentMode := false

    Render()
}

; Ctrl + Shift + F12
^>+F12::
{
    Sleep 100

    global id := WinGetID("A")

    Publish()
}

; Ctrl + F12
^F12::
{
    Sleep 500
    
    global id := WinGetID("A")

    GetWorkItems()
}

; Alt-Gr + F12
<^>!F12::
{
    Sleep 100

    global id := WinGetID("A")

    OpenInWeb()
}
#HotIf

#HotIf IsSet(window)
F12::
{
    global id

    HideWindow()

    WinActivate(id)
    FixRows()
}
#HotIf

; Use window ID instead
#HotIf WinActive("Links and Attachments")
F12::
{
    global id, selectParentMode

    if (selectParentMode == false)
    {
        return
    }
    else
    {
        selectParentMode := false
        Send "^+{Tab}" ; Navigate to 
        Send "{Left 2}"
        Send "{Space}"
        Send "{Tab}"

        SuggestFillWorkItemIDs()
    }
}
#HotIf

GetSuggestedWorkItemIDs()
{
    Sleep 500
    clipboardCopy := A_Clipboard
    result := TryParseIDs(clipboardCopy)

    ; If the value in the clipboard looks like a work item ID, ask whether to use it
    if (result.Success)
    {
        formattedValue := StrReplace(result.Value, ";", "`n - ")

        ; The clipboard contains multiple IDs
        if (InStr(result.Value, ";") > 0)
        {
            formattedValue := "`n - " . formattedValue
        }

        useValueFromClipboard := MsgBox("Use value from clipboard: " formattedValue, "Confirm?", "YesNo") == "Yes"

        if (useValueFromClipboard)
        {
            return result
        }
    }

    return false
}

ApplySuggestedWorkItemIDs(suggestedIds)
{
    ; ASSUMPTION: The work item IDs are in the clipboard
    ; ASSUMPTION: The input field is focused

    if (suggestedIds && suggestedIds.Success)
    {
        Send StrReplace(suggestedIds.Value, ";", ",")
        Send "{Enter}"
        return true
    }

    return false
}

SuggestFillWorkItemIDs()
{
    suggestedIds := GetSuggestedWorkItemIDs()
    return ApplySuggestedWorkItemIDs(suggestedIds)
}

ManageLinks()
{
    global id

    WinActivate(id)
    OpenLinksAndAttachments()
    Send "{Tab 3}"
}

global selectParentMode := false

ReparentWorkItem()
{
    global id, selectParentMode

    WinActivate(id)
    linkToWindow := OpenLinkToDialog()
    WinActivate(linkToWindow)

    Send "p"

    ;Sleep 500

    linkTypeHwnd := ControlGetFocus("ahk_id " . linkToWindow)
    selectedText := ControlGetText(linkTypeHwnd)

    if (selectedText != "Parent")
    {
        Send "{Esc}"
        Sleep 50
        Send "{Tab 3}"
        MsgBox("The work item already has a parent.`n`nPlease select the parent in the list and press F12 to reparent.", "Select existing parent", 48)
        selectParentMode := true
        return
    }
    else
    {
        Send "{Tab}"
        SuggestFillWorkItemIDs()
    }
}

Publish()
{
    global id

    WinActivate(id)
    Send "{Alt down}"
    Send "{Alt up}"
    Sleep 50
    Send "y2p"
}

Refresh()
{
    global id

    WinActivate(id)
    Send "{Alt down}"
    Send "{Alt up}"
    Sleep 50
    Send "y2r"
}

AddTreeLevel()
{
    global id

    WinActivate(id)
    Send "{Alt down}"
    Send "{Alt up}"
    Sleep 50
    Send "y2a"

    addTreeLevelWindow := WinWaitActive("Convert to Tree List")
    WinActivate(addTreeLevelWindow)
    Send "{Enter}"
}

GetWorkItems()
{
    global id

    WinActivate(id)
    Send "{Alt down}"
    Send "{Alt up}"
    Sleep 50
    Send "y2g"

    getWorkItemsWindow := WinWaitActive("Get Work Items")
    WinActivate(getWorkItemsWindow)

    ; Make sure the "IDs" radio button is selected
    Send "{Alt down}ii{Alt up}"
    Send "{Tab}"

    suggestedWorkItemIDs := GetSuggestedWorkItemIDs()
    if (suggestedWorkItemIDs && suggestedWorkItemIDs.Success)
    {
        Send suggestedWorkItemIDs.Value
        Send "{Enter}"
        Send "!s"
        Send "{Space}"
        Send "{Enter}"
    }
}

TryParseIDs(clipboardCopy)
{
    delimiters := Array(";", "`n")
    trimChars := " `r`n"

    if (clipboardCopy = "")
    {
        return Result.Error()
    }

    trimmed := Trim(clipboardCopy, trimChars)

    if (InStr(trimmed, ";") == 0 && InStr(trimmed, "`r`n") == 0)
    {
        ; If the clipboard does not contain a semicolon or newline, it is not a valid ID list
        if (TryParseInteger(trimmed, &out))
        {
            ; If the clipboard contains a single integer, it is a valid ID
            return Result.Success(trimmed)
        }
    }

    /*
     * At this point, we assume the clipboard contains multiple IDs
     */

    lines := []

    for _, delimiter in delimiters
    {
        ; Split the clipboard content by the current delimiter
        lines := StrSplit(clipboardCopy, delimiter, " `r`n")

        ; If we have lines, concatenate them to the main lines array
        if (lines.Length > 1)
        {
            break
        }

        lines := [] ; Reset lines if no valid split was found
    }

    ; If we still have no lines, return false
    if (lines.Length == 0)
    {
        return Result.Error()
    }

    lines := ArrayExcept(lines, (value) => value == "")
    
    if (ArrayAll(lines, (value) => TryParseInteger(value, &out)))
    {
        ; If all values are integers, return true
        concatString := StrConcat(lines, ";")

        return Result.Success(concatString)
    }

    return Result.Error()
}

FixRows()
{
    global id

    if (id != 0)
    {
        WinActivate(id)
    }

    Send "^a" ; Select all cells
    Send "!hw" ; Toggle wrap text
    Send "{Up}{Down}" ; Remove selection
}

OpenLinksAndAttachments()
{
    Send "{Alt down}"
    Send "{Alt up}"
    Sleep 50
    Send "y2l"

    return WinWaitActive("Links and Attachments") 
}

OpenLinkToDialog()
{
    linksWindow := OpenLinksAndAttachments()
    WinActivate(linksWindow)
    Send "!l"

    addLinkWindowHwnd := WinWaitActive("Add Link to")
    WinActivate(addLinkWindowHwnd)
    Send "{Home}"

    return addLinkWindowHwnd
}

AddRelated()
{
    global id

    WinActivate(id)
    linkToWindow := OpenLinkToDialog()
    WinActivate(linkToWindow)

    Send "{r 3}"
    Send "{Tab}"

    filled := SuggestFillWorkItemIDs()

    if (filled)
    {
        Sleep 50
        Send "!p"
        Send "{Esc}"
    }
}

AddPredecessor()
{
    global id

    WinActivate(id)
    linkToWindow := OpenLinkToDialog()
    WinActivate(linkToWindow)

    Send "{p 2}"
    Send "{Tab}"

    filled := SuggestFillWorkItemIDs()

    if (filled)
    {
        Sleep 50
        Send "!p"
        Send "{Esc}"
    }
}

AddSuccessor()
{
    global id

    WinActivate(id)
    linkToWindow := OpenLinkToDialog()
    WinActivate(linkToWindow)

    Send "{s 3}"
    Send "{Tab}"

    Sleep 150
    filled := SuggestFillWorkItemIDs()

    if (filled)
    {
        Sleep 50
        Send "!p"
        Send "{Esc}"
    }
}

OpenInWeb()
{
    global id

    WinActivate(id)
    Send "{Alt down}"
    Send "{Alt up}"
    Sleep 50
    Send "y2"
    Sleep 50
    Send "w"
}