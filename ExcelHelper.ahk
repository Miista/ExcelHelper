#Requires AutoHotkey v2.0
#Include <Sequence>
#Include <StrConcat>
#Include <TryParseInteger>
#Include <Array>
#Include <Result>

id := 0
window := unset
DEBUG := !A_IsCompiled
selectParentWindowHwnd := unset

; GLOBALS
global WindowTitles := {
    CreateWorkItem: "Create Work Item.xlsx ahk_exe EXCEL.EXE ahk_class XLMAIN",
    LinksAndAttachmentsPrefix: "Links and Attachments",
    AddLinkPrefix: "Add Link to",
    GetWorkItems: "Get Work Items",
    ConvertToTreeList: "Convert to Tree List",
    EditLink: "Edit Link",
}

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

#HotIf WinActive(WindowTitles.CreateWorkItem)
F12::
{
    global id := WinGetID("A")
    global selectParentWindowHwnd := unset

    Render()
}

; Ctrl + Shift + F12
^>+F12::
{
    global id := WinGetID("A")

    Publish()
}

; Ctrl + F12
^F12::
{
    global id := WinGetID("A")

    GetWorkItems()
}

; Alt-Gr + F12
<^>!F12::
{
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

; This is only active when reparenting a work item
#HotIf WinActive(WindowTitles.LinksAndAttachmentsPrefix) == selectParentWindowHwnd
F12::
{
    global id

    Send "^+{Tab}" ; Navigate to 
    Send "{Left 2}"
    Send "{Space}"

    WinActivateWait(WindowTitles.EditLink)

    Send "{Tab}"

    filled := SuggestFillWorkItemIDs()

    if (filled)
    {
        Sleep 50
        Send "!p"
        Send "{Esc}"
    }
}
#HotIf

GetSuggestedWorkItemIDs(currentWorkItemId := "")
{
    Sleep 500
    clipboardCopy := A_Clipboard
    result := TryParseIDs(clipboardCopy, currentWorkItemId)

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

        windowClosed := WinWaitClose(WindowTitles.AddLinkPrefix, , 0.5)

        if (windowClosed)
        {
            return true
        }

        ; If the window hasn't closed, an error dialog is most likely displayed on the screen.
        HandleErrors()

        ; Check again if the window is closed
        windowClosed := WinWaitClose(WindowTitles.AddLinkPrefix, , 0.5)

        if (windowClosed)
        {
            return true
        }

        if (!windowClosed)
        {
            MsgBox("Failed to add work item links. Changes will not be published automatically.", "Warning", 48)
            return false
        }
    }

    return false

    HandleErrors()
    {
        if (InStr(GetActiveWindowText(), "TF207015"))
        {
            ; TF207015: The current work item already contains links to the following work items:
            ; Resolution: The duplicate links are removed automatically when pressing Enter
            Send "{Esc}" ; Dismiss error dialog

            WinActivateWait(WindowTitles.AddLinkPrefix)
            Send "{Enter}" ; Try to submit again

            windowClosed := WinWaitClose(WindowTitles.AddLinkPrefix, , 0.5)

            if (!windowClosed)
            {
                newText := GetActiveWindowText()

                if (InStr(newText, "TF207015"))
                {
                    Send "{Esc}" ; Dismiss error dialog

                    MsgBox("It would seem that the work item is already linked to the specified work items. Close this dialog to continue.", "All links exist!", 64)

                    WinWaitActive(WindowTitles.AddLinkPrefix)
                    Send "{Esc}" ; Close the Add Link to window

                    return true
                }

                MsgBox("Failed to add work item links. Changes will not be published automatically.", "Warning", 48)

                return false
            }
        }

        GetActiveWindowText()
        {
            activeWindowHwnd := WinGetID("A")
            windowText := WinGetText(activeWindowHwnd)

            return windowText
        }
    }
}

SuggestFillWorkItemIDs(currentWorkItemId := "")
{
    suggestedIds := GetSuggestedWorkItemIDs(currentWorkItemId)
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
    global id, selectParentWindowHwnd

    WinActivate(id)
    result := OpenLinkToDialog()
    linkToWindow := result.AddLinkWindowHwnd
    WinActivate(linkToWindow)

    Send "p" ; Select "Parent" link type

    linkTypeHwnd := ControlGetFocus(linkToWindow)
    selectedText := ControlGetText(linkTypeHwnd)

    if (selectedText != "Parent")
    {
        ; Not able to select Parent link type, so go back to the Links and Attachments window
        Send "{Esc}"
        WinWaitActive(linkToWindow)

        Send "{Tab 3}" ; Navigate to the link control
        MsgBox("The work item already has a parent.`n`nPlease select the parent in the list and press F12 to reparent.", "Select existing parent", 48)
        selectParentWindowHwnd := result.LinksWindow ; Store the links window handle
        return
    }
    else
    {
        Send "{Tab}"
        filled := SuggestFillWorkItemIDs()

        if (filled)
        {
            Sleep 50
            Send "!p"
            Send "{Esc}"
        }
    }
}

Publish()
{
    global id

    WinActivate(id)
    ExecuteTeamCommand("y2p")
}

Refresh()
{
    global id

    WinActivate(id)
    ExecuteTeamCommand("y2r")
}

AddTreeLevel()
{
    global id

    WinActivate(id)
    ExecuteTeamCommand("y2a")

    WinActivateWait(WindowTitles.ConvertToTreeList)
    Send "{Enter}"
}

GetWorkItems()
{
    global id

    WinActivate(id)
    ExecuteTeamCommand("y2g")
    WinActivateWait(WindowTitles.GetWorkItems)

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

TryParseIDs(clipboardCopy, currentWorkItemId := "")
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

    lines := ArrayExcept(lines, (value) => value == "" or value == currentWorkItemId)
    
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
    ExecuteTeamCommand("y2l")

    return WinActivateWait(WindowTitles.LinksAndAttachmentsPrefix)
}

OpenLinkToDialog()
{
    linksWindow := OpenLinksAndAttachments()
    WinActivate(linksWindow)
    workItemId := TryGetCurrentWorkItemId(linksWindow)
    WinWaitActive(linksWindow)
    Sleep 50
    Send "!l"

    addLinkWindowHwnd := WinActivateWait(WindowTitles.AddLinkPrefix)
    Send "{Home}"

    return {
        AddLinkWindowHwnd: addLinkWindowHwnd,
        LinksWindow: linksWindow,
        WorkItemId: workItemId
    }

    TryGetCurrentWorkItemId(linksWindow)
    {
        windowTitle := WinGetTitle(linksWindow)
        foundMatch := RegExMatch(windowTitle, "i)" . WindowTitles.LinksAndAttachmentsPrefix . " for (?:[A-Za-z]+) (\d+)", &workItemId)

        if (foundMatch)
        {
            ; The first match is the work item ID
            return workItemId.1
        }
        else
        {
            return false
        }
    }
}

AddRelated()
{
    global id

    WinActivate(id)
    result := OpenLinkToDialog()
    linkToWindow := result.AddLinkWindowHwnd
    WinActivate(linkToWindow)

    Send "{r 3}"
    Send "{Tab}"

    filled := SuggestFillWorkItemIDs(result.WorkItemId)

    if (filled)
    {
        Sleep 250
        Send "!p"
        Send "{Esc}"
    }
}

AddPredecessor()
{
    global id

    WinActivate(id)
    result := OpenLinkToDialog()
    linkToWindow := result.AddLinkWindowHwnd
    WinActivate(linkToWindow)

    Send "{p 2}"
    Send "{Tab}"

    filled := SuggestFillWorkItemIDs(result.WorkItemId)

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
    result := OpenLinkToDialog()
    linkToWindow := result.AddLinkWindowHwnd
    WinActivate(linkToWindow)

    Send "{s 3}"
    Send "{Tab}"

    Sleep 150
    filled := SuggestFillWorkItemIDs(result.WorkItemId)

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
    ExecuteTeamCommand("y2w")
}

; HELPER FUNCTIONS
ExecuteTeamCommand(command)
{
    global id

    WinActivate(id)
    Send "{Alt down}"
    Send "{Alt up}"
    Sleep 50
    Send command
}

WinActivateWait(title)
{
    windowHwnd := WinWaitActive(title)
    WinActivate(windowHwnd)

    return windowHwnd
}