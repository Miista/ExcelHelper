#Requires AutoHotkey v2.0
#Include <Sequence>
#Include <StrConcat>
#Include <StrInterpolate>
#Include <TryParseInteger>
#Include <Array>
#Include <Result>

id := 0
window := unset
subWindow := unset
DEBUG := !A_IsCompiled
selectParentWindowHwnd := unset

lastColumnSpecification := ""

; GLOBALS
global WindowTitles := {
    CreateWorkItem: "Create Work Item.xlsx ahk_exe EXCEL.EXE ahk_class XLMAIN",
    LinksAndAttachmentsPrefix: "Links and Attachments",
    AddLinkPrefix: "Add Link to",
    GetWorkItems: "Get Work Items",
    ConvertToTreeList: "Convert to Tree List",
    EditLink: "Edit Link",
    RulesManager: "Conditional Formatting Rules Manager",
    NewRule: "New Formatting Rule",
    EditRule: "Edit Formatting Rule",
    FormatCells: "Format Cells",
    Colors: "Colors"
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

    getWorkItemsInQueryButton := window.AddButton("", "Get Work Items Pending Estimation (&Q)")
    getWorkItemsInQueryButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => GetWorkItemsPendingEstimation()]))

    addTreeLevelButton := window.AddButton("", "Add &Tree Level")
    addTreeLevelButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => AddTreeLevel()]))

    fixRowsButton := window.AddButton("", "&Fix Rows (F12)")
    fixRowsButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => FixRows()]))

    manageLinksButton := window.AddButton("", "&Manage Links")
    manageLinksButton.OnEvent("Click", (*) => Sequence([() => RenderLinkManager()]))

    openInWebButton := window.AddButton("", "&Open in Web (Alt-Gr + F12)")
    openInWebButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => OpenInWeb()]))

    highlightWorkFieldsButton := window.AddButton("", "&Highlight Work Item Fields (Experimental)")
    highlightWorkFieldsButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => HighlightWorkItemFields()]))

    removeHighlightWorkFieldsButton := window.AddButton("", "Remove Highlight Work Item Fiel&ds (Experimental)")
    removeHighlightWorkFieldsButton.OnEvent("Click", (*) => Sequence([() => HideWindow(), () => RemoveHighlightWorkItemFields()]))

    window.Show()
}

RenderLinkManager()
{
    global window, subWindow

    subWindow := Gui("+ToolWindow", "Manage Links")

    subWindow.OnEvent("Escape", (*) => HideLinkManager())

    reparentButton := subWindow.AddButton("", "R&eparent")
    reparentButton.OnEvent("Click", (*) => Sequence([() => HideLinkManager(true), () => ReparentWorkItem()]))

    addRelatedButton := subWindow.AddButton("", "&Related")
    addRelatedButton.OnEvent("Click", (*) => Sequence([() => HideLinkManager(true), () => AddRelated()]))

    addPredecessorButton := subWindow.AddButton("", "&Predecessor")
    addPredecessorButton.OnEvent("Click", (*) => Sequence([() => HideLinkManager(true), () => AddPredecessor()]))

    addSuccessorButton := subWindow.AddButton("", "&Successor")
    addSuccessorButton.OnEvent("Click", (*) => Sequence([() => HideLinkManager(true), () => AddSuccessor()]))

    manageLinksButton := subWindow.AddButton("", "&Manage Links (F12)")
    manageLinksButton.OnEvent("Click", (*) => Sequence([() => HideLinkManager(true), () => ManageLinks()]))

    if (IsSet(window))
    {
        subWindow.Opt("+Owner" window.Hwnd)
        window.OnEvent("Escape", (*) => {})
        window.Opt("+Disabled")
    }

    subWindow.Show()
}

HideLinkManager(includeParent := false)
{
    global subWindow, window

    if (IsSet(window))
    {
        window.Opt("-Disabled")
    }

    subWindow.Destroy()
    subWindow := unset

    if (includeParent && IsSet(window))
    {
        HideWindow()
    }
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

; Right Win + F12
>#F12::
{
    global id := WinGetID("A")

    RenderLinkManager()
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

#HotIf IsSet(subWindow)
F12::
{
    HideLinkManager(true)
    ManageLinks()
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
#HotIf IsSet(selectParentWindowHwnd) && WinActive(WindowTitles.LinksAndAttachmentsPrefix) == selectParentWindowHwnd
F12::
{
    global id, selectParentWindowHwnd

    Send "^+{Tab}" ; Navigate to 
    Send "{Left 2}"
    Send "{Space}"

    WinActivateWait(WindowTitles.EditLink)

    Send "{Tab}"

    filled := SuggestFillWorkItemIDs_LinkWindow()

    if (filled)
    {
        Sleep 50
        Send "!p"
        Send "{Esc}"
        selectParentWindowHwnd := unset
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

SuggestFillWorkItemIDs_LinkWindow(currentWorkItemId := "", windowHwnd := 0)
{
    suggestedIds := GetSuggestedWorkItemIDs(currentWorkItemId)

    /*if (windowHwnd != 0)
    {
        WinActivate(windowHwnd)
        WinActivateWait(windowHwnd)
    }*/

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
        filled := SuggestFillWorkItemIDs_LinkWindow()

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

GetWorkItemsPendingEstimation()
{
    global id

    WinActivate(id)
    ExecuteTeamCommand("y2g")
    WinActivateWait(WindowTitles.GetWorkItems)

    Send "{Alt down}qq{Alt up}"
    Send "{Tab}"

    ; Open the query dropdown
    Send "{Down}"
    Send "{PgUp}"

    ; Open "Travel Retail"
    Send "Travel"
    Send "{NumpadSub}"
    Send "{NumpadAdd}"

    ; Open "My Queries"
    Send "My"
    Send "{NumpadSub}"
    Send "{NumpadAdd}"

    ; Select "Commerce Ops Pending Estimation"
    Send "Commerce"
    Send "{Enter}"

    Send "!n"
    Send "!s"
    Send "{Enter}"
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

    filled := SuggestFillWorkItemIDs_LinkWindow(result.WorkItemId, linkToWindow)

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

    filled := SuggestFillWorkItemIDs_LinkWindow(result.WorkItemId)

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
    filled := SuggestFillWorkItemIDs_LinkWindow(result.WorkItemId)

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

HighlightWorkItemFields()
{
    global id, lastColumnSpecification

    columns := GetColumns()
    workItemTypeColumn := columns[1]
    remainingWorkColumn := columns[2]
    originalEstimateColumn := columns[3]
    stateColumn := columns[4]

    rulePattern := "=AND(${1}1=`"Task`";${2}1=`"`";${3}1<>`"Closed`"; ${3}1<>`"Removed`")"
    remainingWorkRule := GenerateRule(rulePattern, workItemTypeColumn, remainingWorkColumn, stateColumn)
    originalEstimateRule := GenerateRule(rulePattern, workItemTypeColumn, originalEstimateColumn, stateColumn)

    rangePattern := "=${1}:${1}"
    remainingWorkRange := GenerateRange(rangePattern, remainingWorkColumn)
    originalEstimateRange := GenerateRange(rangePattern, originalEstimateColumn)

    ruleFillColor := "#FF5151"

    WinActivate(id)
    ExecuteTeamCommand("hlr")

    WinActivateWait(WindowTitles.RulesManager)
    AddRule(remainingWorkRule, remainingWorkRange, ruleFillColor)

    WinActivateWait(WindowTitles.RulesManager)
    AddRule(originalEstimateRule, originalEstimateRange, ruleFillColor)

    ; Navigate to "Apply"
    WinActivateWait(WindowTitles.RulesManager)
    Send "!s"
    Send "{Escape}"
    Send "+{Tab}"
    Send "{Space}"
    Send "{Escape}"

    GetColumns()
    {
        ib := InputBox("Specify the columns for Work Item Type, Remaining Work, and Original Estimate, State in that order. Separate with comma.", "Specify columns",, lastColumnSpecification)

        if (ib.Result == "Cancel")
        {
            return
        }

        lastColumnSpecification := ib.Value
        columns := SplitColumns(ib.Value)

        return columns
    }

    GenerateRule(pattern, workItemColumn, targetColumn, stateColumn)
    {
        rule := StrInterpolate(pattern, [workItemColumn, targetColumn, stateColumn])
        return rule
    }

    GenerateRange(pattern, column)
    {
        range := StrInterpolate(pattern, [column])
        return range
    }

    SplitColumns(input)
    {
        columns := []

        if (InStr(input, ",") == 0)
        {
            ; Split by character
            columns := StrSplit(input)
        }
        else
        {
            ; Split by comma
            columns := StrSplit(input, ",")
        }

        if (columns.Length < 4)
        {
            throw Error("Insufficient columns specified. You must specify at least four columns.")
        }
        
        return ArrayMap(columns, (value) => StrUpper(Trim(value)))
    }

    AddRule(rule, range, color)
    {
        WinActivateWait(WindowTitles.RulesManager)
        Send "!n"

        WinActivateWait(WindowTitles.NewRule)
        ; Sending End-Tab positions the cursor at the "Format values where this formula is true" input
        Send "{End}"
        Send "{Tab}"
        SendText rule

        SetColorFormatting(color)

        WinActivateWait(WindowTitles.NewRule)
        Send "!o" ; Apparently, Alt-O focuses the OK button
        Send "{Enter}"

        WinActivateWait(WindowTitles.RulesManager)
        Send "{Tab}"
        SendText range

        /* It's crucial that we move focus away from the range input this.
         * If we don't do this, then--when we return from creating the next rule--
         * the window will freeze and turn white, crashing Excel.
         * That's not ideal. So we kindly tab away from the input.
         */
        ; Tab away from the range input
        Send "{LShift down}"
        Send "{Tab}"
        Send "{LShift up}"

        /* For some reason, the newly added rule will--when applied--use some crazy row number
         * like 154815158. I have no idea why.
         * Anyway, to fix this, we need to edit the rule again and re-enter the correct rule.
         */
        Send "!e"
        WinActivateWait(WindowTitles.EditRule)
        Send "!o"
        SendText rule
        Send "{Enter}"
    }

    SetColorFormatting(color)
    {
        WinActivateWait(WindowTitles.NewRule)

        ; Open Format Cells dialog
        Send "!f"
        WinActivateWait(WindowTitles.FormatCells)

        SelectFillTab()

        ; Select "More Colors..."
        Send "!m"

        WinActivateWait(WindowTitles.Colors)
        Send "^{Tab}"
        Send "!h" ; Move to Hex input
        SendText color
        Send "{Enter}"

        WinActivateWait(WindowTitles.FormatCells)
        Send "{Tab 4}" ; Move to OK button
        Send "{Enter}"

        SelectFillTab()
        {
            /* There are four tabs in the Format Cells dialog.
             * They are:
             *   1. Number
             *   2. Font
             *   3. Border
             *   4. Fill
             * 
             * We need to make sure we are on the "Fill" tab.
             * Unfortunately, there is no direct way to check which tab is active,
             * and even if we are on the Fill tab, hitting Ctrl+Tab will cycle us through the tabs.
             * 
             * Therefore, we need to check which controls are visible.
             * On the Fill tab, there are two "MSO Generic Control Container" controls visible.
             * Thus, we cycle through the tabs until there are two such controls visible.
             * 
             * Fortunately, once we are on the Fill tab, we will still be on the Fill tab
             * when we open the window again.
             */
            controls := WinGetControlsHwnd(WindowTitles.FormatCells)
            controls := ArrayFilter(controls, (value) => ControlGetText(value) == "MSO Generic Control Container")

            ; This check shortcircuits the "expensive" loop. I assume that it's expensive.
            if (IsFillTab())
            {
                return true
            }

            while (!IsFillTab())
            {
                Send "^{Tab}"
                Sleep 50
            }

            IsFillTab()
            {
                vs := []

                for index, value in controls
                {
                    v := ControlGetVisible(value)
                    if (v == true)
                    {
                        vs.Push(v)
                    }
                }

                return vs.Length == 2
            }
        }
    }
}

RemoveHighlightWorkItemFields()
{
    global id

    WinActivate(id)

    ExecuteTeamCommand("hlr")

    WinActivateWait(WindowTitles.RulesManager)

    loop 3
    {
        Send "!d"
    }

    Send "{Enter}"
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