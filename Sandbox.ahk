#Requires AutoHotkey v2.0

F5::Reload

MsgBox("Test", "Warning", 48)

F1::
{
    windowHwnd := WinWaitActive("Add Link to", , 2)
    
    if (windowHwnd == 0)
    {
        return
    }

    windowTitle := WinGetTitle(windowHwnd)
    MsgBox("Window Title: " . windowTitle)
    foundMatch := RegExMatch(windowTitle, "i)Add Link to (?:[A-Za-z]+) (\d+)", &workItemId)

    if (foundMatch)
    {
        return workItemId.1
    }
    else
    {
        return false
    }
}

F2::
{
    activeWindowHwnd := WinGetID("A")
    HWNDs := WinGetControlsHwnd(activeWindowHwnd)

    for id, value IN HWNDs
    {
        windowText := WinGetText(activeWindowHwnd)

        if (InStr(windowText, "TF207015"))
        {
            ; TF207015: The current work item already contains links to the following work items:
            Send "{Enter}"
        }
        else if (InStr(windowText, "TF207008"))
        {
            ; TF207008: Azure DevOps does not support linking a work item to itself.
            return
        }
    }
}

WinActivateWait(title, timeout := 10)
{
    windowHwnd := WinWaitActive(title, timeout)
    WinActivate(windowHwnd)

    return windowHwnd
}