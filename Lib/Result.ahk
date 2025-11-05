#Requires AutoHotkey v2.0

class Result
{
    __New(success, value)
    {
        this.Success := success
        this.Value := value
    }

    static Success(value)
    {
        return Result(true, value)
    }

    static Error()
    {
        return Result(false, {})
    }
}
