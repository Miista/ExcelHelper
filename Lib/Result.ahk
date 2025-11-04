#Requires AutoHotkey v2.0

Tuple(item1, item2)
{
    return {
        Item1: item1,
        Item2: item2
    }
}

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
