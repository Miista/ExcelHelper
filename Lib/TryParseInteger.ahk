#Requires AutoHotkey v2.0

/**
 * Attempts to parse a value as an integer.
 * @param {any} value - The value to parse.
 * @param {Integer} out - The output variable to store the parsed integer.
 * @returns {Boolean} - Returns true if the value was successfully parsed as an integer,
 */
TryParseInteger(value, &out)
{
    if (value is Integer)
    {
        out := value
        return true
    }
    else if (value is String)
        try
        {
            out := Integer(value)
            return true
        }
        catch TypeError as err
        {
            out := 0
            return false
        }        
    else
        throw ValueError("Value is not an integer or string: " value)
}