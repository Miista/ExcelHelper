#Requires AutoHotkey v2.0

/**
 * Concatenates an array of strings with a specified delimiter.
 * @param {Array} strings - The array of strings to concatenate.
 * @param {String} delimiter - The delimiter to use between strings.
 * @returns {String} - The concatenated string.
 */
StrConcat(strings, delimiter)
{
    result := ""

    for _, value in strings
    {
        result .= value . delimiter
    }

    return RTrim(result, delimiter)
}