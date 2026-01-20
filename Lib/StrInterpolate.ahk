#Requires AutoHotkey v2.0

/**
 * Applies a range of values to a pattern.
 * @remarks The pattern should use {1}, {2}, etc. as placeholders for the values. Note that indexing starts at 1.
 * @param {Array} pattern - The pattern to apply values to.
 * @param {String} values - The values to apply.
 * @returns {String} - The complete string.
 */
StrInterpolate(pattern, values)
{
    result := pattern

    for _, value in values
    {
        result := StrReplace(result, "{" . A_Index . "}", value)
    }

    return result
}