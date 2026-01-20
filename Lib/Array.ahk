#Requires AutoHotkey v2.0

/**
 * Checks if all elements in an array satisfy a given condition.
 * @param {Array} array - The array to check.
 * @param {Function} fn - The function to apply to each element.
 * @returns {Boolean} - Returns true if all elements satisfy the condition, false otherwise.
 */
ArrayAll(array, fn)
{
    for _, value in array
    {
        if (!fn(value))
        {
            return false
        }
    }

    return true
}

/**
 * Creates a new array excluding elements that satisfy a given condition.
 * @param {Array} array - The original array.
 * @param {Function} fn - The function to determine which elements to exclude.
 * @returns {Array} - The new array with excluded elements.
 */
ArrayExcept(array, fn)
{
    newArray := []

    for _, value in array
    {
        if (fn(value))
        {
            continue
        }

        newArray.Push(value)
    }

    return newArray
}

/**
 * Filters an array to contain only elements that satisfy a given condition.
 * @param {Array} array - The original array.
 * @param {Function} fn - The function to determine which elements to keep.
 * @returns {Array} - The new array with included elements.
 */
ArrayFilter(array, fn)
{
    newArray := []

    for _, value in array
    {
        if (fn(value))
        {
            newArray.Push(value)
        }
    }

    return newArray
}

/**
 * Maps the given function over all elements in an array.
 * @param {Array} array - The original array.
 * @param {Function} fn - The mapping function.
 * @returns {Array} - The new array with mapped elements.
 */
ArrayMap(array, fn)
{
    newArray := []

    for _, value in array
    {
        newValue := fn(value)
        newArray.Push(newValue)
    }

    return newArray
}