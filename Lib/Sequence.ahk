#Requires AutoHotkey v2.0

/**
 * Executes a sequence of functions in order.
 * @param {Array} fns - An array of functions to execute.
 */
Sequence(fns)
{
    for _, fn in fns
    {
        fn()
    }
}