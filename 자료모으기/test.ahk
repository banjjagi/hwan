FileSelectFile, FileName, S16,, Create a new file:
if FileName =
    return
GENERIC_WRITE = 0x40000000  ; Open the file for writing rather than reading.
CREATE_ALWAYS = 2  ; Create new file (overwriting any existing file).
hFile := DllCall("CreateFile", str, FileName, Uint, GENERIC_WRITE, Uint, 0, UInt, 0, UInt, CREATE_ALWAYS, Uint, 0, UInt, 0)
if not hFile
{
    MsgBox Can't open "%FileName%" for writing.
    return
}
TestString = This is a test string.`r`n  ; When writing a file this way, use `r`n rather than `n to start a new line.
DllCall("WriteFile", UInt, hFile, str, TestString, UInt, StrLen(TestString), UIntP, BytesActuallyWritten, UInt, 0)
DllCall("CloseHandle", UInt, hFile)  ; Close the file.