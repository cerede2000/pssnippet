function SendMessage([IntPtr] $hWnd, [Int32] $message, [Int32] $wParam, [Int32] $lParam) {
  $parameterTypes = [IntPtr], [Int32], [Int32], [Int32]
  $parameters = $hWnd, $message, $wParam, $lParam

  Invoke-Win32 "user32.dll" ([Int32]) "SendMessage" $parameterTypes $parameters
}

function GetConsoleWindow() {
  Invoke-Win32 "kernel32" ([IntPtr]) "GetConsoleWindow"
}

function ShowWindow([IntPtr] $hWnd, [Int32] $nCmdShow) {
  $parameterTypes = [IntPtr], [Int32]
  $parameters = $hWnd, $nCmdShow

  Invoke-Win32 "user32" ([IntPtr]) "ShowWindow" $parameterTypes $parameters
}

function SetConsoleIcon([IntPtr] $IconHandle) {
  $parameterTypes = [IntPtr]
  $parameters = $IconHandle

  Invoke-Win32 "kernel32.dll" ([bool]) "SetConsoleIcon" $parameterTypes $parameters
}
