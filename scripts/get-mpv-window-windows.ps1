param(
  [ValidateSet('geometry')]
  [string]$Mode = 'geometry',
  [string]$SocketPath
)

$ErrorActionPreference = 'Stop'

try {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class SubMinerWindowsHelper {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

  [StructLayout(LayoutKind.Sequential)]
  public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }

  [DllImport("user32.dll")]
  public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern bool IsIconic(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern IntPtr GetForegroundWindow();

  [DllImport("user32.dll", SetLastError = true)]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

  [DllImport("dwmapi.dll")]
  public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);
}
"@

  $DWMWA_EXTENDED_FRAME_BOUNDS = 9

  function Get-WindowBounds {
    param([IntPtr]$hWnd)

    $rect = New-Object SubMinerWindowsHelper+RECT
    $size = [System.Runtime.InteropServices.Marshal]::SizeOf($rect)
    $dwmResult = [SubMinerWindowsHelper]::DwmGetWindowAttribute(
      $hWnd,
      $DWMWA_EXTENDED_FRAME_BOUNDS,
      [ref]$rect,
      $size
    )

    if ($dwmResult -ne 0) {
      if (-not [SubMinerWindowsHelper]::GetWindowRect($hWnd, [ref]$rect)) {
        return $null
      }
    }

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    if ($width -le 0 -or $height -le 0) {
      return $null
    }

    return [PSCustomObject]@{
      X = $rect.Left
      Y = $rect.Top
      Width = $width
      Height = $height
      Area = $width * $height
    }
  }

  $commandLineByPid = @{}
  if (-not [string]::IsNullOrWhiteSpace($SocketPath)) {
    foreach ($process in Get-CimInstance Win32_Process) {
      $commandLineByPid[[uint32]$process.ProcessId] = $process.CommandLine
    }
  }

  $matches = New-Object System.Collections.Generic.List[object]
  $foregroundWindow = [SubMinerWindowsHelper]::GetForegroundWindow()
  $callback = [SubMinerWindowsHelper+EnumWindowsProc]{
    param([IntPtr]$hWnd, [IntPtr]$lParam)

    if (-not [SubMinerWindowsHelper]::IsWindowVisible($hWnd)) {
      return $true
    }

    if ([SubMinerWindowsHelper]::IsIconic($hWnd)) {
      return $true
    }

    [uint32]$windowProcessId = 0
    [void][SubMinerWindowsHelper]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)
    if ($windowProcessId -eq 0) {
      return $true
    }

    try {
      $process = Get-Process -Id $windowProcessId -ErrorAction Stop
    } catch {
      return $true
    }

    if ($process.ProcessName -ine 'mpv') {
      return $true
    }

    if (-not [string]::IsNullOrWhiteSpace($SocketPath)) {
      $commandLine = $commandLineByPid[[uint32]$windowProcessId]
      if ([string]::IsNullOrWhiteSpace($commandLine)) {
        return $true
      }
      if (
        ($commandLine -notlike "*--input-ipc-server=$SocketPath*") -and
        ($commandLine -notlike "*--input-ipc-server $SocketPath*")
      ) {
        return $true
      }
    }

    $bounds = Get-WindowBounds -hWnd $hWnd
    if ($null -eq $bounds) {
      return $true
    }

    $matches.Add([PSCustomObject]@{
      HWnd = $hWnd
      X = $bounds.X
      Y = $bounds.Y
      Width = $bounds.Width
      Height = $bounds.Height
      Area = $bounds.Area
      IsForeground = ($foregroundWindow -ne [IntPtr]::Zero -and $hWnd -eq $foregroundWindow)
    })

    return $true
  }

  [void][SubMinerWindowsHelper]::EnumWindows($callback, [IntPtr]::Zero)

  $focusedMatch = $matches | Where-Object { $_.IsForeground } | Select-Object -First 1
  if ($null -ne $focusedMatch) {
    [Console]::Error.WriteLine('focus=focused')
  } else {
    [Console]::Error.WriteLine('focus=not-focused')
  }

  if ($matches.Count -eq 0) {
    Write-Output 'not-found'
    exit 0
  }

  $bestMatch = if ($null -ne $focusedMatch) {
    $focusedMatch
  } else {
    $matches | Sort-Object -Property Area, Width, Height -Descending | Select-Object -First 1
  }
  Write-Output "$($bestMatch.X),$($bestMatch.Y),$($bestMatch.Width),$($bestMatch.Height)"
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
