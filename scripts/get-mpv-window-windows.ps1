param(
  [ValidateSet('geometry', 'foreground-process', 'bind-overlay', 'lower-overlay', 'set-owner', 'clear-owner')]
  [string]$Mode = 'geometry',
  [string]$SocketPath,
  [string]$OverlayWindowHandle
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
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool SetWindowPos(
    IntPtr hWnd,
    IntPtr hWndInsertAfter,
    int X,
    int Y,
    int cx,
    int cy,
    uint uFlags
  );

  [DllImport("user32.dll", SetLastError = true)]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

  [DllImport("user32.dll", SetLastError = true)]
  public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

  [DllImport("user32.dll", SetLastError = true)]
  public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

  [DllImport("user32.dll", SetLastError = true)]
  public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

  [DllImport("dwmapi.dll")]
  public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);
}
"@

  $DWMWA_EXTENDED_FRAME_BOUNDS = 9
  $SWP_NOSIZE = 0x0001
  $SWP_NOMOVE = 0x0002
  $SWP_NOACTIVATE = 0x0010
  $SWP_NOOWNERZORDER = 0x0200
  $SWP_FLAGS = $SWP_NOSIZE -bor $SWP_NOMOVE -bor $SWP_NOACTIVATE -bor $SWP_NOOWNERZORDER
  $GWL_EXSTYLE = -20
  $WS_EX_TOPMOST = 0x00000008
  $GWLP_HWNDPARENT = -8
  $HWND_TOP = [IntPtr]::Zero
  $HWND_BOTTOM = [IntPtr]::One
  $HWND_TOPMOST = [IntPtr](-1)
  $HWND_NOTOPMOST = [IntPtr](-2)

  if ($Mode -eq 'foreground-process') {
    $foregroundWindow = [SubMinerWindowsHelper]::GetForegroundWindow()
    if ($foregroundWindow -eq [IntPtr]::Zero) {
      Write-Output 'not-found'
      exit 0
    }

    [uint32]$foregroundProcessId = 0
    [void][SubMinerWindowsHelper]::GetWindowThreadProcessId($foregroundWindow, [ref]$foregroundProcessId)
    if ($foregroundProcessId -eq 0) {
      Write-Output 'not-found'
      exit 0
    }

    try {
      $foregroundProcess = Get-Process -Id $foregroundProcessId -ErrorAction Stop
    } catch {
      Write-Output 'not-found'
      exit 0
    }

    Write-Output "process=$($foregroundProcess.ProcessName)"
    exit 0
  }

  if ($Mode -eq 'clear-owner') {
    if ([string]::IsNullOrWhiteSpace($OverlayWindowHandle)) {
      [Console]::Error.WriteLine('overlay-window-handle-required')
      exit 1
    }

    [IntPtr]$overlayWindow = [IntPtr]([int64]$OverlayWindowHandle)
    [void][SubMinerWindowsHelper]::SetWindowLongPtr($overlayWindow, $GWLP_HWNDPARENT, [IntPtr]::Zero)
    Write-Output 'ok'
    exit 0
  }

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

  $mpvMatches = New-Object System.Collections.Generic.List[object]
  $targetWindowState = 'not-found'
  $foregroundWindow = [SubMinerWindowsHelper]::GetForegroundWindow()
  $callback = [SubMinerWindowsHelper+EnumWindowsProc]{
    param([IntPtr]$hWnd, [IntPtr]$lParam)

    if (-not [SubMinerWindowsHelper]::IsWindowVisible($hWnd)) {
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

    if ([SubMinerWindowsHelper]::IsIconic($hWnd)) {
      if (-not [string]::IsNullOrWhiteSpace($SocketPath) -and $targetWindowState -ne 'visible') {
        $targetWindowState = 'minimized'
      }
      return $true
    }

    $bounds = Get-WindowBounds -hWnd $hWnd
    if ($null -eq $bounds) {
      return $true
    }

    if (-not [string]::IsNullOrWhiteSpace($SocketPath)) {
      $targetWindowState = 'visible'
    }

    $mpvMatches.Add([PSCustomObject]@{
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

  if ($Mode -eq 'lower-overlay') {
    if ([string]::IsNullOrWhiteSpace($OverlayWindowHandle)) {
      [Console]::Error.WriteLine('overlay-window-handle-required')
      exit 1
    }

    [IntPtr]$overlayWindow = [IntPtr]([int64]$OverlayWindowHandle)

    [void][SubMinerWindowsHelper]::SetWindowPos(
      $overlayWindow,
      $HWND_NOTOPMOST,
      0,
      0,
      0,
      0,
      $SWP_FLAGS
    )
    [void][SubMinerWindowsHelper]::SetWindowPos(
      $overlayWindow,
      $HWND_BOTTOM,
      0,
      0,
      0,
      0,
      $SWP_FLAGS
    )
    Write-Output 'ok'
    exit 0
  }

  $focusedMatch = $mpvMatches | Where-Object { $_.IsForeground } | Select-Object -First 1
  if ($null -ne $focusedMatch) {
    [Console]::Error.WriteLine('focus=focused')
  } else {
    [Console]::Error.WriteLine('focus=not-focused')
  }
  if (-not [string]::IsNullOrWhiteSpace($SocketPath)) {
    [Console]::Error.WriteLine("state=$targetWindowState")
  }

  if ($mpvMatches.Count -eq 0) {
    Write-Output 'not-found'
    exit 0
  }

  $bestMatch = if ($null -ne $focusedMatch) {
    $focusedMatch
  } else {
    $mpvMatches | Sort-Object -Property Area, Width, Height -Descending | Select-Object -First 1
  }

  if ($Mode -eq 'set-owner') {
    if ([string]::IsNullOrWhiteSpace($OverlayWindowHandle)) {
      [Console]::Error.WriteLine('overlay-window-handle-required')
      exit 1
    }

    [IntPtr]$overlayWindow = [IntPtr]([int64]$OverlayWindowHandle)
    $targetWindow = [IntPtr]$bestMatch.HWnd
    [void][SubMinerWindowsHelper]::SetWindowLongPtr($overlayWindow, $GWLP_HWNDPARENT, $targetWindow)
    Write-Output 'ok'
    exit 0
  }

  if ($Mode -eq 'bind-overlay') {
    if ([string]::IsNullOrWhiteSpace($OverlayWindowHandle)) {
      [Console]::Error.WriteLine('overlay-window-handle-required')
      exit 1
    }

    [IntPtr]$overlayWindow = [IntPtr]([int64]$OverlayWindowHandle)
    $targetWindow = [IntPtr]$bestMatch.HWnd
    $targetWindowExStyle = [SubMinerWindowsHelper]::GetWindowLong($targetWindow, $GWL_EXSTYLE)
    $targetWindowIsTopmost = ($targetWindowExStyle -band $WS_EX_TOPMOST) -ne 0

    $overlayExStyle = [SubMinerWindowsHelper]::GetWindowLong($overlayWindow, $GWL_EXSTYLE)
    $overlayIsTopmost = ($overlayExStyle -band $WS_EX_TOPMOST) -ne 0
    if ($targetWindowIsTopmost -and -not $overlayIsTopmost) {
      [void][SubMinerWindowsHelper]::SetWindowPos(
        $overlayWindow, $HWND_TOPMOST, 0, 0, 0, 0, $SWP_FLAGS
      )
    } elseif (-not $targetWindowIsTopmost -and $overlayIsTopmost) {
      [void][SubMinerWindowsHelper]::SetWindowPos(
        $overlayWindow, $HWND_NOTOPMOST, 0, 0, 0, 0, $SWP_FLAGS
      )
    }

    $GW_HWNDPREV = 3
    $windowAboveMpv = [SubMinerWindowsHelper]::GetWindow($targetWindow, $GW_HWNDPREV)

    if ($windowAboveMpv -ne [IntPtr]::Zero -and $windowAboveMpv -eq $overlayWindow) {
      Write-Output 'ok'
      exit 0
    }

    $insertAfter = $HWND_TOP
    if ($windowAboveMpv -ne [IntPtr]::Zero) {
      $aboveExStyle = [SubMinerWindowsHelper]::GetWindowLong($windowAboveMpv, $GWL_EXSTYLE)
      $aboveIsTopmost = ($aboveExStyle -band $WS_EX_TOPMOST) -ne 0
      if ($aboveIsTopmost -eq $targetWindowIsTopmost) {
        $insertAfter = $windowAboveMpv
      }
    }

    [void][SubMinerWindowsHelper]::SetWindowPos(
      $overlayWindow, $insertAfter, 0, 0, 0, 0, $SWP_FLAGS
    )
    Write-Output 'ok'
    exit 0
  }

  Write-Output "$($bestMatch.X),$($bestMatch.Y),$($bestMatch.Width),$($bestMatch.Height)"
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
