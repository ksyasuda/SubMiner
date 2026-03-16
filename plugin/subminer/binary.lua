local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local utils = ctx.utils
	local opts = ctx.opts
	local state = ctx.state
	local environment = ctx.environment
	local subminer_log = ctx.log.subminer_log

	local function normalize_binary_path_candidate(candidate)
		if type(candidate) ~= "string" then
			return nil
		end
		local trimmed = candidate:match("^%s*(.-)%s*$") or ""
		if trimmed == "" then
			return nil
		end
		if #trimmed >= 2 then
			local first = trimmed:sub(1, 1)
			local last = trimmed:sub(-1)
			if (first == '"' and last == '"') or (first == "'" and last == "'") then
				trimmed = trimmed:sub(2, -2)
			end
		end
		return trimmed ~= "" and trimmed or nil
	end

	local function binary_candidates_from_app_path(app_path)
		if environment.is_windows() then
			return {
				utils.join_path(app_path, "SubMiner.exe"),
				utils.join_path(app_path, "subminer.exe"),
			}
		end

		return {
			utils.join_path(app_path, "Contents", "MacOS", "SubMiner"),
			utils.join_path(app_path, "Contents", "MacOS", "subminer"),
		}
	end

	local function file_exists(path)
		local info = utils.file_info(path)
		if not info then
			return false
		end
		if info.is_dir ~= nil then
			return not info.is_dir
		end
		return true
	end

	local function directory_exists(path)
		local info = utils.file_info(path)
		return info ~= nil and info.is_dir == true
	end

	local function resolve_binary_candidate(candidate)
		local normalized = normalize_binary_path_candidate(candidate)
		if not normalized then
			return nil
		end

		if file_exists(normalized) then
			return normalized
		end

		if environment.is_windows() then
			if not normalized:lower():match("%.exe$") then
				local with_exe = normalized .. ".exe"
				if file_exists(with_exe) then
					return with_exe
				end
			end

			if directory_exists(normalized) then
				for _, path in ipairs(binary_candidates_from_app_path(normalized)) do
					if file_exists(path) then
						return path
					end
				end
			end

			return nil
		end

		if not normalized:lower():find("%.app") then
			return nil
		end

		local app_root = normalized
		if not app_root:lower():match("%.app$") then
			app_root = normalized:match("(.+%.app)")
		end
		if not app_root then
			return nil
		end

		for _, path in ipairs(binary_candidates_from_app_path(app_root)) do
			if file_exists(path) then
				return path
			end
		end

		return nil
	end

	local function find_binary_override()
		for _, env_name in ipairs({ "SUBMINER_APPIMAGE_PATH", "SUBMINER_BINARY_PATH" }) do
			local path = resolve_binary_candidate(os.getenv(env_name))
			if path and path ~= "" then
				return path
			end
		end

		return nil
	end

	local function add_search_path(search_paths, candidate)
		if type(candidate) == "string" and candidate ~= "" then
			search_paths[#search_paths + 1] = candidate
		end
	end

	local function trim_subprocess_stdout(value)
		if type(value) ~= "string" then
			return nil
		end
		local trimmed = value:match("^%s*(.-)%s*$") or ""
		if trimmed == "" then
			return nil
		end
		return trimmed
	end

	local function find_windows_binary_via_system_lookup()
		if not environment.is_windows() then
			return nil
		end
		if not mp or type(mp.command_native) ~= "function" then
			return nil
		end

		local script = [=[
function Emit-FirstExistingPath {
  param([string[]]$Candidates)

  foreach ($candidate in $Candidates) {
    if ([string]::IsNullOrWhiteSpace($candidate)) {
      continue
    }
    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
      Write-Output $candidate
      exit 0
    }
  }
}

$runningProcess = Get-CimInstance Win32_Process |
  Where-Object { $_.Name -ieq 'SubMiner.exe' -or $_.Name -ieq 'subminer.exe' } |
  Select-Object -First 1 -Property ExecutablePath, CommandLine
if ($null -ne $runningProcess) {
  Emit-FirstExistingPath @($runningProcess.ExecutablePath)
}

$localAppData = [Environment]::GetFolderPath('LocalApplicationData')
$programFiles = [Environment]::GetFolderPath('ProgramFiles')
$programFilesX86 = ${env:ProgramFiles(x86)}

Emit-FirstExistingPath @(
  $(if (-not [string]::IsNullOrWhiteSpace($localAppData)) { Join-Path $localAppData 'Programs\SubMiner\SubMiner.exe' } else { $null }),
  $(if (-not [string]::IsNullOrWhiteSpace($programFiles)) { Join-Path $programFiles 'SubMiner\SubMiner.exe' } else { $null }),
  $(if (-not [string]::IsNullOrWhiteSpace($programFilesX86)) { Join-Path $programFilesX86 'SubMiner\SubMiner.exe' } else { $null }),
  'C:\SubMiner\SubMiner.exe'
)

foreach ($registryPath in @(
  'HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\SubMiner.exe',
  'HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\SubMiner.exe',
  'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\SubMiner.exe'
)) {
  try {
    $appPath = (Get-ItemProperty -Path $registryPath -ErrorAction Stop).'(default)'
    Emit-FirstExistingPath @($appPath)
  } catch {
  }
}

try {
  $commandPath = Get-Command SubMiner.exe -ErrorAction Stop | Select-Object -First 1 -ExpandProperty Source
  Emit-FirstExistingPath @($commandPath)
} catch {
}
]=]

		local result = mp.command_native({
			name = "subprocess",
			args = {
				"powershell.exe",
				"-NoProfile",
				"-ExecutionPolicy",
				"Bypass",
				"-Command",
				script,
			},
			playback_only = false,
			capture_stdout = true,
			capture_stderr = false,
		})
		if not result or result.status ~= 0 then
			return nil
		end

		local candidate = trim_subprocess_stdout(result.stdout)
		if not candidate then
			return nil
		end

		return resolve_binary_candidate(candidate)
	end

	local function find_binary()
		local override = find_binary_override()
		if override then
			return override
		end

		local configured = resolve_binary_candidate(opts.binary_path)
		if configured then
			return configured
		end

		local system_lookup_binary = find_windows_binary_via_system_lookup()
		if system_lookup_binary then
			subminer_log("info", "binary", "Found Windows binary via system lookup at: " .. system_lookup_binary)
			return system_lookup_binary
		end

		local home = os.getenv("HOME") or os.getenv("USERPROFILE") or ""
		local app_data = os.getenv("APPDATA") or ""
		local app_data_local = app_data ~= "" and app_data:gsub("[/\\][Rr][Oo][Aa][Mm][Ii][Nn][Gg]$", "\\Local") or ""
		local local_app_data = os.getenv("LOCALAPPDATA") or utils.join_path(home, "AppData", "Local")
		local program_files = os.getenv("ProgramFiles") or "C:\\Program Files"
		local program_files_x86 = os.getenv("ProgramFiles(x86)") or "C:\\Program Files (x86)"
		local search_paths = {}

		if environment.is_windows() then
			add_search_path(search_paths, utils.join_path(app_data_local, "Programs", "SubMiner", "SubMiner.exe"))
			add_search_path(search_paths, utils.join_path(local_app_data, "Programs", "SubMiner", "SubMiner.exe"))
			add_search_path(search_paths, utils.join_path(program_files, "SubMiner", "SubMiner.exe"))
			add_search_path(search_paths, utils.join_path(program_files_x86, "SubMiner", "SubMiner.exe"))
			add_search_path(search_paths, "C:\\SubMiner\\SubMiner.exe")
		else
			add_search_path(search_paths, "/Applications/SubMiner.app/Contents/MacOS/SubMiner")
			add_search_path(search_paths, utils.join_path(home, "Applications", "SubMiner.app", "Contents", "MacOS", "SubMiner"))
			add_search_path(search_paths, utils.join_path(home, ".local", "bin", "SubMiner.AppImage"))
			add_search_path(search_paths, "/opt/SubMiner/SubMiner.AppImage")
			add_search_path(search_paths, "/usr/local/bin/SubMiner")
			add_search_path(search_paths, "/usr/local/bin/subminer")
			add_search_path(search_paths, "/usr/bin/SubMiner")
			add_search_path(search_paths, "/usr/bin/subminer")
		end

		for _, path in ipairs(search_paths) do
			if file_exists(path) then
				subminer_log("info", "binary", "Found binary at: " .. path)
				return path
			end
		end

		return nil
	end

	local function ensure_binary_available()
		if state.binary_available and state.binary_path and file_exists(state.binary_path) then
			return true
		end

		local discovered = find_binary()
		if discovered then
			state.binary_path = discovered
			state.binary_available = true
			return true
		end

		state.binary_path = nil
		state.binary_available = false
		return false
	end

	return {
		normalize_binary_path_candidate = normalize_binary_path_candidate,
		file_exists = file_exists,
		find_binary = find_binary,
		ensure_binary_available = ensure_binary_available,
		is_windows = environment.is_windows,
	}
end

return M
