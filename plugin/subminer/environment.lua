local M = {}

function M.create(ctx)
	local mp = ctx.mp

	local detected_backend = nil

	local function is_windows()
		return package.config:sub(1, 1) == "\\"
	end

	local function is_macos()
		local platform = mp.get_property("platform") or ""
		if platform == "macos" or platform == "darwin" then
			return true
		end
		local ostype = os.getenv("OSTYPE") or ""
		return ostype:find("darwin") ~= nil
	end

	local function default_socket_path()
		if is_windows() then
			return "\\\\.\\pipe\\subminer-socket"
		end
		return "/tmp/subminer-socket"
	end

	local function is_linux()
		return not is_windows() and not is_macos()
	end

	local function is_subminer_process_running()
		local command = is_windows() and { "tasklist", "/FO", "CSV", "/NH" } or { "ps", "-A", "-o", "args=" }
		local result = mp.command_native({
			name = "subprocess",
			args = command,
			playback_only = false,
			capture_stdout = true,
			capture_stderr = false,
		})
		if not result or type(result.stdout) ~= "string" or result.status ~= 0 then
			return false
		end

		local process_list = result.stdout:lower()
		for line in process_list:gmatch("[^\n]+") do
			if is_windows() then
				local image = line:match('^"([^"]+)","')
				if not image then
					image = line:match('^"([^"]+)"')
				end
				if not image then
					goto continue
				end
				if image == "subminer" or image == "subminer.exe" or image == "subminer.appimage" or image == "subminer.app" then
					return true
				end
				if image:find("subminer", 1, true) and not image:find(".lua", 1, true) then
					return true
				end
			else
				local argv0 = line:match('^"([^"]+)"') or line:match("^%s*([^%s]+)")
				if not argv0 then
					goto continue
				end
				if argv0:find("subminer.lua", 1, true) or argv0:find("subminer.conf", 1, true) then
					goto continue
				end
				local exe = argv0:match("([^/\\]+)$") or argv0
				if exe == "SubMiner" or exe == "SubMiner.AppImage" or exe == "SubMiner.exe" or exe == "subminer" or exe == "subminer.appimage" or exe == "subminer.exe" then
					return true
				end
				if exe:find("subminer", 1, true) and exe:find("%.lua", 1, true) == nil and exe:find("%.app", 1, true) == nil then
					return true
				end
			end

			::continue::
		end
		return false
	end

	local function is_subminer_app_running()
		return is_subminer_process_running()
	end

	local function detect_backend()
		if detected_backend then
			return detected_backend
		end

		local backend = nil
		local subminer_log = ctx.log and ctx.log.subminer_log or function() end

		if is_macos() then
			backend = "macos"
		elseif is_windows() then
			backend = nil
		elseif os.getenv("HYPRLAND_INSTANCE_SIGNATURE") then
			backend = "hyprland"
		elseif os.getenv("SWAYSOCK") then
			backend = "sway"
		elseif os.getenv("XDG_SESSION_TYPE") == "x11" or os.getenv("DISPLAY") then
			backend = "x11"
		else
			subminer_log("warn", "backend", "Could not detect window manager, falling back to x11")
			backend = "x11"
		end

		detected_backend = backend
		if backend then
			subminer_log("info", "backend", "Detected backend: " .. backend)
		else
			subminer_log("info", "backend", "No backend detected")
		end
		return backend
	end

	return {
		is_windows = is_windows,
		is_macos = is_macos,
		is_linux = is_linux,
		default_socket_path = default_socket_path,
		is_subminer_process_running = is_subminer_process_running,
		is_subminer_app_running = is_subminer_app_running,
		detect_backend = detect_backend,
	}
end

return M
