local M = {}

function M.create(ctx)
	local mp = ctx.mp

	local detected_backend = nil
	local app_running_cache_value = nil
	local app_running_cache_time = nil
	local app_running_check_inflight = false
	local app_running_waiters = {}
	local APP_RUNNING_CACHE_TTL_SECONDS = 2

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

	local function now_seconds()
		if type(mp.get_time) == "function" then
			local value = tonumber(mp.get_time())
			if value then
				return value
			end
		end
		return os.time()
	end

	local function process_list_has_subminer(raw_process_list)
		if type(raw_process_list) ~= "string" then
			return false
		end
		local process_list = raw_process_list:lower()
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

	local function process_scan_command()
		if is_windows() then
			return { "tasklist", "/FO", "CSV", "/NH" }
		end
		return { "ps", "-A", "-o", "args=" }
	end

	local function is_subminer_process_running()
		local result = mp.command_native({
			name = "subprocess",
			args = process_scan_command(),
			playback_only = false,
			capture_stdout = true,
			capture_stderr = false,
		})
		if not result or result.status ~= 0 then
			return false
		end
		return process_list_has_subminer(result.stdout)
	end

	local function flush_app_running_waiters(value)
		local waiters = app_running_waiters
		app_running_waiters = {}
		for _, waiter in ipairs(waiters) do
			waiter(value)
		end
	end

	local function is_subminer_app_running_async(callback, opts)
		opts = opts or {}
		local force_refresh = opts.force_refresh == true
		local now = now_seconds()
		if not force_refresh and app_running_cache_value ~= nil and app_running_cache_time ~= nil then
			if (now - app_running_cache_time) <= APP_RUNNING_CACHE_TTL_SECONDS then
				callback(app_running_cache_value)
				return
			end
		end

		app_running_waiters[#app_running_waiters + 1] = callback
		if app_running_check_inflight then
			return
		end
		app_running_check_inflight = true

		mp.command_native_async({
			name = "subprocess",
			args = process_scan_command(),
			playback_only = false,
			capture_stdout = true,
			capture_stderr = false,
		}, function(success, result)
			app_running_check_inflight = false
			local running = false
			if success and result and result.status == 0 then
				running = process_list_has_subminer(result.stdout)
			end
			app_running_cache_value = running
			app_running_cache_time = now_seconds()
			flush_app_running_waiters(running)
		end)
	end

	local function is_subminer_app_running()
		local running = is_subminer_process_running()
		app_running_cache_value = running
		app_running_cache_time = now_seconds()
		return running
	end

	local function set_subminer_app_running_cache(running)
		app_running_cache_value = running == true
		app_running_cache_time = now_seconds()
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
		is_subminer_app_running_async = is_subminer_app_running_async,
		set_subminer_app_running_cache = set_subminer_app_running_cache,
		detect_backend = detect_backend,
	}
end

return M
