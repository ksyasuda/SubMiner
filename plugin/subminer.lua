local input = require("mp.input")
local mp = require("mp")
local msg = require("mp.msg")
local options = require("mp.options")
local utils = require("mp.utils")

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
	return {
		utils.join_path(app_path, "Contents", "MacOS", "SubMiner"),
		utils.join_path(app_path, "Contents", "MacOS", "subminer"),
	}
end

local opts = {
	binary_path = "",
	socket_path = default_socket_path(),
	texthooker_enabled = true,
	texthooker_port = 5174,
	backend = "auto",
	auto_start = true,
	auto_start_overlay = false, -- legacy alias, maps to auto_start_visible_overlay
	auto_start_visible_overlay = false,
	auto_start_invisible_overlay = "platform-default", -- platform-default | visible | hidden
	osd_messages = true,
	log_level = "info",
}

options.read_options(opts, "subminer")

local state = {
	overlay_running = false,
	texthooker_running = false,
	overlay_process = nil,
	binary_available = false,
	binary_path = nil,
	detected_backend = nil,
	invisible_overlay_visible = false,
}

local LOG_LEVEL_PRIORITY = {
	debug = 10,
	info = 20,
	warn = 30,
	error = 40,
}

local function normalize_log_level(level)
	local normalized = (level or "info"):lower()
	if LOG_LEVEL_PRIORITY[normalized] then
		return normalized
	end
	return "info"
end

local function should_log(level)
	local current = normalize_log_level(opts.log_level)
	local target = normalize_log_level(level)
	return LOG_LEVEL_PRIORITY[target] >= LOG_LEVEL_PRIORITY[current]
end

local function subminer_log(level, scope, message)
	if not should_log(level) then
		return
	end
	local timestamp = os.date("%Y-%m-%d %H:%M:%S")
	local line = string.format("[subminer] - %s - %s - [%s] %s", timestamp, string.upper(level), scope, message)
	if level == "error" then
		msg.error(line)
	elseif level == "warn" then
		msg.warn(line)
	elseif level == "debug" then
		msg.debug(line)
	else
		msg.info(line)
	end
end

local function show_osd(message)
	if opts.osd_messages then
		mp.osd_message("SubMiner: " .. message, 3)
	end
end

local function detect_backend()
	if state.detected_backend then
		return state.detected_backend
	end

	local backend = nil

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

	state.detected_backend = backend
	if backend then
		subminer_log("info", "backend", "Detected backend: " .. backend)
	else
		subminer_log("info", "backend", "No backend detected")
	end
	return backend
end

local function file_exists(path)
	local info = utils.file_info(path)
	if not info then return false end
	if info.is_dir ~= nil then
		return not info.is_dir
	end
	return true
end

local function resolve_binary_candidate(candidate)
	local normalized = normalize_binary_path_candidate(candidate)
	if not normalized then
		return nil
	end

	if file_exists(normalized) then
		return normalized
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
	local candidates = {
		resolve_binary_candidate(os.getenv("SUBMINER_APPIMAGE_PATH")),
		resolve_binary_candidate(os.getenv("SUBMINER_BINARY_PATH")),
	}

	for _, path in ipairs(candidates) do
		if path and path ~= "" then
			return path
		end
	end

	return nil
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

	local search_paths = {
		"/Applications/SubMiner.app/Contents/MacOS/SubMiner",
		utils.join_path(os.getenv("HOME") or "", "Applications/SubMiner.app/Contents/MacOS/SubMiner"),
		"C:\\Program Files\\SubMiner\\SubMiner.exe",
		"C:\\Program Files (x86)\\SubMiner\\SubMiner.exe",
		"C:\\SubMiner\\SubMiner.exe",
		utils.join_path(os.getenv("HOME") or "", ".local/bin/SubMiner.AppImage"),
		"/opt/SubMiner/SubMiner.AppImage",
		"/usr/local/bin/SubMiner",
		"/usr/bin/SubMiner",
	}

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

local function resolve_backend(override_backend)
	local selected = override_backend
	if selected == nil or selected == "" then
		selected = opts.backend
	end
	if selected == "auto" then
		return detect_backend()
	end
	return selected
end

local function build_command_args(action, overrides)
	overrides = overrides or {}
	local args = { state.binary_path }

	table.insert(args, "--" .. action)
	local log_level = normalize_log_level(overrides.log_level or opts.log_level)
	if log_level ~= "info" then
		table.insert(args, "--log-level")
		table.insert(args, log_level)
	end

	local needs_start_context = action == "start"

	if needs_start_context then
		local backend = resolve_backend(overrides.backend)
		if backend and backend ~= "" then
			table.insert(args, "--backend")
			table.insert(args, backend)
		end

		local socket_path = overrides.socket_path or opts.socket_path
		table.insert(args, "--socket")
		table.insert(args, socket_path)
	end

	return args
end

local function run_control_command(action)
	local args = build_command_args(action)
	subminer_log("debug", "process", "Control command: " .. table.concat(args, " "))
	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})
	return result and result.status == 0
end

local function coerce_bool(value, fallback)
	if type(value) == "boolean" then
		return value
	end
	if type(value) == "string" then
		local normalized = value:lower()
		if normalized == "yes" or normalized == "true" or normalized == "1" or normalized == "on" then
			return true
		end
		if normalized == "no" or normalized == "false" or normalized == "0" or normalized == "off" then
			return false
		end
	end
	return fallback
end

local function parse_start_script_message_overrides(...)
	local overrides = {}
	for i = 1, select("#", ...) do
		local token = select(i, ...)
		if type(token) == "string" and token ~= "" then
			local key, value = token:match("^([%w_%-]+)=(.+)$")
			if key and value then
				local normalized_key = key:lower()
				if normalized_key == "backend" then
					local backend = value:lower()
					if backend == "auto" or backend == "hyprland" or backend == "sway" or backend == "x11" or backend == "macos" then
						overrides.backend = backend
					end
				elseif normalized_key == "socket" or normalized_key == "socket_path" then
					overrides.socket_path = value
				elseif normalized_key == "texthooker" or normalized_key == "texthooker_enabled" then
					local parsed = coerce_bool(value, nil)
					if parsed ~= nil then
						overrides.texthooker_enabled = parsed
					end
				elseif normalized_key == "log-level" or normalized_key == "log_level" then
					overrides.log_level = normalize_log_level(value)
				end
			end
		end
	end
	return overrides
end

local function resolve_visible_overlay_startup()
	local visible = coerce_bool(opts.auto_start_visible_overlay, false)
	-- Backward compatibility for old config key.
	if coerce_bool(opts.auto_start_overlay, false) then
		visible = true
	end
	return visible
end

local function resolve_invisible_overlay_startup()
	local raw = opts.auto_start_invisible_overlay
	if type(raw) == "boolean" then
		return raw
	end

	local mode = type(raw) == "string" and raw:lower() or "platform-default"
	if mode == "visible" or mode == "show" or mode == "yes" or mode == "true" or mode == "on" then
		return true
	end
	if mode == "hidden" or mode == "hide" or mode == "no" or mode == "false" or mode == "off" then
		return false
	end

	-- platform-default
	return not is_linux()
end

local function apply_startup_overlay_preferences()
	local should_show_visible = resolve_visible_overlay_startup()
	local should_show_invisible = resolve_invisible_overlay_startup()

	local visible_action = should_show_visible and "show-visible-overlay" or "hide-visible-overlay"
	if not run_control_command(visible_action) then
		subminer_log("warn", "process", "Failed to apply visible startup action: " .. visible_action)
	end

	local invisible_action = should_show_invisible and "show-invisible-overlay" or "hide-invisible-overlay"
	if not run_control_command(invisible_action) then
		subminer_log("warn", "process", "Failed to apply invisible startup action: " .. invisible_action)
	end

	state.invisible_overlay_visible = should_show_invisible
end

local function build_texthooker_args()
	local args = { state.binary_path, "--texthooker", "--port", tostring(opts.texthooker_port) }
	local log_level = normalize_log_level(opts.log_level)
	if log_level ~= "info" then
		table.insert(args, "--log-level")
		table.insert(args, log_level)
	end
	return args
end

local function ensure_texthooker_running(callback)
	if not opts.texthooker_enabled then
		callback()
		return
	end

	if state.texthooker_running then
		callback()
		return
	end

	local args = build_texthooker_args()
	subminer_log("info", "texthooker", "Starting texthooker process: " .. table.concat(args, " "))
	state.texthooker_running = true

	mp.command_native_async({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	}, function(success, result, error)
		if not success or (result and result.status ~= 0) then
			state.texthooker_running = false
			subminer_log(
				"warn",
				"texthooker",
				"Texthooker process exited unexpectedly: " .. (error or (result and result.stderr) or "unknown error")
			)
		end
	end)

	-- Give the process a moment to acquire the app lock before sending --start.
	mp.add_timeout(0.35, callback)
end

local function start_overlay(overrides)
	if not ensure_binary_available() then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	if state.overlay_running then
		subminer_log("info", "process", "Overlay already running")
		show_osd("Already running")
		return
	end

	overrides = overrides or {}
	local texthooker_enabled = overrides.texthooker_enabled
	if texthooker_enabled == nil then
		texthooker_enabled = opts.texthooker_enabled
	end

	local function launch_overlay()
		local args = build_command_args("start", overrides)
		subminer_log("info", "process", "Starting overlay: " .. table.concat(args, " "))

		show_osd("Starting...")
		state.overlay_running = true

		mp.command_native_async({
			name = "subprocess",
			args = args,
			playback_only = false,
			capture_stdout = true,
			capture_stderr = true,
		}, function(success, result, error)
			if not success or (result and result.status ~= 0) then
				state.overlay_running = false
				subminer_log(
					"error",
					"process",
					"Overlay start failed: " .. (error or (result and result.stderr) or "unknown error")
				)
				show_osd("Overlay start failed")
			end
		end)

		-- Apply explicit startup visibility for each overlay layer.
		mp.add_timeout(0.6, function()
			apply_startup_overlay_preferences()
		end)
	end

	if texthooker_enabled then
		ensure_texthooker_running(launch_overlay)
	else
		launch_overlay()
	end
end

local function start_overlay_from_script_message(...)
	local overrides = parse_start_script_message_overrides(...)
	start_overlay(overrides)
end

local function stop_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("stop")
	subminer_log("info", "process", "Stopping overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	state.overlay_running = false
	state.texthooker_running = false
	if result.status == 0 then
		subminer_log("info", "process", "Overlay stopped")
	else
		subminer_log("warn", "process", "Stop command returned non-zero status: " .. tostring(result.status))
	end
	show_osd("Stopped")
end

local function toggle_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("toggle")
	subminer_log("info", "process", "Toggling overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Toggle command failed")
		show_osd("Toggle failed")
	end
end

local function toggle_invisible_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("toggle-invisible-overlay")
	subminer_log("info", "process", "Toggling invisible overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Invisible toggle command failed")
		show_osd("Invisible toggle failed")
		return
	end

	state.invisible_overlay_visible = not state.invisible_overlay_visible
	show_osd("Invisible overlay: " .. (state.invisible_overlay_visible and "visible" or "hidden"))
end

local function show_invisible_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("show-invisible-overlay")
	subminer_log("info", "process", "Showing invisible overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Show invisible command failed")
		show_osd("Show invisible failed")
		return
	end

	state.invisible_overlay_visible = true
	show_osd("Invisible overlay: visible")
end

local function hide_invisible_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("hide-invisible-overlay")
	subminer_log("info", "process", "Hiding invisible overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Hide invisible command failed")
		show_osd("Hide invisible failed")
		return
	end

	state.invisible_overlay_visible = false
	show_osd("Invisible overlay: hidden")
end

local function open_options()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end
	local args = build_command_args("settings")
	subminer_log("info", "process", "Opening options: " .. table.concat(args, " "))
	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})
	if result.status == 0 then
		subminer_log("info", "process", "Options window opened")
		show_osd("Options opened")
	else
		subminer_log("warn", "process", "Failed to open options")
		show_osd("Failed to open options")
	end
end

local restart_overlay
local check_status

local function show_menu()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local items = {
		"Start overlay",
		"Stop overlay",
		"Toggle overlay",
		"Toggle invisible overlay",
		"Open options",
		"Restart overlay",
		"Check status",
	}

	local actions = {
		start_overlay,
		stop_overlay,
		toggle_overlay,
		toggle_invisible_overlay,
		open_options,
		restart_overlay,
		check_status,
	}

	input.select({
		prompt = "SubMiner: ",
		items = items,
		submit = function(index)
			if index and actions[index] then
				actions[index]()
			end
		end,
	})
end

restart_overlay = function()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	subminer_log("info", "process", "Restarting overlay...")
	show_osd("Restarting...")

	local stop_args = build_command_args("stop")
	mp.command_native({
		name = "subprocess",
		args = stop_args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	state.overlay_running = false
	state.texthooker_running = false

	ensure_texthooker_running(function()
		local start_args = build_command_args("start")
		subminer_log("info", "process", "Starting overlay: " .. table.concat(start_args, " "))

		state.overlay_running = true
		mp.command_native_async({
			name = "subprocess",
			args = start_args,
			playback_only = false,
			capture_stdout = true,
			capture_stderr = true,
		}, function(success, result, error)
			if not success or (result and result.status ~= 0) then
				state.overlay_running = false
				subminer_log(
					"error",
					"process",
					"Overlay start failed: " .. (error or (result and result.stderr) or "unknown error")
				)
				show_osd("Restart failed")
			else
				show_osd("Restarted successfully")
			end
		end)
	end)
end

check_status = function()
	if not state.binary_available then
		show_osd("Status: binary not found")
		return
	end

	local status = state.overlay_running and "running" or "stopped"
	show_osd("Status: overlay is " .. status)
	subminer_log("info", "process", "Status check: overlay is " .. status)
end

local function on_file_loaded()
	state.binary_path = find_binary()
	if state.binary_path then
		state.binary_available = true
		subminer_log("info", "lifecycle", "SubMiner ready (binary: " .. state.binary_path .. ")")
		local should_auto_start = coerce_bool(opts.auto_start, false)
		if should_auto_start then
			start_overlay()
		end
	else
		state.binary_available = false
		subminer_log("warn", "binary", "SubMiner binary not found - overlay features disabled")
		if opts.binary_path ~= "" then
			subminer_log("warn", "binary", "Configured path '" .. opts.binary_path .. "' does not exist")
		end
	end
end

local function on_shutdown()
	if (state.overlay_running or state.texthooker_running) and state.binary_available then
		subminer_log("info", "lifecycle", "mpv shutting down, stopping SubMiner process")
		show_osd("Shutting down...")
		stop_overlay()
	end
end

local function register_keybindings()
	mp.add_key_binding("y-s", "subminer-start", start_overlay)
	mp.add_key_binding("y-S", "subminer-stop", stop_overlay)
	mp.add_key_binding("y-t", "subminer-toggle", toggle_overlay)
	mp.add_key_binding("y-i", "subminer-toggle-invisible", toggle_invisible_overlay)
	mp.add_key_binding("y-I", "subminer-show-invisible", show_invisible_overlay)
	mp.add_key_binding("y-u", "subminer-hide-invisible", hide_invisible_overlay)
	mp.add_key_binding("y-y", "subminer-menu", show_menu)
	mp.add_key_binding("y-o", "subminer-options", open_options)
	mp.add_key_binding("y-r", "subminer-restart", restart_overlay)
	mp.add_key_binding("y-c", "subminer-status", check_status)
end

local function register_script_messages()
	mp.register_script_message("subminer-start", start_overlay_from_script_message)
	mp.register_script_message("subminer-stop", stop_overlay)
	mp.register_script_message("subminer-toggle", toggle_overlay)
	mp.register_script_message("subminer-toggle-invisible", toggle_invisible_overlay)
	mp.register_script_message("subminer-show-invisible", show_invisible_overlay)
	mp.register_script_message("subminer-hide-invisible", hide_invisible_overlay)
	mp.register_script_message("subminer-menu", show_menu)
	mp.register_script_message("subminer-options", open_options)
	mp.register_script_message("subminer-restart", restart_overlay)
	mp.register_script_message("subminer-status", check_status)
end

local function init()
	register_keybindings()
	register_script_messages()

	mp.register_event("file-loaded", on_file_loaded)
	mp.register_event("shutdown", on_shutdown)

	subminer_log("info", "lifecycle", "SubMiner plugin loaded")
end

init()
