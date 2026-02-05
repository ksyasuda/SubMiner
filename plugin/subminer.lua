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

local opts = {
	binary_path = "",
	socket_path = default_socket_path(),
	texthooker_enabled = true,
	texthooker_port = 5174,
	backend = "auto",
	auto_start = false,
	auto_start_overlay = true,
	osd_messages = true,
}

options.read_options(opts, "subminer")

local state = {
	overlay_running = false,
	overlay_process = nil,
	binary_available = false,
	binary_path = nil,
	detected_backend = nil,
}

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
		msg.warn("Could not detect window manager, falling back to x11")
		backend = "x11"
	end

	state.detected_backend = backend
	if backend then
		msg.info("Detected backend: " .. backend)
	else
		msg.info("No backend detected")
	end
	return backend
end

local function file_exists(path)
	local info = utils.file_info(path)
	return info ~= nil
end

local function find_binary()
	if opts.binary_path ~= "" and file_exists(opts.binary_path) then
		return opts.binary_path
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
			msg.info("Found binary at: " .. path)
			return path
		end
	end

	return nil
end

local function build_command_args(action)
	local args = { state.binary_path }

	table.insert(args, "--" .. action)

	if action == "start" then
		local backend = opts.backend == "auto" and detect_backend() or opts.backend
		if backend and backend ~= "" then
			table.insert(args, "--backend")
			table.insert(args, backend)
		end

		if opts.texthooker_enabled then
			table.insert(args, "--texthooker")
			table.insert(args, "--port")
			table.insert(args, tostring(opts.texthooker_port))
		end

		table.insert(args, "--socket")
		table.insert(args, opts.socket_path)

		if opts.auto_start_overlay then
			table.insert(args, "--auto-start-overlay")
		end
	end

	return args
end

local function start_overlay()
	if not state.binary_available then
		msg.error("SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	if state.overlay_running then
		msg.info("Overlay already running")
		show_osd("Already running")
		return
	end

	local args = build_command_args("start")
	msg.info("Starting overlay: " .. table.concat(args, " "))

	show_osd("Starting...")
	state.overlay_running = true

	mp.command_native_async({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	}, function(success, result, error)
		state.overlay_running = false
		if not success or (result and result.status ~= 0) then
			msg.error("Overlay stopped unexpectedly: " .. (error or (result and result.stderr) or "unknown error"))
			show_osd("Overlay stopped unexpectedly")
		end
	end)
end

local function stop_overlay()
	if not state.binary_available then
		msg.error("SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("stop")
	msg.info("Stopping overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	state.overlay_running = false
	if result.status == 0 then
		msg.info("Overlay stopped")
	else
		msg.warn("Stop command returned non-zero status: " .. tostring(result.status))
	end
	show_osd("Stopped")
end

local function toggle_overlay()
	if not state.binary_available then
		msg.error("SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("toggle")
	msg.info("Toggling overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		msg.warn("Toggle command failed")
		show_osd("Toggle failed")
	end
end

local function open_options()
	if not state.binary_available then
		msg.error("SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end
	local args = build_command_args("settings")
	msg.info("Opening options: " .. table.concat(args, " "))
	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})
	if result.status == 0 then
		msg.info("Options window opened")
		show_osd("Options opened")
	else
		msg.warn("Failed to open options")
		show_osd("Failed to open options")
	end
end

local restart_overlay
local check_status

local function show_menu()
	if not state.binary_available then
		msg.error("SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local items = {
		"Start overlay",
		"Stop overlay",
		"Toggle overlay",
		"Open options",
		"Restart overlay",
		"Check status",
	}

	local actions = {
		start_overlay,
		stop_overlay,
		toggle_overlay,
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
		msg.error("SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	msg.info("Restarting overlay...")
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

	local start_args = build_command_args("start")
	msg.info("Starting overlay: " .. table.concat(start_args, " "))

	state.overlay_running = true
	mp.command_native_async({
		name = "subprocess",
		args = start_args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	}, function(success, result, error)
		state.overlay_running = false
		if not success or (result and result.status ~= 0) then
			msg.error("Overlay stopped unexpectedly: " .. (error or (result and result.stderr) or "unknown error"))
			show_osd("Restart failed - overlay stopped unexpectedly")
		else
			show_osd("Restarted successfully")
		end
	end)
end

check_status = function()
	if not state.binary_available then
		show_osd("Status: binary not found")
		return
	end

	local status = state.overlay_running and "running" or "stopped"
	show_osd("Status: overlay is " .. status)
	msg.info("Status check: overlay is " .. status)
end

local function on_file_loaded()
	state.binary_path = find_binary()
	if state.binary_path then
		state.binary_available = true
		msg.info("SubMiner ready (binary: " .. state.binary_path .. ")")

		if opts.auto_start then
			start_overlay()
		end
	else
		state.binary_available = false
		msg.warn("SubMiner binary not found - overlay features disabled")
		if opts.binary_path ~= "" then
			msg.warn("Configured path '" .. opts.binary_path .. "' does not exist")
		end
	end
end

local function on_shutdown()
	if state.overlay_running and state.binary_available then
		msg.info("mpv shutting down, stopping overlay")
		show_osd("Shutting down...")
		stop_overlay()
	end
end

local function register_keybindings()
	mp.add_key_binding("y-s", "subminer-start", start_overlay)
	mp.add_key_binding("y-S", "subminer-stop", stop_overlay)
	mp.add_key_binding("y-t", "subminer-toggle", toggle_overlay)
	mp.add_key_binding("y-y", "subminer-menu", show_menu)
	mp.add_key_binding("y-o", "subminer-options", open_options)
	mp.add_key_binding("y-r", "subminer-restart", restart_overlay)
	mp.add_key_binding("y-c", "subminer-status", check_status)
end

local function register_script_messages()
	mp.register_script_message("subminer-start", start_overlay)
	mp.register_script_message("subminer-stop", stop_overlay)
	mp.register_script_message("subminer-toggle", toggle_overlay)
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

	msg.info("SubMiner plugin loaded")
end

init()
