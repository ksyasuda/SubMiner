local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local opts = ctx.opts
	local state = ctx.state
	local binary = ctx.binary
	local environment = ctx.environment
	local options_helper = ctx.options_helper
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd
	local normalize_log_level = ctx.log.normalize_log_level

	local function resolve_backend(override_backend)
		local selected = override_backend
		if selected == nil or selected == "" then
			selected = opts.backend
		end
		if selected == "auto" then
			return environment.detect_backend()
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
						local parsed = options_helper.coerce_bool(value, nil)
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
		local visible = options_helper.coerce_bool(opts.auto_start_visible_overlay, false)
		if options_helper.coerce_bool(opts.auto_start_overlay, false) then
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

		return not environment.is_linux()
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
				subminer_log("warn", "texthooker", "Texthooker process exited unexpectedly: " .. (error or (result and result.stderr) or "unknown error"))
			end
		end)

		mp.add_timeout(0.35, callback)
	end

	local function start_overlay(overrides)
		if not binary.ensure_binary_available() then
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
					subminer_log("error", "process", "Overlay start failed: " .. (error or (result and result.stderr) or "unknown error"))
					show_osd("Overlay start failed")
				end
			end)

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
		if not binary.ensure_binary_available() then
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
		if not binary.ensure_binary_available() then
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
		if not binary.ensure_binary_available() then
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
		if not binary.ensure_binary_available() then
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
		if not binary.ensure_binary_available() then
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

	local function restart_overlay()
		if not binary.ensure_binary_available() then
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
					subminer_log("error", "process", "Overlay start failed: " .. (error or (result and result.stderr) or "unknown error"))
					show_osd("Restart failed")
				else
					show_osd("Restarted successfully")
				end
			end)
		end)
	end

	local function check_status()
		if not state.binary_available then
			show_osd("Status: binary not found")
			return
		end
		local status = state.overlay_running and "running" or "stopped"
		show_osd("Status: overlay is " .. status)
		subminer_log("info", "process", "Status check: overlay is " .. status)
	end

	return {
		build_command_args = build_command_args,
		run_control_command = run_control_command,
		parse_start_script_message_overrides = parse_start_script_message_overrides,
		apply_startup_overlay_preferences = apply_startup_overlay_preferences,
		ensure_texthooker_running = ensure_texthooker_running,
		start_overlay = start_overlay,
		start_overlay_from_script_message = start_overlay_from_script_message,
		stop_overlay = stop_overlay,
		toggle_overlay = toggle_overlay,
		toggle_invisible_overlay = toggle_invisible_overlay,
		show_invisible_overlay = show_invisible_overlay,
		hide_invisible_overlay = hide_invisible_overlay,
		open_options = open_options,
		restart_overlay = restart_overlay,
		check_status = check_status,
	}
end

return M
