local M = {}

local OVERLAY_START_RETRY_DELAY_SECONDS = 0.2
local OVERLAY_START_MAX_ATTEMPTS = 6
local AUTO_PLAY_READY_LOADING_OSD = "Loading subtitle tokenization..."
local AUTO_PLAY_READY_READY_OSD = "Subtitle tokenization ready"
local DEFAULT_AUTO_PLAY_READY_TIMEOUT_SECONDS = 15

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
	local run_control_command_async

	local function resolve_visible_overlay_startup()
		local raw_visible_overlay = opts.auto_start_visible_overlay
		if raw_visible_overlay == nil then
			raw_visible_overlay = opts["auto-start-visible-overlay"]
		end
		return options_helper.coerce_bool(raw_visible_overlay, false)
	end

	local function resolve_pause_until_ready()
		local raw_pause_until_ready = opts.auto_start_pause_until_ready
		if raw_pause_until_ready == nil then
			raw_pause_until_ready = opts["auto-start-pause-until-ready"]
		end
		return options_helper.coerce_bool(raw_pause_until_ready, false)
	end

	local function resolve_pause_until_ready_timeout_seconds()
		local raw_timeout_seconds = opts.auto_start_pause_until_ready_timeout_seconds
		if raw_timeout_seconds == nil then
			raw_timeout_seconds = opts["auto-start-pause-until-ready-timeout-seconds"]
		end
		if type(raw_timeout_seconds) == "number" then
			return raw_timeout_seconds
		end
		if type(raw_timeout_seconds) == "string" then
			local parsed = tonumber(raw_timeout_seconds)
			if parsed ~= nil then
				return parsed
			end
		end
		return DEFAULT_AUTO_PLAY_READY_TIMEOUT_SECONDS
	end

	local function normalize_socket_path(path)
		if type(path) ~= "string" then
			return nil
		end
		local trimmed = path:match("^%s*(.-)%s*$")
		if trimmed == "" then
			return nil
		end
		return trimmed
	end

	local function has_matching_mpv_ipc_socket(target_socket_path)
		local expected_socket = normalize_socket_path(target_socket_path or opts.socket_path)
		local active_socket = normalize_socket_path(mp.get_property("input-ipc-server"))
		if expected_socket == nil or active_socket == nil then
			return false
		end
		return expected_socket == active_socket
	end

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

	local function clear_auto_play_ready_timeout()
		local timeout = state.auto_play_ready_timeout
		if timeout and timeout.kill then
			timeout:kill()
		end
		state.auto_play_ready_timeout = nil
	end

	local function clear_auto_play_ready_osd_timer()
		local timer = state.auto_play_ready_osd_timer
		if timer and timer.kill then
			timer:kill()
		end
		state.auto_play_ready_osd_timer = nil
	end

	local function disarm_auto_play_ready_gate(options)
		local should_resume = options == nil or options.resume_playback ~= false
		local was_armed = state.auto_play_ready_gate_armed
		clear_auto_play_ready_timeout()
		clear_auto_play_ready_osd_timer()
		state.auto_play_ready_gate_armed = false
		if was_armed and should_resume then
			mp.set_property_native("pause", false)
		end
	end

	local function release_auto_play_ready_gate(reason)
		if not state.auto_play_ready_gate_armed then
			return
		end
		disarm_auto_play_ready_gate({ resume_playback = false })
		mp.set_property_native("pause", false)
		show_osd(AUTO_PLAY_READY_READY_OSD)
		subminer_log("info", "process", "Resuming playback after startup gate: " .. tostring(reason or "ready"))
	end

	local function arm_auto_play_ready_gate()
		if state.auto_play_ready_gate_armed then
			clear_auto_play_ready_timeout()
			clear_auto_play_ready_osd_timer()
		end
		state.auto_play_ready_gate_armed = true
		mp.set_property_native("pause", true)
		show_osd(AUTO_PLAY_READY_LOADING_OSD)
		if type(mp.add_periodic_timer) == "function" then
			state.auto_play_ready_osd_timer = mp.add_periodic_timer(2.5, function()
				if state.auto_play_ready_gate_armed then
					show_osd(AUTO_PLAY_READY_LOADING_OSD)
				end
			end)
		end
		subminer_log("info", "process", "Pausing playback until SubMiner overlay/tokenization readiness signal")
		local timeout_seconds = resolve_pause_until_ready_timeout_seconds()
		if timeout_seconds and timeout_seconds > 0 then
			state.auto_play_ready_timeout = mp.add_timeout(timeout_seconds, function()
				if not state.auto_play_ready_gate_armed then
					return
				end
				subminer_log(
					"warn",
					"process",
					"Startup readiness signal timed out; resuming playback to avoid stalled pause"
				)
				release_auto_play_ready_gate("timeout")
			end)
		end
	end

	local function notify_auto_play_ready()
		release_auto_play_ready_gate("tokenization-ready")
		if state.suppress_ready_overlay_restore then
			return
		end
		if state.overlay_running and resolve_visible_overlay_startup() then
			run_control_command_async("show-visible-overlay", {
				socket_path = opts.socket_path,
			})
		end
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

		if action == "start" then
			local backend = resolve_backend(overrides.backend)
			if backend and backend ~= "" then
				table.insert(args, "--backend")
				table.insert(args, backend)
			end

			local socket_path = overrides.socket_path or opts.socket_path
			table.insert(args, "--socket")
			table.insert(args, socket_path)

			local should_show_visible = resolve_visible_overlay_startup()
			if should_show_visible then
				table.insert(args, "--show-visible-overlay")
			else
				table.insert(args, "--hide-visible-overlay")
			end

			local texthooker_enabled = overrides.texthooker_enabled
			if texthooker_enabled == nil then
				texthooker_enabled = opts.texthooker_enabled
			end
			if texthooker_enabled then
				table.insert(args, "--texthooker")
			end
		end

		return args
	end

	run_control_command_async = function(action, overrides, callback)
		local args = build_command_args(action, overrides)
		subminer_log("debug", "process", "Control command: " .. table.concat(args, " "))
		mp.command_native_async({
			name = "subprocess",
			args = args,
			playback_only = false,
			capture_stdout = true,
			capture_stderr = true,
		}, function(success, result, error)
			local ok = success and (result == nil or result.status == 0)
			if callback then
				callback(ok, result, error)
			end
		end)
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

	local function ensure_texthooker_running(callback)
		if callback then
			callback()
		end
	end

	local function start_overlay(overrides)
		overrides = overrides or {}
		if overrides.auto_start_trigger == true then
			state.suppress_ready_overlay_restore = false
		end

		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end

		if state.overlay_running then
			if overrides.auto_start_trigger == true then
				subminer_log("debug", "process", "Auto-start ignored because overlay is already running")
				local socket_path = overrides.socket_path or opts.socket_path
				local should_pause_until_ready = (
					resolve_visible_overlay_startup()
					and resolve_pause_until_ready()
					and has_matching_mpv_ipc_socket(socket_path)
				)
				if should_pause_until_ready then
					arm_auto_play_ready_gate()
				else
					disarm_auto_play_ready_gate()
				end
				local visibility_action = resolve_visible_overlay_startup()
						and "show-visible-overlay"
					or "hide-visible-overlay"
				run_control_command_async(visibility_action, {
					socket_path = socket_path,
					log_level = overrides.log_level,
				})
				return
			end
			subminer_log("info", "process", "Overlay already running")
			show_osd("Already running")
			return
		end

		local texthooker_enabled = overrides.texthooker_enabled
		if texthooker_enabled == nil then
			texthooker_enabled = opts.texthooker_enabled
		end
		local socket_path = overrides.socket_path or opts.socket_path
		local should_pause_until_ready = (
			overrides.auto_start_trigger == true
			and resolve_visible_overlay_startup()
			and resolve_pause_until_ready()
			and has_matching_mpv_ipc_socket(socket_path)
		)
		if should_pause_until_ready then
			arm_auto_play_ready_gate()
		else
			disarm_auto_play_ready_gate()
		end

		local function launch_overlay_with_retry(attempt)
			local args = build_command_args("start", overrides)
			if attempt == 1 then
				subminer_log("info", "process", "Starting overlay: " .. table.concat(args, " "))
			else
				subminer_log(
					"warn",
					"process",
					"Retrying overlay start (attempt " .. tostring(attempt) .. "): " .. table.concat(args, " ")
				)
			end

			if attempt == 1 and not state.auto_play_ready_gate_armed then
				show_osd("Starting...")
			end
			state.overlay_running = true

			mp.command_native_async({
				name = "subprocess",
				args = args,
				playback_only = false,
				capture_stdout = true,
				capture_stderr = true,
			}, function(success, result, error)
				if not success or (result and result.status ~= 0) then
					local reason = error or (result and result.stderr) or "unknown error"
					if attempt < OVERLAY_START_MAX_ATTEMPTS then
						mp.add_timeout(OVERLAY_START_RETRY_DELAY_SECONDS, function()
							launch_overlay_with_retry(attempt + 1)
						end)
						return
					end

					state.overlay_running = false
					subminer_log("error", "process", "Overlay start failed after retries: " .. reason)
					show_osd("Overlay start failed")
					release_auto_play_ready_gate("overlay-start-failed")
					return
				end

				if overrides.auto_start_trigger == true then
					local visibility_action = resolve_visible_overlay_startup()
							and "show-visible-overlay"
						or "hide-visible-overlay"
						run_control_command_async(visibility_action, {
							socket_path = socket_path,
							log_level = overrides.log_level,
						})
				end

			end)
		end

		launch_overlay_with_retry(1)
		if texthooker_enabled then
			ensure_texthooker_running(function() end)
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

		run_control_command_async("stop", nil, function(ok, result)
			if ok then
				subminer_log("info", "process", "Overlay stopped")
			else
				subminer_log(
					"warn",
					"process",
					"Stop command returned non-zero status: " .. tostring(result and result.status or "unknown")
				)
			end
		end)

		state.overlay_running = false
		state.texthooker_running = false
		disarm_auto_play_ready_gate()
		show_osd("Stopped")
	end

	local function hide_visible_overlay()
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			return
		end
		state.suppress_ready_overlay_restore = true

		run_control_command_async("hide-visible-overlay", nil, function(ok, result)
			if ok then
				subminer_log("info", "process", "Visible overlay hidden")
			else
				subminer_log(
					"warn",
					"process",
					"Hide-visible-overlay command returned non-zero status: "
						.. tostring(result and result.status or "unknown")
				)
			end
		end)

		disarm_auto_play_ready_gate()
	end

	local function toggle_overlay()
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end
		state.suppress_ready_overlay_restore = true

		run_control_command_async("toggle-visible-overlay", nil, function(ok)
			if not ok then
				subminer_log("warn", "process", "Toggle command failed")
				show_osd("Toggle failed")
			end
		end)
	end

	local function open_options()
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end

		run_control_command_async("settings", nil, function(ok)
			if ok then
				subminer_log("info", "process", "Options window opened")
				show_osd("Options opened")
			else
				subminer_log("warn", "process", "Failed to open options")
				show_osd("Failed to open options")
			end
		end)
	end

	local function restart_overlay()
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end

		subminer_log("info", "process", "Restarting overlay...")
		show_osd("Restarting...")

		run_control_command_async("stop", nil, function()
			state.overlay_running = false
			state.texthooker_running = false
			disarm_auto_play_ready_gate()

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

			if opts.texthooker_enabled then
				ensure_texthooker_running(function() end)
			end
		end)
	end

	local function check_status()
		if not binary.ensure_binary_available() then
			show_osd("Status: binary not found")
			return
		end

		local status = state.overlay_running and "running" or "stopped"
		show_osd("Status: overlay is " .. status)
		subminer_log("info", "process", "Status check: overlay is " .. status)
	end

	local function check_binary_available()
		return binary.ensure_binary_available()
	end

	return {
		build_command_args = build_command_args,
		has_matching_mpv_ipc_socket = has_matching_mpv_ipc_socket,
		run_control_command_async = run_control_command_async,
		parse_start_script_message_overrides = parse_start_script_message_overrides,
		ensure_texthooker_running = ensure_texthooker_running,
		start_overlay = start_overlay,
		start_overlay_from_script_message = start_overlay_from_script_message,
		stop_overlay = stop_overlay,
		hide_visible_overlay = hide_visible_overlay,
		toggle_overlay = toggle_overlay,
		open_options = open_options,
		restart_overlay = restart_overlay,
		check_status = check_status,
		check_binary_available = check_binary_available,
		notify_auto_play_ready = notify_auto_play_ready,
		disarm_auto_play_ready_gate = disarm_auto_play_ready_gate,
	}
end

return M
