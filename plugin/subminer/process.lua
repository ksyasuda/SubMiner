local M = {}

local OVERLAY_START_RETRY_DELAY_SECONDS = 0.2
local OVERLAY_START_MAX_ATTEMPTS = 6
local OVERLAY_RESTART_PING_RETRY_DELAY_SECONDS = 0.2
local OVERLAY_RESTART_PING_MAX_ATTEMPTS = 20
local OVERLAY_LOADING_OSD_PREFIX = "Overlay loading "
local OVERLAY_LOADING_OSD_FRAMES = { "|", "/", "-", "\\" }
local OVERLAY_LOADING_OSD_REFRESH_SECONDS = 0.18
local AUTO_PLAY_READY_LOADING_OSD = "Loading subtitle tokenization..."
local AUTO_PLAY_READY_READY_OSD = "Subtitle tokenization ready"
local DEFAULT_AUTO_PLAY_READY_TIMEOUT_SECONDS = 30
local DUPLICATE_VISIBLE_OVERLAY_TOGGLE_SECONDS = 0.25

function M.create(ctx)
	local mp = ctx.mp
	local utils = ctx.utils
	local opts = ctx.opts
	local state = ctx.state
	local binary = ctx.binary
	local environment = ctx.environment
	local options_helper = ctx.options_helper
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd
	local normalize_log_level = ctx.log.normalize_log_level
	local run_control_command_async
	local APP_ARGC_ENV = "SUBMINER_APP_ARGC"
	local APP_ARG_PREFIX = "SUBMINER_APP_ARG_"

	local function resolve_visible_overlay_startup()
		local raw_visible_overlay = opts.auto_start_visible_overlay
		if raw_visible_overlay == nil then
			raw_visible_overlay = opts["auto-start-visible-overlay"]
		end
		return options_helper.coerce_bool(raw_visible_overlay, false)
	end

	local function resolve_auto_start_visibility_action()
		if resolve_visible_overlay_startup() then
			if state.suppress_ready_overlay_restore then
				return nil
			end
			return "show-visible-overlay"
		end
		if state.visible_overlay_requested == true then
			return nil
		end
		return "hide-visible-overlay"
	end

	local function resolve_pause_until_ready()
		local raw_pause_until_ready = opts.auto_start_pause_until_ready
		if raw_pause_until_ready == nil then
			raw_pause_until_ready = opts["auto-start-pause-until-ready"]
		end
		return options_helper.coerce_bool(raw_pause_until_ready, false)
	end

	local function resolve_osd_messages_enabled()
		local raw_osd_messages = opts.osd_messages
		if raw_osd_messages == nil then
			raw_osd_messages = opts["osd-messages"]
		end
		return options_helper.coerce_bool(raw_osd_messages, false)
	end

	local function resolve_pause_until_ready_owns_initial_pause()
		local raw_owns_initial_pause = opts.auto_start_pause_until_ready_owns_initial_pause
		if raw_owns_initial_pause == nil then
			raw_owns_initial_pause = opts["auto-start-pause-until-ready-owns-initial-pause"]
		end
		return options_helper.coerce_bool(raw_owns_initial_pause, false)
	end

	local function consume_pause_until_ready_initial_pause_ownership()
		if state.auto_play_ready_initial_pause_ownership_consumed then
			return false
		end
		if not resolve_pause_until_ready_owns_initial_pause() then
			return false
		end
		state.auto_play_ready_initial_pause_ownership_consumed = true
		return true
	end

	local function resolve_texthooker_enabled(override_value)
		if override_value ~= nil then
			return options_helper.coerce_bool(override_value, false)
		end
		local raw_texthooker_enabled = opts.texthooker_enabled
		if raw_texthooker_enabled == nil then
			raw_texthooker_enabled = opts["texthooker-enabled"]
		end
		return options_helper.coerce_bool(raw_texthooker_enabled, false)
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

	local function record_visible_overlay_action(action)
		if action == "show-visible-overlay" then
			state.visible_overlay_requested = true
			state.suppress_ready_overlay_restore = false
		elseif action == "hide-visible-overlay" then
			state.visible_overlay_requested = false
		elseif action == "toggle-visible-overlay" and state.visible_overlay_requested ~= nil then
			state.visible_overlay_requested = not state.visible_overlay_requested
			if state.visible_overlay_requested then
				state.suppress_ready_overlay_restore = false
			end
		end
	end

	local function record_visible_overlay_visibility(visible)
		if visible then
			state.visible_overlay_requested = true
			state.suppress_ready_overlay_restore = false
			return
		end
		state.visible_overlay_requested = false
		state.suppress_ready_overlay_restore = true
	end

	local function record_start_visibility_args(args)
		for _, arg in ipairs(args) do
			if arg == "--show-visible-overlay" then
				record_visible_overlay_action("show-visible-overlay")
				return
			end
			if arg == "--hide-visible-overlay" then
				record_visible_overlay_action("hide-visible-overlay")
				return
			end
		end
	end

	local function should_run_visibility_action(action)
		if action == "show-visible-overlay" and state.visible_overlay_requested == true then
			return false
		end
		if action == "hide-visible-overlay" and state.visible_overlay_requested == false then
			return false
		end
		return true
	end

	local function run_visibility_action_if_needed(action, overrides, callback)
		if action == nil then
			if callback then
				callback(true)
			end
			return
		end
		if not should_run_visibility_action(action) then
			subminer_log("debug", "process", "Skipping duplicate visible overlay action: " .. tostring(action))
			if callback then
				callback(true)
			end
			return
		end
		run_control_command_async(action, overrides, callback)
	end

	local function should_ignore_duplicate_visible_overlay_toggle()
		if type(mp.get_time) ~= "function" then
			return false
		end
		local now = mp.get_time()
		if type(now) ~= "number" then
			return false
		end

		local previous = state.last_visible_overlay_toggle_time
		state.last_visible_overlay_toggle_time = now
		if type(previous) ~= "number" then
			return false
		end

		local elapsed = now - previous
		return elapsed >= 0 and elapsed < DUPLICATE_VISIBLE_OVERLAY_TOGGLE_SECONDS
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

	local function get_mpv_ipc_socket_match(target_socket_path)
		local expected_socket = normalize_socket_path(target_socket_path or opts.socket_path)
		local active_socket = normalize_socket_path(mp.get_property("input-ipc-server"))
		return {
			expected_socket = expected_socket,
			active_socket = active_socket,
			matching = expected_socket ~= nil and active_socket ~= nil and expected_socket == active_socket,
		}
	end

	local function has_matching_mpv_ipc_socket(target_socket_path)
		local match = get_mpv_ipc_socket_match(target_socket_path)
		return match.matching
	end

	local function describe_mpv_ipc_socket_match(target_socket_path)
		local match = get_mpv_ipc_socket_match(target_socket_path)
		local expected_socket = match.expected_socket or "<empty>"
		local active_socket = match.active_socket or "<empty>"
		if expected_socket == nil or active_socket == nil then
			return "expected=" .. expected_socket .. "; active=" .. active_socket .. "; matching=no"
		end
		return "expected=" .. expected_socket .. "; active=" .. active_socket .. "; matching=" .. (match.matching and "yes" or "no")
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

	local function clear_overlay_loading_osd_timer()
		local timer = state.overlay_loading_osd_timer
		if timer and timer.kill then
			timer:kill()
		end
		state.overlay_loading_osd_timer = nil
	end

	local function stop_overlay_loading_osd()
		state.overlay_loading_osd_active = false
		state.overlay_loading_osd_frame = 1
		clear_overlay_loading_osd_timer()
	end

	local function start_overlay_loading_osd()
		if state.overlay_loading_osd_active then
			return
		end
		state.overlay_loading_osd_active = true
		state.overlay_loading_osd_frame = 1
		local function show_next_overlay_loading_frame()
			local frame_index = state.overlay_loading_osd_frame or 1
			local frame = OVERLAY_LOADING_OSD_FRAMES[frame_index] or OVERLAY_LOADING_OSD_FRAMES[1]
			show_osd(OVERLAY_LOADING_OSD_PREFIX .. frame, { force = true })
			state.overlay_loading_osd_frame = (frame_index % #OVERLAY_LOADING_OSD_FRAMES) + 1
		end
		show_next_overlay_loading_frame()
		if type(mp.add_periodic_timer) == "function" then
			state.overlay_loading_osd_timer = mp.add_periodic_timer(OVERLAY_LOADING_OSD_REFRESH_SECONDS, function()
				if state.overlay_loading_osd_active then
					show_next_overlay_loading_frame()
				end
			end)
		end
	end

	local function disarm_auto_play_ready_gate(options)
		local should_resume = options == nil or options.resume_playback ~= false
		local was_armed = state.auto_play_ready_gate_armed
		local should_resume_playback = state.auto_play_ready_should_resume_playback == true
		clear_auto_play_ready_timeout()
		clear_auto_play_ready_osd_timer()
		state.auto_play_ready_gate_armed = false
		state.auto_play_ready_should_resume_playback = false
		if was_armed and should_resume and should_resume_playback then
			mp.set_property_native("pause", false)
		end
	end

	local function release_auto_play_ready_gate(reason)
		if not state.auto_play_ready_gate_armed then
			return false
		end
		local should_resume_playback = state.auto_play_ready_should_resume_playback == true
		if resolve_osd_messages_enabled() then
			stop_overlay_loading_osd()
			show_osd(AUTO_PLAY_READY_READY_OSD)
		end
		disarm_auto_play_ready_gate({ resume_playback = false })
		if should_resume_playback then
			mp.set_property_native("pause", false)
			subminer_log("info", "process", "Resuming playback after startup gate: " .. tostring(reason or "ready"))
		else
			subminer_log("info", "process", "Startup gate ready; leaving playback paused: " .. tostring(reason or "ready"))
		end
		return true
	end

	local function arm_auto_play_ready_gate()
		local was_armed = state.auto_play_ready_gate_armed
		if was_armed then
			clear_auto_play_ready_timeout()
			clear_auto_play_ready_osd_timer()
		end
		if not was_armed then
			state.auto_play_ready_should_resume_playback = consume_pause_until_ready_initial_pause_ownership()
				or mp.get_property_native("pause") ~= true
		end
		state.auto_play_ready_gate_armed = true
		mp.set_property_native("pause", true)
		if resolve_osd_messages_enabled() then
			stop_overlay_loading_osd()
			show_osd(AUTO_PLAY_READY_LOADING_OSD)
		end
		if resolve_osd_messages_enabled() and type(mp.add_periodic_timer) == "function" then
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
		state.auto_play_ready_signal_seen = true
		local released_ready_gate = release_auto_play_ready_gate("tokenization-ready")
		local force_ready_overlay_restore = state.force_ready_overlay_restore == true
		state.force_ready_overlay_restore = false
		if not released_ready_gate and not force_ready_overlay_restore then
			return
		end
		if state.suppress_ready_overlay_restore and not force_ready_overlay_restore then
			return
		end
		if force_ready_overlay_restore then
			state.suppress_ready_overlay_restore = false
		end
		if state.overlay_running and (force_ready_overlay_restore or resolve_visible_overlay_startup()) then
			run_visibility_action_if_needed("show-visible-overlay", {
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
			if overrides.background ~= false then
				table.insert(args, "--background")
			end
			table.insert(args, "--managed-playback")

			local backend = resolve_backend(overrides.backend)
			if backend and backend ~= "" then
				table.insert(args, "--backend")
				table.insert(args, backend)
			end

			local socket_path = overrides.socket_path or opts.socket_path
			table.insert(args, "--socket")
			table.insert(args, socket_path)

			local should_show_visible = overrides.show_visible_overlay
			if should_show_visible == nil then
				should_show_visible = resolve_visible_overlay_startup() and not state.suppress_ready_overlay_restore
			end
			if should_show_visible then
				table.insert(args, "--show-visible-overlay")
			else
				table.insert(args, "--hide-visible-overlay")
			end

			local texthooker_enabled = resolve_texthooker_enabled(overrides.texthooker_enabled)
			if texthooker_enabled then
				table.insert(args, "--texthooker")
			end
		end
		if action == "playback-feedback" and type(overrides.message) == "string" and overrides.message ~= "" then
			table.insert(args, overrides.message)
		end

		return args
	end

	local function is_appimage_binary(path)
		return environment.is_linux() and type(path) == "string" and path:lower():match("%.appimage$") ~= nil
	end

	local function append_transport_env(env, args)
		local count = math.max(#args - 1, 0)
		env[#env + 1] = APP_ARGC_ENV .. "=" .. tostring(count)
		for index = 2, #args do
			env[#env + 1] = APP_ARG_PREFIX .. tostring(index - 2) .. "=" .. tostring(args[index])
		end
	end

	local function env_has_name(env, name)
		local prefix = name .. "="
		for _, value in ipairs(env) do
			if type(value) == "string" and value:sub(1, #prefix) == prefix then
				return true
			end
		end
		return false
	end

	local function append_default_app_log_env(env)
		local log_dir = environment.join_path(environment.resolve_subminer_config_dir(), "logs")
		local date = os.date("%Y-%m-%d")
		if not env_has_name(env, "SUBMINER_APP_LOG") then
			env[#env + 1] = "SUBMINER_APP_LOG=" .. environment.join_path(log_dir, "app-" .. date .. ".log")
		end
		if not env_has_name(env, "SUBMINER_MPV_LOG") then
			env[#env + 1] = "SUBMINER_MPV_LOG=" .. environment.join_path(log_dir, "mpv-" .. date .. ".log")
		end
	end

	local function build_appimage_subprocess_env(args)
		local env = {}
		if utils and type(utils.get_env_list) == "function" then
			for _, value in ipairs(utils.get_env_list()) do
				if
					type(value) == "string"
					and not value:match("^" .. APP_ARGC_ENV .. "=")
					and not value:match("^" .. APP_ARG_PREFIX .. "%d+=")
				then
					env[#env + 1] = value
				end
			end
		end
		append_default_app_log_env(env)
		append_transport_env(env, args)
		return env
	end

	local function build_subprocess_command(args)
		if is_appimage_binary(args[1]) then
			return {
				args = { args[1] },
				env = build_appimage_subprocess_env(args),
			}
		end
		return { args = args }
	end

	run_control_command_async = function(action, overrides, callback)
		local args = build_command_args(action, overrides)
		local command = build_subprocess_command(args)
		subminer_log("debug", "process", "Control command: " .. table.concat(args, " "))
		mp.command_native_async({
			name = "subprocess",
			args = command.args,
			env = command.env,
			playback_only = false,
			capture_stdout = true,
			capture_stderr = true,
		}, function(success, result, error)
			local ok = success and (result == nil or result.status == 0)
			if ok then
				record_visible_overlay_action(action)
			end
			if callback then
				callback(ok, result, error)
			end
		end)
	end

	local function notify_playback_feedback(message, fallback)
		if type(message) ~= "string" or message == "" then
			return
		end
		if resolve_osd_messages_enabled() then
			show_osd(message)
			return
		end
		if not binary.ensure_binary_available() then
			if fallback then
				fallback()
			end
			return
		end
		run_control_command_async("playback-feedback", { message = message }, function(ok)
			if not ok and fallback then
				fallback()
			end
		end)
	end

	local function wait_for_app_ping_state(expected_running, label, on_ready, on_timeout, attempt)
		attempt = attempt or 1
		run_control_command_async("app-ping", nil, function(_ok, result)
			local status = result and result.status
			local is_running = status == 0
			local is_not_running = status == 1
			if (expected_running and is_running) or ((not expected_running) and is_not_running) then
				on_ready()
				return
			end
			if attempt >= OVERLAY_RESTART_PING_MAX_ATTEMPTS then
				subminer_log("warn", "process", "Timed out waiting for SubMiner app to " .. label)
				if on_timeout then
					on_timeout()
				end
				return
			end
			mp.add_timeout(OVERLAY_RESTART_PING_RETRY_DELAY_SECONDS, function()
				wait_for_app_ping_state(expected_running, label, on_ready, on_timeout, attempt + 1)
			end)
		end)
	end

	local function run_binary_command_async(args, callback)
		local command = build_subprocess_command(args)
		subminer_log("debug", "process", "Binary command: " .. table.concat(args, " "))
		mp.command_native_async({
			name = "subprocess",
			args = command.args,
			env = command.env,
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

		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			stop_overlay_loading_osd()
			show_osd("Error: binary not found")
			return
		end

		if state.overlay_running then
			if overrides.auto_start_trigger == true then
				subminer_log("debug", "process", "Auto-start ignored because overlay is already running")
				local socket_path = overrides.socket_path or opts.socket_path
				local should_pause_until_ready = (
					overrides.rearm_pause_until_ready == true
					and resolve_visible_overlay_startup()
					and resolve_pause_until_ready()
					and has_matching_mpv_ipc_socket(socket_path)
				)
				if should_pause_until_ready then
					arm_auto_play_ready_gate()
				elseif not state.auto_play_ready_gate_armed then
					disarm_auto_play_ready_gate()
				end
				local visibility_action = resolve_auto_start_visibility_action()
				run_visibility_action_if_needed(visibility_action, {
					socket_path = socket_path,
					log_level = overrides.log_level,
				})
				return
			end
			subminer_log("info", "process", "Overlay already running")
			show_osd("Already running")
			return
		end

		local texthooker_enabled = resolve_texthooker_enabled(overrides.texthooker_enabled)
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

			local command = build_subprocess_command(args)
			record_start_visibility_args(args)
			mp.command_native_async({
				name = "subprocess",
				args = command.args,
				env = command.env,
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
					state.auto_play_ready_signal_seen = false
					subminer_log("error", "process", "Overlay start failed after retries: " .. reason)
					stop_overlay_loading_osd()
					show_osd("Overlay start failed")
					release_auto_play_ready_gate("overlay-start-failed")
					return
				end

				if overrides.auto_start_trigger == true then
					local visibility_action = resolve_auto_start_visibility_action()
					run_visibility_action_if_needed(visibility_action, {
						socket_path = socket_path,
						log_level = overrides.log_level,
					})
				end

			end)
		end

		environment.is_subminer_app_running_async(function(app_running)
			overrides.background = not app_running
			launch_overlay_with_retry(1)
			if texthooker_enabled then
				ensure_texthooker_running(function() end)
			end
		end, { force_refresh = true })
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
		state.auto_play_ready_signal_seen = false
		stop_overlay_loading_osd()
		disarm_auto_play_ready_gate()
		show_osd("Stopped")
	end

	local function hide_visible_overlay(options)
		options = options or {}
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			return
		end
		state.suppress_ready_overlay_restore = true
		stop_overlay_loading_osd()

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

		disarm_auto_play_ready_gate({
			resume_playback = options.resume_playback ~= false,
		})
	end

	local function toggle_overlay()
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end
		if should_ignore_duplicate_visible_overlay_toggle() then
			subminer_log("debug", "process", "Ignoring duplicate visible overlay toggle")
			return
		end
		if state.visible_overlay_requested == true then
			state.suppress_ready_overlay_restore = true
			hide_visible_overlay({ resume_playback = false })
			return
		end
		if state.visible_overlay_requested == false then
			state.suppress_ready_overlay_restore = false
			disarm_auto_play_ready_gate({ resume_playback = false })
			run_control_command_async("show-visible-overlay", nil, function(ok)
				if not ok then
					subminer_log("warn", "process", "Show-visible-overlay command failed")
					show_osd("Toggle failed")
				end
			end)
			return
		end
		if not state.overlay_running then
			state.suppress_ready_overlay_restore = false
			disarm_auto_play_ready_gate({ resume_playback = false })
			start_overlay({
				show_visible_overlay = true,
			})
			return
		end
		state.suppress_ready_overlay_restore = true
		disarm_auto_play_ready_gate({ resume_playback = false })

		run_control_command_async("toggle-visible-overlay", nil, function(ok)
			if not ok then
				subminer_log("warn", "process", "Toggle command failed")
				show_osd("Toggle failed")
			end
		end)
	end

	local function toggle_primary_subtitle_bar()
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end

		run_control_command_async("toggle-primary-subtitle-bar", nil, function(ok)
			if not ok then
				subminer_log("warn", "process", "Primary subtitle bar toggle command failed")
				show_osd("Primary subtitle toggle failed")
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

		local function show_restart_feedback(message)
			notify_playback_feedback(message, function()
				show_osd(message)
			end)
		end

		start_overlay_loading_osd()
		subminer_log("info", "process", "Restarting overlay...")
		show_restart_feedback("Restarting...")

		run_control_command_async("stop", nil, function(ok, result)
			if not ok then
				local reason = result and result.stderr or "unknown error"
				subminer_log("warn", "process", "Restart stop command failed: " .. reason)
				stop_overlay_loading_osd()
				show_restart_feedback("Restart failed")
				return
			end

			state.overlay_running = false
			state.texthooker_running = false
			state.auto_play_ready_signal_seen = false
			state.suppress_ready_overlay_restore = false
			state.force_ready_overlay_restore = true
			disarm_auto_play_ready_gate({ resume_playback = false })

			wait_for_app_ping_state(false, "release the single-instance lock", function()
				local start_args = build_command_args("start", {
					show_visible_overlay = true,
				})
				subminer_log("info", "process", "Starting overlay: " .. table.concat(start_args, " "))

				state.overlay_running = true
				local command = build_subprocess_command(start_args)
				mp.command_native_async({
					name = "subprocess",
					args = command.args,
					env = command.env,
					playback_only = false,
					capture_stdout = true,
					capture_stderr = true,
				}, function(success, result, error)
					if not success or (result and result.status ~= 0) then
						state.overlay_running = false
						state.auto_play_ready_signal_seen = false
						subminer_log(
							"error",
							"process",
							"Overlay start failed: " .. (error or (result and result.stderr) or "unknown error")
						)
						stop_overlay_loading_osd()
						show_restart_feedback("Restart failed")
					else
						wait_for_app_ping_state(true, "own the single-instance lock", function()
							run_control_command_async("show-visible-overlay", nil, function()
								show_restart_feedback("Restarted successfully")
							end)
						end, function()
							run_control_command_async("show-visible-overlay", nil, function()
								show_restart_feedback("Restarted successfully")
							end)
						end)
					end
				end)

				if resolve_texthooker_enabled(nil) then
					ensure_texthooker_running(function() end)
				end
			end, function()
				stop_overlay_loading_osd()
				show_restart_feedback("Restart failed")
			end)
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
		describe_mpv_ipc_socket_match = describe_mpv_ipc_socket_match,
		has_matching_mpv_ipc_socket = has_matching_mpv_ipc_socket,
		run_control_command_async = run_control_command_async,
		notify_playback_feedback = notify_playback_feedback,
		record_visible_overlay_visibility = record_visible_overlay_visibility,
		run_binary_command_async = run_binary_command_async,
		parse_start_script_message_overrides = parse_start_script_message_overrides,
		ensure_texthooker_running = ensure_texthooker_running,
		start_overlay = start_overlay,
		start_overlay_from_script_message = start_overlay_from_script_message,
		stop_overlay = stop_overlay,
		hide_visible_overlay = hide_visible_overlay,
		toggle_overlay = toggle_overlay,
		toggle_primary_subtitle_bar = toggle_primary_subtitle_bar,
		open_options = open_options,
		restart_overlay = restart_overlay,
		check_status = check_status,
		check_binary_available = check_binary_available,
		notify_auto_play_ready = notify_auto_play_ready,
		disarm_auto_play_ready_gate = disarm_auto_play_ready_gate,
		start_overlay_loading_osd = start_overlay_loading_osd,
		stop_overlay_loading_osd = stop_overlay_loading_osd,
	}
end

return M
