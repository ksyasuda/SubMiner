local M = {}

local OVERLAY_START_RETRY_DELAY_SECONDS = 0.2
local OVERLAY_START_MAX_ATTEMPTS = 6
local STARTUP_OVERLAY_ACTION_RETRY_DELAY_SECONDS = 0.2
local STARTUP_OVERLAY_ACTION_MAX_ATTEMPTS = 6

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

		if action == "start" then
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

	local function run_control_command_async(action, overrides, callback)
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

	local function resolve_visible_overlay_startup()
		return options_helper.coerce_bool(opts.auto_start_visible_overlay, false)
	end

	local function apply_startup_overlay_preferences()
		local should_show_visible = resolve_visible_overlay_startup()
		local visible_action = should_show_visible and "show-visible-overlay" or "hide-visible-overlay"

		local function try_apply(attempt)
			run_control_command_async(visible_action, nil, function(ok)
				if ok then
					subminer_log(
						"debug",
						"process",
						"Applied visible startup action: " .. visible_action .. " (attempt " .. tostring(attempt) .. ")"
					)
					return
				end

				if attempt >= STARTUP_OVERLAY_ACTION_MAX_ATTEMPTS then
					subminer_log("warn", "process", "Failed to apply visible startup action: " .. visible_action)
					return
				end

				mp.add_timeout(STARTUP_OVERLAY_ACTION_RETRY_DELAY_SECONDS, function()
					try_apply(attempt + 1)
				end)
			end)
		end

		try_apply(1)
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

		-- Start overlay immediately; overlay start path retries on readiness failures.
		callback()
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

			if attempt == 1 then
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
					return
				end

				apply_startup_overlay_preferences()
			end)
		end

		if texthooker_enabled then
			ensure_texthooker_running(function()
				launch_overlay_with_retry(1)
			end)
		else
			launch_overlay_with_retry(1)
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
		show_osd("Stopped")
	end

	local function toggle_overlay()
		if not binary.ensure_binary_available() then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end

		run_control_command_async("toggle", nil, function(ok)
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
		run_control_command_async = run_control_command_async,
		parse_start_script_message_overrides = parse_start_script_message_overrides,
		apply_startup_overlay_preferences = apply_startup_overlay_preferences,
		ensure_texthooker_running = ensure_texthooker_running,
		start_overlay = start_overlay,
		start_overlay_from_script_message = start_overlay_from_script_message,
		stop_overlay = stop_overlay,
		toggle_overlay = toggle_overlay,
		open_options = open_options,
		restart_overlay = restart_overlay,
		check_status = check_status,
		check_binary_available = check_binary_available,
	}
end

return M
