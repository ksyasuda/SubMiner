local M = {}

local unpack_fn = table.unpack or unpack

local KEY_NAME_MAP = {
	Space = "SPACE",
	Tab = "TAB",
	Enter = "ENTER",
	Escape = "ESC",
	Backspace = "BS",
	Delete = "DEL",
	ArrowUp = "UP",
	ArrowDown = "DOWN",
	ArrowLeft = "LEFT",
	ArrowRight = "RIGHT",
	Slash = "/",
	Backslash = "\\",
	Minus = "-",
	Equal = "=",
	Comma = ",",
	Period = ".",
	Quote = "'",
	Semicolon = ";",
	BracketLeft = "[",
	BracketRight = "]",
	Backquote = "`",
}

local MODIFIER_MAP = {
	ctrl = "Ctrl",
	alt = "Alt",
	shift = "Shift",
	meta = "Meta",
}

function M.create(ctx)
	local mp = ctx.mp
	local utils = ctx.utils
	local state = ctx.state
	local process = ctx.process
	local environment = ctx.environment
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd

	local function read_file(path)
		local handle = io.open(path, "r")
		if not handle then
			return nil
		end
		local content = handle:read("*a")
		handle:close()
		return content
	end

	local function remove_binding_names(names)
		for _, name in ipairs(names) do
			mp.remove_key_binding(name)
		end
		for index = #names, 1, -1 do
			names[index] = nil
		end
	end

	local function key_code_to_mpv_name(code)
		if KEY_NAME_MAP[code] then
			return KEY_NAME_MAP[code]
		end

		local letter = code:match("^Key([A-Z])$")
		if letter then
			return string.lower(letter)
		end

		local digit = code:match("^Digit([0-9])$")
		if digit then
			return digit
		end

		local function_key = code:match("^(F%d+)$")
		if function_key then
			return function_key
		end

		return nil
	end

	local function key_spec_to_mpv_binding(key)
		if type(key) ~= "table" then
			return nil
		end

		if type(key.code) ~= "string" then
			return nil
		end
		if type(key.modifiers) ~= "table" then
			return nil
		end

		local key_name = key_code_to_mpv_name(key.code)
		if not key_name then
			return nil
		end

		local parts = {}
		for _, modifier in ipairs(key.modifiers) do
			local mapped = MODIFIER_MAP[modifier]
			if mapped then
				parts[#parts + 1] = mapped
			end
		end
		parts[#parts + 1] = key_name
		return table.concat(parts, "+")
	end

	local function build_cli_args(action_id, payload)
		if action_id == "toggleVisibleOverlay" then
			return { "--toggle-visible-overlay" }
		elseif action_id == "toggleStatsOverlay" then
			return { "--toggle-stats-overlay" }
		elseif action_id == "copySubtitle" then
			return { "--copy-subtitle" }
		elseif action_id == "copySubtitleMultiple" then
			return { "--copy-subtitle-count", tostring(payload and payload.count or 1) }
		elseif action_id == "updateLastCardFromClipboard" then
			return { "--update-last-card-from-clipboard" }
		elseif action_id == "triggerFieldGrouping" then
			return { "--trigger-field-grouping" }
		elseif action_id == "triggerSubsync" then
			return { "--trigger-subsync" }
		elseif action_id == "mineSentence" then
			return { "--mine-sentence" }
		elseif action_id == "mineSentenceMultiple" then
			return { "--mine-sentence-count", tostring(payload and payload.count or 1) }
		elseif action_id == "toggleSecondarySub" then
			return { "--toggle-secondary-sub" }
		elseif action_id == "toggleSubtitleSidebar" then
			return { "--toggle-subtitle-sidebar" }
		elseif action_id == "markAudioCard" then
			return { "--mark-audio-card" }
		elseif action_id == "openRuntimeOptions" then
			return { "--open-runtime-options" }
		elseif action_id == "openJimaku" then
			return { "--open-jimaku" }
		elseif action_id == "openYoutubePicker" then
			return { "--open-youtube-picker" }
		elseif action_id == "openSessionHelp" then
			return { "--open-session-help" }
		elseif action_id == "openControllerSelect" then
			return { "--open-controller-select" }
		elseif action_id == "openControllerDebug" then
			return { "--open-controller-debug" }
		elseif action_id == "openPlaylistBrowser" then
			return { "--open-playlist-browser" }
		elseif action_id == "replayCurrentSubtitle" then
			return { "--replay-current-subtitle" }
		elseif action_id == "playNextSubtitle" then
			return { "--play-next-subtitle" }
		elseif action_id == "shiftSubDelayPrevLine" then
			return { "--shift-sub-delay-prev-line" }
		elseif action_id == "shiftSubDelayNextLine" then
			return { "--shift-sub-delay-next-line" }
		elseif action_id == "cycleRuntimeOption" then
			local runtime_option_id = payload and payload.runtimeOptionId or nil
			if type(runtime_option_id) ~= "string" or runtime_option_id == "" then
				return nil
			end
			local direction = payload and payload.direction == -1 and "prev" or "next"
			return { "--cycle-runtime-option", runtime_option_id .. ":" .. direction }
		end

		return nil
	end

	local function invoke_cli_action(action_id, payload)
		if not process.check_binary_available() then
			show_osd("Error: binary not found")
			return
		end

		local cli_args = build_cli_args(action_id, payload)
		if not cli_args then
			subminer_log("warn", "session-bindings", "No CLI mapping for action: " .. tostring(action_id))
			return
		end

		local args = { state.binary_path }
		for _, arg in ipairs(cli_args) do
			args[#args + 1] = arg
		end
		local runner = process.run_binary_command_async
		if type(runner) ~= "function" then
			runner = function(binary_args, callback)
				mp.command_native_async({
					name = "subprocess",
					args = binary_args,
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
		end
		runner(args, function(ok, result, error)
			if ok then
				return
			end
			local reason = error or (result and result.stderr) or "unknown error"
			subminer_log("warn", "session-bindings", "Session action failed: " .. tostring(reason))
			show_osd("Session action failed")
		end)
	end

	local function clear_numeric_selection(show_cancelled)
		if state.session_numeric_selection and state.session_numeric_selection.timeout then
			state.session_numeric_selection.timeout:kill()
		end
		state.session_numeric_selection = nil
		remove_binding_names(state.session_numeric_binding_names)
		if show_cancelled then
			show_osd("Cancelled")
		end
	end

	local function start_numeric_selection(action_id, timeout_ms)
		clear_numeric_selection(false)
		for digit = 1, 9 do
			local digit_string = tostring(digit)
			local name = "subminer-session-digit-" .. digit_string
			state.session_numeric_binding_names[#state.session_numeric_binding_names + 1] = name
			mp.add_forced_key_binding(digit_string, name, function()
				clear_numeric_selection(false)
				invoke_cli_action(action_id, { count = digit })
			end)
		end

		state.session_numeric_binding_names[#state.session_numeric_binding_names + 1] =
			"subminer-session-digit-cancel"
		mp.add_forced_key_binding("ESC", "subminer-session-digit-cancel", function()
			clear_numeric_selection(true)
		end)

		state.session_numeric_selection = {
			action_id = action_id,
			timeout = mp.add_timeout((timeout_ms or 3000) / 1000, function()
				clear_numeric_selection(false)
				show_osd(action_id == "copySubtitleMultiple" and "Copy timeout" or "Mine timeout")
			end),
		}

		show_osd(
			action_id == "copySubtitleMultiple"
					and "Copy how many lines? Press 1-9 (Esc to cancel)"
				or "Mine how many lines? Press 1-9 (Esc to cancel)"
		)
	end

	local function execute_mpv_command(command)
		if type(command) ~= "table" or command[1] == nil then
			return
		end
		mp.commandv(unpack_fn(command))
	end

	local function handle_binding(binding, numeric_selection_timeout_ms)
		if binding.actionType == "mpv-command" then
			execute_mpv_command(binding.command)
			return
		end

		if binding.actionId == "copySubtitleMultiple" or binding.actionId == "mineSentenceMultiple" then
			start_numeric_selection(binding.actionId, numeric_selection_timeout_ms)
			return
		end

		invoke_cli_action(binding.actionId, binding.payload)
	end

	local function load_artifact()
		local artifact_path = environment.resolve_session_bindings_artifact_path()
		local raw = read_file(artifact_path)
		if not raw or raw == "" then
			return nil, "Missing session binding artifact: " .. tostring(artifact_path)
		end

		local parsed, parse_error = utils.parse_json(raw)
		if not parsed then
			return nil, "Failed to parse session binding artifact: " .. tostring(parse_error)
		end
		if type(parsed) ~= "table" or type(parsed.bindings) ~= "table" then
			return nil, "Invalid session binding artifact"
		end

		return parsed, nil
	end

	local function clear_bindings()
		clear_numeric_selection(false)
		remove_binding_names(state.session_binding_names)
	end

	local function register_bindings()
		local artifact, load_error = load_artifact()
		if not artifact then
			subminer_log("warn", "session-bindings", load_error)
			return false
		end

		clear_numeric_selection(false)

		local previous_binding_names = state.session_binding_names
		local next_binding_names = {}
		state.session_binding_generation = (state.session_binding_generation or 0) + 1
		local generation = state.session_binding_generation

		local timeout_ms = tonumber(artifact.numericSelectionTimeoutMs) or 3000
		for index, binding in ipairs(artifact.bindings) do
			local key_name = key_spec_to_mpv_binding(binding.key)
			if key_name then
				local name = "subminer-session-binding-" .. tostring(generation) .. "-" .. tostring(index)
				next_binding_names[#next_binding_names + 1] = name
				mp.add_forced_key_binding(key_name, name, function()
					handle_binding(binding, timeout_ms)
				end)
			else
				subminer_log(
					"warn",
					"session-bindings",
					"Skipped unsupported key code from artifact: " .. tostring(binding.key and binding.key.code or "unknown")
				)
			end
		end

		remove_binding_names(previous_binding_names)
		state.session_binding_names = next_binding_names

		subminer_log(
			"info",
			"session-bindings",
			"Registered " .. tostring(#next_binding_names) .. " shared session bindings"
		)
		return true
	end

	local function reload_bindings()
		return register_bindings()
	end

	return {
		register_bindings = register_bindings,
		reload_bindings = reload_bindings,
		clear_bindings = clear_bindings,
	}
end

return M
