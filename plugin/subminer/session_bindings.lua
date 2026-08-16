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
	MBTN_LEFT = "MBTN_LEFT",
	MBTN_MID = "MBTN_MID",
	MBTN_RIGHT = "MBTN_RIGHT",
	MBTN_BACK = "MBTN_BACK",
	MBTN_FORWARD = "MBTN_FORWARD",
}

local MODIFIER_MAP = {
	ctrl = "Ctrl",
	alt = "Alt",
	shift = "Shift",
	meta = "Meta",
}

local SHIFTED_KEY_NAME_MAP = {
	Digit1 = "!",
	Digit2 = "@",
	Digit3 = "SHARP",
	Digit4 = "$",
	Digit5 = "%",
	Digit6 = "^",
	Digit7 = "&",
	Digit8 = "*",
	Digit9 = "(",
	Digit0 = ")",
	Minus = "_",
	Equal = "+",
	BracketLeft = "{",
	BracketRight = "}",
	Backslash = "|",
	Semicolon = ":",
	Quote = '"',
	Comma = "<",
	Period = ">",
	Slash = "?",
	Backquote = "~",
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

	local function contains_value(values, target)
		for _, value in ipairs(values) do
			if value == target then
				return true
			end
		end
		return false
	end

	local function append_unique(values, value)
		if not contains_value(values, value) then
			values[#values + 1] = value
		end
	end

	local function key_spec_to_mpv_bindings(key)
		if type(key) ~= "table" then
			return nil
		end

		if type(key.code) ~= "string" then
			return nil
		end
		if type(key.modifiers) ~= "table" then
			return nil
		end

		local shifted_letter = key.code:match("^Key([A-Z])$")
		local has_shift = false
		for _, modifier in ipairs(key.modifiers) do
			if modifier == "shift" then
				has_shift = true
				break
			end
		end

		local key_name = key_code_to_mpv_name(key.code)
		if shifted_letter and has_shift then
			key_name = shifted_letter
		end
		if not key_name then
			return nil
		end

		local parts = {}
		for _, modifier in ipairs(key.modifiers) do
			if not (modifier == "shift" and shifted_letter) then
				local mapped = MODIFIER_MAP[modifier]
				if mapped then
					parts[#parts + 1] = mapped
				end
			end
		end
		parts[#parts + 1] = key_name
		local bindings = { table.concat(parts, "+") }

		local shifted_key_name = SHIFTED_KEY_NAME_MAP[key.code]
		if has_shift and shifted_key_name then
			local shifted_parts = {}
			for _, modifier in ipairs(key.modifiers) do
				if modifier ~= "shift" then
					local mapped = MODIFIER_MAP[modifier]
					if mapped then
						shifted_parts[#shifted_parts + 1] = mapped
					end
				end
			end
			shifted_parts[#shifted_parts + 1] = shifted_key_name
			append_unique(bindings, table.concat(shifted_parts, "+"))
		end

		return bindings
	end

	local function normalize_cli_args(cli_args)
		if type(cli_args) ~= "table" then
			return nil
		end

		local normalized = {}
		for _, arg in ipairs(cli_args) do
			if type(arg) ~= "string" and type(arg) ~= "number" then
				return nil
			end
			local value = tostring(arg)
			if value == "" then
				return nil
			end
			normalized[#normalized + 1] = value
		end

		if #normalized == 0 then
			return nil
		end
		return normalized
	end

	local function build_cli_args(action_id, payload, artifact_cli_args)
		local cli_args = normalize_cli_args(artifact_cli_args)
		if cli_args then
			return cli_args
		end

		if action_id == "toggleVisibleOverlay" then
			return { "--toggle-visible-overlay" }
		elseif action_id == "toggleStatsOverlay" then
			return { "--toggle-stats-overlay" }
		elseif action_id == "copySubtitle" then
			return { "--copy-subtitle" }
		elseif action_id == "copySubtitleMultiple" then
			if payload and payload.count then
				return { "--copy-subtitle-count", tostring(payload.count) }
			end
			return { "--copy-subtitle-multiple" }
		elseif action_id == "updateLastCardFromClipboard" then
			return { "--update-last-card-from-clipboard" }
		elseif action_id == "triggerFieldGrouping" then
			return { "--trigger-field-grouping" }
		elseif action_id == "triggerSubsync" then
			return { "--session-action", '{"actionId":"triggerSubsync"}' }
		elseif action_id == "mineSentence" then
			return { "--mine-sentence" }
		elseif action_id == "mineSentenceMultiple" then
			if payload and payload.count then
				return { "--mine-sentence-count", tostring(payload.count) }
			end
			return { "--mine-sentence-multiple" }
		elseif action_id == "toggleSecondarySub" then
			return { "--toggle-secondary-sub" }
		elseif action_id == "toggleSubtitleSidebar" then
			return { "--toggle-subtitle-sidebar" }
		elseif action_id == "toggleNotificationHistory" then
			return { "--session-action", '{"actionId":"toggleNotificationHistory"}' }
		elseif action_id == "markAudioCard" then
			return { "--mark-audio-card" }
		elseif action_id == "markWatched" then
			return { "--mark-watched" }
		elseif action_id == "openRuntimeOptions" then
			return { "--session-action", '{"actionId":"openRuntimeOptions"}' }
		elseif action_id == "openJimaku" then
			return { "--open-jimaku" }
		elseif action_id == "openTsukihime" or action_id == "openAnimetosho" then
			return { "--open-tsukihime" }
		elseif action_id == "openYoutubePicker" then
			return { "--open-youtube-picker" }
		elseif action_id == "openSessionHelp" then
			return { "--session-action", '{"actionId":"openSessionHelp"}' }
		elseif action_id == "openCharacterDictionaryManager" then
			return { "--session-action", '{"actionId":"openCharacterDictionaryManager"}' }
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

	local function invoke_cli_action(action_id, payload, artifact_cli_args)
		if not process.check_binary_available() then
			show_osd("Error: binary not found")
			return
		end

		local cli_args = build_cli_args(action_id, payload, artifact_cli_args)
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

	local function execute_mpv_command(command)
		if type(command) ~= "table" or command[1] == nil then
			return
		end
		mp.commandv(unpack_fn(command))
	end

	local function handle_binding(binding)
		if binding.actionType == "mpv-command" then
			execute_mpv_command(binding.command)
			return
		end

		if binding.actionId == "toggleVisibleOverlay" and type(process.toggle_overlay) == "function" then
			process.toggle_overlay()
			return
		end

		invoke_cli_action(binding.actionId, binding.payload, binding.cliArgs)
	end

	local function is_supported_binding(binding)
		if binding.actionType == "mpv-command" then
			return type(binding.command) == "table" and binding.command[1] ~= nil
		end
		if binding.actionType == "session-action" then
			return build_cli_args(binding.actionId, binding.payload, binding.cliArgs) ~= nil
		end
		return false
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
		remove_binding_names(state.session_binding_names)
	end

	local function register_bindings()
		local artifact, load_error = load_artifact()
		if not artifact then
			subminer_log("warn", "session-bindings", load_error)
			return false
		end

		local previous_binding_names = state.session_binding_names
		local next_binding_names = {}
		state.session_binding_generation = (state.session_binding_generation or 0) + 1
		local generation = state.session_binding_generation

		for index, binding in ipairs(artifact.bindings) do
			if not is_supported_binding(binding) then
				subminer_log(
					"warn",
					"session-bindings",
					"Skipped unsupported session binding from artifact"
				)
			else
				local key_names = key_spec_to_mpv_bindings(binding.key)
				if key_names then
					for key_index, key_name in ipairs(key_names) do
						local name = "subminer-session-binding-"
							.. tostring(generation)
							.. "-"
							.. tostring(index)
							.. "-"
							.. tostring(key_index)
						next_binding_names[#next_binding_names + 1] = name
						mp.add_forced_key_binding(key_name, name, function()
							handle_binding(binding)
						end)
					end
				else
					subminer_log(
						"warn",
						"session-bindings",
						"Skipped unsupported key code from artifact: " .. tostring(binding.key and binding.key.code or "unknown")
					)
				end
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
