local function run_plugin_scenario(config)
	config = config or {}

	local recorded = {
		async_calls = {},
		sync_calls = {},
		script_messages = {},
		events = {},
		observers = {},
		key_bindings = {},
		osd = {},
		logs = {},
		property_sets = {},
		periodic_timers = {},
		timeouts = {},
	}

	local function make_mp_stub()
		local mp = {}

		function mp.get_property(name)
			if name == "platform" then
				return config.platform or "linux"
			end
			if name == "input-ipc-server" then
				if config.input_ipc_server_sequence then
					config.input_ipc_server_sequence_index = (config.input_ipc_server_sequence_index or 0) + 1
					local index = config.input_ipc_server_sequence_index
					return config.input_ipc_server_sequence[index] or config.input_ipc_server_sequence[#config.input_ipc_server_sequence] or ""
				end
				return config.input_ipc_server or ""
			end
			if name == "filename/no-ext" then
				return config.filename_no_ext or ""
			end
			if name == "filename" then
				return config.filename or ""
			end
			if name == "path" then
				return config.path or ""
			end
			if name == "media-title" then
				return config.media_title or ""
			end
			return ""
		end

		function mp.get_property_native(name)
			if name == "pause" then
				return config.paused == true
			end
			if name == "osd-dimensions" then
				return config.osd_dimensions or {
					w = 1280,
					h = config.osd_height or 720,
				}
			end
			return config.chapter_list or {}
		end

		function mp.get_property_number(name)
			if name == "time-pos" then
				return config.time_pos
			end
			if name == "sub-pos" then
				return config.sub_pos or 100
			end
			if name == "osd-height" then
				return config.osd_height or 720
			end
			return nil
		end

		function mp.get_script_directory()
			return "plugin/subminer"
		end

		function mp.command_native(command)
			recorded.sync_calls[#recorded.sync_calls + 1] = command
			local args = command.args or {}
			if args[1] == "ps" then
				return {
					status = 0,
					stdout = config.process_list or "",
					stderr = "",
				}
			end
			if args[1] == "curl" then
				local url = args[#args] or ""
				if type(url) == "string" and url:find("myanimelist", 1, true) then
					return { status = 0, stdout = config.mal_lookup_stdout or "{}", stderr = "" }
				end
				if type(url) == "string" and url:find("api.aniskip.com", 1, true) then
					return { status = 0, stdout = config.aniskip_stdout or "{}", stderr = "" }
				end
				return { status = 0, stdout = "{}", stderr = "" }
			end
			return { status = 0, stdout = "", stderr = "" }
		end

		function mp.command_native_async(command, callback)
			recorded.async_calls[#recorded.async_calls + 1] = command
			if callback then
				local args = command.args or {}
				if args[1] == "ps" then
					callback(true, { status = 0, stdout = config.process_list or "", stderr = "" }, nil)
					return
				end
				if args[1] == "curl" then
					local url = args[#args] or ""
					if type(url) == "string" and url:find("myanimelist", 1, true) then
						callback(true, { status = 0, stdout = config.mal_lookup_stdout or "{}", stderr = "" }, nil)
						return
					end
					if type(url) == "string" and url:find("api.aniskip.com", 1, true) then
						callback(true, { status = 0, stdout = config.aniskip_stdout or "{}", stderr = "" }, nil)
						return
					end
				end
				for _, value in ipairs(args) do
					if value == "--app-ping" then
						config.app_ping_index = (config.app_ping_index or 0) + 1
						local statuses = config.app_ping_statuses or { 1 }
						local status = statuses[config.app_ping_index] or statuses[#statuses] or 1
						callback(status == 0, { status = status, stdout = "", stderr = "" }, nil)
						return
					end
					if value == "--stop" and config.stop_command_fails then
						local stderr = config.stop_command_stderr or "stop failed"
						callback(false, { status = 1, stdout = "", stderr = stderr }, stderr)
						return
					end
				end
				callback(true, { status = 0, stdout = "", stderr = "" }, nil)
			end
		end

		function mp.add_timeout(seconds, callback)
			recorded.timeouts[#recorded.timeouts + 1] = seconds
			local timeout = {
				killed = false,
			}
			function timeout:kill()
				self.killed = true
			end

			local delay = tonumber(seconds) or 0
			if callback and delay < 5 then
				callback()
			end
			return timeout
		end

		function mp.add_periodic_timer(seconds, callback)
			local timer = {
				seconds = seconds,
				killed = false,
				callback = callback,
			}
			function timer:kill()
				self.killed = true
			end
			recorded.periodic_timers[#recorded.periodic_timers + 1] = timer
			return timer
		end

		function mp.register_script_message(name, fn)
			recorded.script_messages[name] = fn
		end

		function mp.add_key_binding(keys, name, fn)
			recorded.key_bindings[#recorded.key_bindings + 1] = {
				keys = keys,
				name = name,
				fn = fn,
			}
		end
		function mp.add_forced_key_binding(keys, name, fn)
			recorded.key_bindings[#recorded.key_bindings + 1] = {
				keys = keys,
				name = name,
				fn = fn,
				forced = true,
			}
		end
		function mp.register_event(name, fn)
			if recorded.events[name] == nil then
				recorded.events[name] = {}
			end
			recorded.events[name][#recorded.events[name] + 1] = fn
		end
		function mp.add_hook(_name, _prio, _fn) end
		function mp.observe_property(name, _kind, fn)
			if recorded.observers[name] == nil then
				recorded.observers[name] = {}
			end
			recorded.observers[name][#recorded.observers[name] + 1] = fn
		end
		function mp.osd_message(message, _duration)
			recorded.osd[#recorded.osd + 1] = message
		end
		function mp.set_osd_ass(...) end
		function mp.get_time()
			return 0
		end
		function mp.commandv(...) end
		function mp.set_property_native(name, value)
			recorded.property_sets[#recorded.property_sets + 1] = {
				name = name,
				value = value,
			}
			if name == "pause" then
				config.paused = value == true
			end
		end
		function mp.set_property(name, value)
			recorded.property_sets[#recorded.property_sets + 1] = {
				name = name,
				value = value,
			}
		end
		function mp.get_script_name()
			return "subminer"
		end

		return mp
	end

	local mp = make_mp_stub()
	local options = {}
	local utils = {}

	function options.read_options(target, _name)
		for key, value in pairs(config.option_overrides or {}) do
			target[key] = value
		end
	end

	function utils.file_info(path)
		local exists = config.files and config.files[path]
		if exists then
			return { is_dir = false }
		end
		return nil
	end

	function utils.join_path(...)
		local parts = { ... }
		return table.concat(parts, "/")
	end

	function utils.get_env_list()
		return config.env_list or {}
	end

	function utils.parse_json(json)
		if json == '{"enabled":true,"amount":125}' then
			return {
				enabled = true,
				amount = 125,
			}, nil
		end
		if json == "__MAL_FOUND__" then
			return {
				categories = {
					{
						items = {
							{
								id = 99,
								name = "Sample Show",
							},
						},
					},
				},
			}, nil
		end
		if json == "__ANISKIP_FOUND__" then
			return {
				found = true,
				results = {
					{
						skip_type = "op",
						interval = {
							start_time = 12.3,
							end_time = 45.6,
						},
					},
				},
			}, nil
		end
		return {}, nil
	end

	package.loaded["mp"] = nil
	package.loaded["mp.input"] = nil
	package.loaded["mp.msg"] = nil
	package.loaded["mp.options"] = nil
	package.loaded["mp.utils"] = nil
	package.loaded["binary"] = nil
	package.loaded["bootstrap"] = nil
	package.loaded["environment"] = nil
	package.loaded["hover"] = nil
	package.loaded["init"] = nil
	package.loaded["lifecycle"] = nil
	package.loaded["log"] = nil
	package.loaded["messages"] = nil
	package.loaded["options"] = nil
	package.loaded["process"] = nil
	package.loaded["state"] = nil
	package.loaded["ui"] = nil
	package.loaded["aniskip"] = nil
	_G.__subminer_plugin_bootstrapped = nil
	local original_package_config = package.config
	if config.platform == "windows" then
		package.config = "\\\n;\n?\n!\n-\n"
	else
		package.config = "/\n;\n?\n!\n-\n"
	end

	package.preload["mp"] = function()
		return mp
	end
	package.preload["mp.input"] = function()
		return {
			select = function(_) end,
		}
	end
	package.preload["mp.msg"] = function()
		return {
			info = function(line)
				recorded.logs[#recorded.logs + 1] = line
			end,
			warn = function(line)
				recorded.logs[#recorded.logs + 1] = line
			end,
			error = function(line)
				recorded.logs[#recorded.logs + 1] = line
			end,
			debug = function(line)
				recorded.logs[#recorded.logs + 1] = line
			end,
		}
	end
	package.preload["mp.options"] = function()
		return options
	end
	package.preload["mp.utils"] = function()
		return utils
	end

	local ok, err = pcall(dofile, "plugin/subminer/main.lua")
	package.config = original_package_config
	if not ok then
		return nil, err, recorded
	end
	return recorded, nil, recorded
end

local function assert_true(condition, message)
	if condition then
		return
	end
	error(message)
end

local function find_start_call(async_calls)
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		for i = 1, #args do
			if args[i] == "--start" then
				return call
			end
		end
	end
	return nil
end

local function count_start_calls(async_calls)
	local count = 0
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		for _, value in ipairs(args) do
			if value == "--start" then
				count = count + 1
				break
			end
		end
	end
	return count
end

local function find_texthooker_call(async_calls)
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		for i = 1, #args do
			if args[i] == "--texthooker" then
				return call
			end
		end
	end
	return nil
end

local function find_call_index(async_calls, target_call)
	for index, call in ipairs(async_calls) do
		if call == target_call then
			return index
		end
	end
	return nil
end

local function find_control_call(async_calls, flag)
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		local has_flag = false
		local has_start = false
		for _, value in ipairs(args) do
			if value == flag then
				has_flag = true
			elseif value == "--start" then
				has_start = true
			end
		end
		if has_flag and not has_start then
			return call
		end
	end
	return nil
end

local function find_nth_control_call(async_calls, flag, target_count)
	local count = 0
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		local has_flag = false
		local has_start = false
		for _, value in ipairs(args) do
			if value == flag then
				has_flag = true
			elseif value == "--start" then
				has_start = true
			end
		end
		if has_flag and not has_start then
			count = count + 1
			if count == target_count then
				return call
			end
		end
	end
	return nil
end

local function count_control_calls(async_calls, flag)
	local count = 0
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		local has_flag = false
		local has_start = false
		for _, value in ipairs(args) do
			if value == flag then
				has_flag = true
			elseif value == "--start" then
				has_start = true
			end
		end
		if has_flag and not has_start then
			count = count + 1
		end
	end
	return count
end

local function call_has_arg(call, target)
	local args = (call and call.args) or {}
	for _, value in ipairs(args) do
		if value == target then
			return true
		end
	end
	return false
end

local function has_sync_command(sync_calls, executable)
	for _, call in ipairs(sync_calls) do
		local args = call.args or {}
		if args[1] == executable then
			return true
		end
	end
	return false
end

local function has_async_command(async_calls, executable)
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		if args[1] == executable then
			return true
		end
	end
	return false
end

local function has_async_curl_for(async_calls, needle)
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		if args[1] == "curl" then
			local url = args[#args] or ""
			if type(url) == "string" and url:find(needle, 1, true) then
				return true
			end
		end
	end
	return false
end

local function count_async_curl_for(async_calls, needle)
	local count = 0
	for _, call in ipairs(async_calls) do
		local args = call.args or {}
		if args[1] == "curl" then
			local url = args[#args] or ""
			if type(url) == "string" and url:find(needle, 1, true) then
				count = count + 1
			end
		end
	end
	return count
end

local function has_property_set(property_sets, name, value)
	for _, call in ipairs(property_sets) do
		if call.name == name and call.value == value then
			return true
		end
	end
	return false
end

local function has_osd_message(messages, target)
	for _, message in ipairs(messages) do
		if message == target then
			return true
		end
	end
	return false
end

local function count_osd_message(messages, target)
	local count = 0
	for _, message in ipairs(messages) do
		if message == target then
			count = count + 1
		end
	end
	return count
end

local function has_timeout(timeouts, target)
	for _, seconds in ipairs(timeouts) do
		if math.abs(seconds - target) < 0.0001 then
			return true
		end
	end
	return false
end

local function env_has(call, target)
	local env = (call and call.env) or {}
	for _, value in ipairs(env) do
		if value == target then
			return true
		end
	end
	return false
end

local function env_has_prefix(call, target)
	local env = (call and call.env) or {}
	for _, value in ipairs(env) do
		if type(value) == "string" and value:sub(1, #target) == target then
			return true
		end
	end
	return false
end

local function count_property_set(property_sets, name, value)
	local count = 0
	for _, call in ipairs(property_sets) do
		if call.name == name and call.value == value then
			count = count + 1
		end
	end
	return count
end

local function fire_event(recorded, name, ...)
	local listeners = recorded.events[name] or {}
	for _, listener in ipairs(listeners) do
		listener(...)
	end
end

local function fire_observer(recorded, name, value)
	local listeners = recorded.observers[name] or {}
	for _, listener in ipairs(listeners) do
		listener(name, value)
	end
end

local function has_key_binding(recorded, keys, name)
	for _, binding in ipairs(recorded.key_bindings or {}) do
		if binding.keys == keys and binding.name == name then
			return true
		end
	end
	return false
end

local binary_path = "/tmp/subminer-binary"
local appimage_path = "/tmp/SubMiner.AppImage"

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
		},
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for cold-start scenario: " .. tostring(err))
	assert_true(recorded.script_messages["subminer-start"] ~= nil, "subminer-start script message not registered")
	recorded.script_messages["subminer-start"]("texthooker=no")
	assert_true(find_start_call(recorded.async_calls) ~= nil, "expected cold-start to invoke --start command when process is absent")
	assert_true(
		not has_sync_command(recorded.sync_calls, "ps"),
		"expected cold-start start command to avoid synchronous process list scan"
	)
end

do
	local scenario = {
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		path = "/media/episode-01.mkv",
		media_title = "Episode 1",
		files = {
			[binary_path] = true,
		},
	}
	local recorded, err = run_plugin_scenario(scenario)
	assert_true(recorded ~= nil, "plugin failed to load for new-media rearm scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	recorded.script_messages["subminer-autoplay-ready"]()
	fire_event(recorded, "end-file", { reason = "eof" })
	scenario.path = "/media/episode-02.mkv"
	scenario.media_title = "Episode 2"
	fire_event(recorded, "file-loaded")
	assert_true(
		count_start_calls(recorded.async_calls) == 1,
		"new media after prior playback should reuse the running overlay"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", true) == 2,
		"new media after prior playback should re-arm pause-until-ready"
	)
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 2,
		"new media after prior playback should resume only after readiness"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server_sequence = { "", "", "/tmp/subminer-socket" },
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for delayed socket auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(find_start_call(recorded.async_calls) ~= nil, "delayed socket auto-start should eventually issue --start")
	assert_true(
		has_property_set(recorded.property_sets, "pause", true),
		"delayed socket auto-start should arm pause-until-ready once the socket is available"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		platform = "osx",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for macOS platform alias scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(start_call ~= nil, "macOS platform alias auto-start should issue --start")
	assert_true(
		call_has_arg(start_call, "macos"),
		"macOS platform alias auto-start should pass macos backend instead of falling back to x11"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = appimage_path,
			auto_start = "no",
			socket_path = "/tmp/subminer-socket",
		},
		files = {
			[appimage_path] = true,
		},
		env_list = {
			"PATH=/usr/bin",
			"SUBMINER_APP_ARGC=stale",
			"SUBMINER_APP_ARG_0=--stale",
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for AppImage env transport scenario: " .. tostring(err))
	recorded.script_messages["subminer-start"]("texthooker=no")
	local call = recorded.async_calls[#recorded.async_calls]
	assert_true(call ~= nil, "AppImage start should issue an async subprocess")
	assert_true(#call.args == 1 and call.args[1] == appimage_path, "AppImage subprocess should not receive raw CLI flags")
	assert_true(env_has(call, "PATH=/usr/bin"), "AppImage subprocess should preserve existing environment")
	assert_true(env_has(call, "SUBMINER_APP_ARGC=8"), "AppImage subprocess should transport app arg count")
	assert_true(env_has(call, "SUBMINER_APP_ARG_0=--start"), "AppImage subprocess should transport --start")
	assert_true(
		env_has(call, "SUBMINER_APP_ARG_1=--background"),
		"AppImage subprocess should transport --background"
	)
	assert_true(
		env_has(call, "SUBMINER_APP_ARG_2=--managed-playback"),
		"AppImage subprocess should transport --managed-playback"
	)
	assert_true(env_has(call, "SUBMINER_APP_ARG_7=--hide-visible-overlay"), "AppImage subprocess should transport visibility flag")
	assert_true(env_has_prefix(call, "SUBMINER_APP_LOG="), "AppImage subprocess should include app log env")
	assert_true(env_has_prefix(call, "SUBMINER_MPV_LOG="), "AppImage subprocess should include mpv log env")
	assert_true(
		not env_has(call, "SUBMINER_APP_ARG_0=--stale"),
		"AppImage subprocess should remove stale transported args"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		app_ping_statuses = { 0, 1, 0 },
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			auto_start_visible_overlay = "no",
		},
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for manual visible restart scenario: " .. tostring(err))
	local restart_binding = nil
	for _, candidate in ipairs(recorded.key_bindings) do
		if candidate.name == "subminer-restart" then
			restart_binding = candidate
			break
		end
	end
	assert_true(restart_binding ~= nil, "restart binding should be registered")
	restart_binding.fn()
	local start_call = find_start_call(recorded.async_calls)
	assert_true(start_call ~= nil, "manual restart should issue --start command")
	local start_index = find_call_index(recorded.async_calls, start_call) or 0
	local old_app_ping = find_nth_control_call(recorded.async_calls, "--app-ping", 1)
	local old_app_stopped_ping = find_nth_control_call(recorded.async_calls, "--app-ping", 2)
	local new_app_started_ping = find_nth_control_call(recorded.async_calls, "--app-ping", 3)
	assert_true(old_app_ping ~= nil, "manual restart should ping before waiting for old app shutdown")
	assert_true(old_app_stopped_ping ~= nil, "manual restart should keep pinging until old app shutdown")
	assert_true(new_app_started_ping ~= nil, "manual restart should ping after start until the new app is running")
	assert_true(
		(find_call_index(recorded.async_calls, old_app_ping) or 0) < start_index,
		"manual restart should wait for old app ping before starting"
	)
	assert_true(
		(find_call_index(recorded.async_calls, old_app_stopped_ping) or 0) < start_index,
		"manual restart should wait for old app stopped ping before starting"
	)
	assert_true(
		start_index < (find_call_index(recorded.async_calls, new_app_started_ping) or 0),
		"manual restart should wait for new app running ping after starting"
	)
	assert_true(
		call_has_arg(start_call, "--show-visible-overlay"),
		"manual restart should bring the visible overlay back after process reload"
	)
	assert_true(
		not call_has_arg(start_call, "--hide-visible-overlay"),
		"manual restart should not restart into hidden visible-overlay state"
	)
	assert_true(
		not has_timeout(recorded.timeouts, 0.5),
		"manual restart should use app-ping readiness instead of a fixed 0.5s start delay"
	)
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 1,
		"manual restart should re-assert visible overlay after the restarted app is launched"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		app_ping_statuses = { 0, 2, 1, 0 },
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			auto_start_visible_overlay = "no",
		},
		files = {
			[binary_path] = true,
		},
	})
	assert_true(
		recorded ~= nil,
		"plugin failed to load for transient app-ping failure restart scenario: " .. tostring(err)
	)
	recorded.script_messages["subminer-restart"]()
	local start_call = find_start_call(recorded.async_calls)
	assert_true(start_call ~= nil, "manual restart should start after app-ping reports stopped")
	local start_index = find_call_index(recorded.async_calls, start_call) or 0
	local failed_ping = find_nth_control_call(recorded.async_calls, "--app-ping", 2)
	local stopped_ping = find_nth_control_call(recorded.async_calls, "--app-ping", 3)
	assert_true(failed_ping ~= nil, "manual restart should retry after transient app-ping failure")
	assert_true(stopped_ping ~= nil, "manual restart should observe stopped app-ping status")
	assert_true(
		(find_call_index(recorded.async_calls, failed_ping) or 0) < start_index,
		"manual restart should not treat app-ping status 2 as stopped"
	)
	assert_true(
		(find_call_index(recorded.async_calls, stopped_ping) or 0) < start_index,
		"manual restart should wait for explicit stopped app-ping status"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		app_ping_statuses = { 0, 1, 0 },
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for gated restart pause scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(
		count_property_set(recorded.property_sets, "pause", true) == 1,
		"gated restart should start from an armed pause gate"
	)
	recorded.script_messages["subminer-restart"]()
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 0,
		"manual restart should clear a startup gate without resuming playback"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		app_ping_statuses = { 1, 0 },
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			auto_start_visible_overlay = "no",
		},
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for restart ready restore scenario: " .. tostring(err))
	assert_true(
		recorded.script_messages["subminer-toggle"] ~= nil,
		"subminer-toggle script message not registered"
	)
	assert_true(
		recorded.script_messages["subminer-restart"] ~= nil,
		"subminer-restart script message not registered"
	)
	assert_true(
		recorded.script_messages["subminer-autoplay-ready"] ~= nil,
		"subminer-autoplay-ready script message not registered"
	)
	recorded.script_messages["subminer-toggle"]()
	recorded.script_messages["subminer-restart"]()
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 2,
		"manual restart should re-assert visible overlay after launch and readiness even when auto-start visibility is disabled"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		stop_command_fails = true,
		stop_command_stderr = "stop refused",
		option_overrides = {
			binary_path = binary_path,
		},
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for failed restart-stop scenario: " .. tostring(err))
	recorded.script_messages["subminer-restart"]()
	assert_true(find_control_call(recorded.async_calls, "--stop") ~= nil, "restart should attempt stop")
	assert_true(count_start_calls(recorded.async_calls) == 0, "restart should not start overlay when stop fails")
	assert_true(
		has_osd_message(recorded.osd, "SubMiner: Restart failed"),
		"restart stop failure should show failure OSD"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			aniskip_enabled = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for pre-ready duplicate auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	fire_event(recorded, "file-loaded")
	assert_true(recorded.script_messages["subminer-autoplay-ready"] ~= nil, "subminer-autoplay-ready script message not registered")
	assert_true(
		count_start_calls(recorded.async_calls) == 1,
		"pre-ready duplicate auto-start should not issue duplicate --start commands"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", true) == 1,
		"pre-ready duplicate auto-start should not repeat the pause gate"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 0,
		"pre-ready duplicate auto-start should not resume playback before tokenization is ready"
	)
	assert_true(
		count_osd_message(recorded.osd, "SubMiner: Loading subtitle tokenization...") == 1,
		"pre-ready duplicate auto-start should not repeat the loading OSD"
	)
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 1,
		"autoplay-ready should resume the original pre-ready gate"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			aniskip_enabled = "yes",
		},
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for primary subtitle bar binding scenario: " .. tostring(err))
	local binding = nil
	for _, candidate in ipairs(recorded.key_bindings) do
		if candidate.name == "subminer-toggle-primary-subtitle-bar" then
			binding = candidate
			break
		end
	end
	assert_true(binding ~= nil, "primary subtitle bar v binding should be registered")
	assert_true(binding.keys == "v", "primary subtitle bar binding should use bare v")
	assert_true(binding.forced == true, "primary subtitle bar binding should override mpv's built-in v binding")
	binding.fn()
	assert_true(
		count_control_calls(recorded.async_calls, "--toggle-primary-subtitle-bar") == 1,
		"primary subtitle bar binding should issue primary subtitle toggle command"
	)
	assert_true(
		count_control_calls(recorded.async_calls, "--toggle-visible-overlay") == 0,
		"primary subtitle bar binding should not toggle the whole visible overlay"
	)
end

do
	local media_path = "/media/Sample Show S01E01.mkv"
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			aniskip_enabled = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		path = media_path,
		media_title = "Sample Show S01E01",
		mal_lookup_stdout = "__MAL_FOUND__",
		aniskip_stdout = "__ANISKIP_FOUND__",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for same-media reload scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	recorded.script_messages["subminer-autoplay-ready"]()
	fire_event(recorded, "end-file", { reason = "reload" })
	fire_event(recorded, "file-loaded")
	assert_true(
		count_control_calls(recorded.async_calls, "--hide-visible-overlay") == 0,
		"same-media reload should not hide the visible overlay"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", true) == 1,
		"same-media reload should not re-arm pause-until-ready"
	)
	assert_true(
		count_async_curl_for(recorded.async_calls, "api.aniskip.com") == 1,
		"same-media reload should not repeat AniSkip lookup"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "no",
			texthooker_enabled = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for subtitle rearm scenario: " .. tostring(err))
	fire_event(recorded, "start-file")
	assert_true(
		has_property_set(recorded.property_sets, "sub-auto", "fuzzy"),
		"managed start-file should rearm sub-auto before mpv loads tracks"
	)
	assert_true(
		has_property_set(recorded.property_sets, "sid", "auto"),
		"managed start-file should rearm primary subtitle selection before mpv loads tracks"
	)
	assert_true(
		has_property_set(recorded.property_sets, "secondary-sid", "auto"),
		"managed start-file should rearm secondary subtitle selection before mpv loads tracks"
	)
	fire_event(recorded, "file-loaded")
	assert_true(
		count_property_set(recorded.property_sets, "sid", "auto") == 1,
		"managed file-loaded should not reset primary subtitle selection after mpv loads tracks"
	)
	assert_true(
		count_property_set(recorded.property_sets, "secondary-sid", "auto") == 1,
		"managed file-loaded should not reset secondary subtitle selection after mpv loads tracks"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for attached subtitle rearm scenario: " .. tostring(err))
	fire_event(recorded, "start-file")
	fire_event(recorded, "file-loaded")
	assert_true(
		count_property_set(recorded.property_sets, "sid", "auto") == 1,
		"attached background app path should select primary subtitle before load only"
	)
	assert_true(
		count_property_set(recorded.property_sets, "secondary-sid", "auto") == 1,
		"attached background app path should select secondary subtitle before load only"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "no",
			socket_path = "/tmp/subminer-socket",
			texthooker_enabled = "no",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for disabled texthooker auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(start_call ~= nil, "disabled texthooker auto-start should still issue --start command")
	assert_true(not call_has_arg(start_call, "--texthooker"), "disabled texthooker should not include --texthooker on --start")
	assert_true(find_control_call(recorded.async_calls, "--texthooker") == nil, "disabled texthooker should not issue a helper texthooker command")
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			aniskip_enabled = "yes",
		},
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for non-subminer file-load scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(not has_sync_command(recorded.sync_calls, "ps"), "file-loaded should avoid synchronous process checks")
	assert_true(not has_sync_command(recorded.sync_calls, "curl"), "file-loaded should avoid synchronous AniSkip network calls")
	assert_true(
		not has_async_curl_for(recorded.async_calls, "myanimelist.net/search/prefix.json"),
		"file-loaded without SubMiner context should skip AniSkip MAL lookup"
	)
	assert_true(
		not has_async_curl_for(recorded.async_calls, "api.aniskip.com"),
		"file-loaded without SubMiner context should skip AniSkip API lookup"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "no",
			texthooker_enabled = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		path = "https://www.youtube.com/watch?v=lJI7uL4JDkE",
		media_title = "【文字起こし】マジで役立つ！！恋愛術！【告radio】",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for URL overlay-start AniSkip scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(find_start_call(recorded.async_calls) ~= nil, "URL auto-start should still invoke --start command")
	assert_true(
		not has_async_curl_for(recorded.async_calls, "myanimelist.net/search/prefix.json"),
		"URL playback should skip AniSkip MAL lookup even after overlay-start"
	)
	assert_true(
		not has_async_curl_for(recorded.async_calls, "api.aniskip.com"),
		"URL playback should skip AniSkip API lookup even after overlay-start"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			aniskip_enabled = "yes",
		},
		media_title = "Sample Show S01E01",
		mal_lookup_stdout = "__MAL_FOUND__",
		aniskip_stdout = "__ANISKIP_FOUND__",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for script-message AniSkip scenario: " .. tostring(err))
	assert_true(recorded.script_messages["subminer-aniskip-refresh"] ~= nil, "subminer-aniskip-refresh script message not registered")
	recorded.script_messages["subminer-aniskip-refresh"]()
	assert_true(not has_sync_command(recorded.sync_calls, "curl"), "AniSkip refresh should not perform synchronous curl calls")
	assert_true(has_async_command(recorded.async_calls, "curl"), "AniSkip refresh should perform async curl calls")
	assert_true(
		has_async_curl_for(recorded.async_calls, "myanimelist.net/search/prefix.json"),
		"AniSkip refresh should perform MAL lookup even when app is not running"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
			aniskip_enabled = "yes",
		},
		media_title = "Sample Show S01E01",
		time_pos = 13,
		mal_lookup_stdout = "__MAL_FOUND__",
		aniskip_stdout = "__ANISKIP_FOUND__",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for default AniSkip keybinding scenario: " .. tostring(err))
	assert_true(
		has_key_binding(recorded, "TAB", "subminer-skip-intro"),
		"default AniSkip keybinding should register TAB"
	)
	assert_true(
		not has_key_binding(recorded, "y-k", "subminer-skip-intro-fallback"),
		"default AniSkip keybinding should not also register legacy y-k fallback"
	)
	recorded.script_messages["subminer-aniskip-refresh"]()
	fire_observer(recorded, "time-pos", 13)
	assert_true(
		has_osd_message(recorded.osd, "You can skip by pressing TAB"),
		"default AniSkip prompt should mention TAB"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "no",
			texthooker_enabled = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for visible auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(start_call ~= nil, "auto-start should issue --start command")
	assert_true(
		call_has_arg(start_call, "--background"),
		"auto-start should launch SubMiner in background/tray mode"
	)
	assert_true(
		call_has_arg(start_call, "--managed-playback"),
		"auto-start should mark SubMiner as managed playback"
	)
	assert_true(call_has_arg(start_call, "--texthooker"), "auto-start should include --texthooker on the main --start command when enabled")
	assert_true(find_control_call(recorded.async_calls, "--texthooker") == nil, "auto-start should not issue a separate texthooker helper command")
	assert_true(
		call_has_arg(start_call, "--show-visible-overlay"),
		"auto-start with visible overlay enabled should include --show-visible-overlay on --start"
	)
	assert_true(
		not call_has_arg(start_call, "--hide-visible-overlay"),
		"auto-start with visible overlay enabled should not include --hide-visible-overlay on --start"
	)
	assert_true(
		find_control_call(recorded.async_calls, "--show-visible-overlay") ~= nil,
		"auto-start with visible overlay enabled should issue a separate --show-visible-overlay command"
	)
	assert_true(
		not has_property_set(recorded.property_sets, "pause", true),
		"auto-start visible overlay should not force pause without explicit pause-until-ready option"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for duplicate auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	fire_event(recorded, "file-loaded")
	assert_true(
		count_start_calls(recorded.async_calls) == 1,
		"duplicate file-loaded events should not issue duplicate --start commands while overlay is already running"
	)
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 2,
		"duplicate auto-start should re-assert visible overlay state when overlay is already running"
	)
	assert_true(
		count_osd_message(recorded.osd, "SubMiner: Already running") == 0,
		"duplicate auto-start events should not show Already running OSD"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for duplicate auto-start pause-until-ready scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(recorded.script_messages["subminer-autoplay-ready"] ~= nil, "subminer-autoplay-ready script message not registered")
	recorded.script_messages["subminer-autoplay-ready"]()
	fire_event(recorded, "file-loaded")
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_start_calls(recorded.async_calls) == 1,
		"duplicate pause-until-ready auto-start should not issue duplicate --start commands while overlay is already running"
	)
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 3,
		"duplicate pause-until-ready auto-start should re-assert visible overlay on initial start, ready, and later file load"
	)
	assert_true(
		count_osd_message(recorded.osd, "SubMiner: Loading subtitle tokenization...") == 1,
		"duplicate pause-until-ready auto-start should not repeat tokenization loading gate after overlay is running"
	)
	assert_true(
		count_osd_message(recorded.osd, "SubMiner: Subtitle tokenization ready") == 1,
		"duplicate pause-until-ready auto-start should not wait for a second readiness signal after overlay is running"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", true) == 1,
		"duplicate pause-until-ready auto-start should not force pause after overlay is running"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 1,
		"duplicate pause-until-ready auto-start should not resume a gate that was never rearmed"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for pause-until-ready scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(
		has_property_set(recorded.property_sets, "pause", true),
		"pause-until-ready auto-start should pause mpv before overlay ready"
	)
	assert_true(recorded.script_messages["subminer-autoplay-ready"] ~= nil, "subminer-autoplay-ready script message not registered")
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		has_property_set(recorded.property_sets, "pause", false),
		"autoplay-ready script message should resume mpv playback"
	)
	assert_true(
		has_osd_message(recorded.osd, "SubMiner: Loading subtitle tokenization..."),
		"pause-until-ready auto-start should show loading OSD message"
	)
	assert_true(
		not has_osd_message(recorded.osd, "SubMiner: Starting..."),
		"pause-until-ready auto-start should avoid replacing loading OSD with generic starting OSD"
	)
	assert_true(
		has_osd_message(recorded.osd, "SubMiner: Subtitle tokenization ready"),
		"autoplay-ready should show loaded OSD message"
	)
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 2,
		"autoplay-ready should re-assert visible overlay state"
	)
	assert_true(
		#recorded.periodic_timers == 1,
		"pause-until-ready auto-start should create periodic loading OSD refresher"
	)
	assert_true(
		recorded.periodic_timers[1].killed == true,
		"autoplay-ready should stop periodic loading OSD refresher"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for duplicate autoplay-ready scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(recorded.script_messages["subminer-autoplay-ready"] ~= nil, "subminer-autoplay-ready script message not registered")
	recorded.script_messages["subminer-autoplay-ready"]()
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 2,
		"duplicate autoplay-ready signals should not repeatedly spawn visible overlay restore commands"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		paused = true,
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for pre-paused pause-until-ready scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(
		count_property_set(recorded.property_sets, "pause", true) == 1,
		"pre-paused pause-until-ready should still arm the gate"
	)
	assert_true(recorded.script_messages["subminer-autoplay-ready"] ~= nil, "subminer-autoplay-ready script message not registered")
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 0,
		"pre-paused pause-until-ready should leave playback paused when ready"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for manual toggle-off ready scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(recorded.script_messages["subminer-toggle"] ~= nil, "subminer-toggle script message not registered")
	recorded.script_messages["subminer-toggle"]()
	assert_true(
		count_control_calls(recorded.async_calls, "--toggle-visible-overlay") == 1,
		"manual toggle should use explicit visible-overlay toggle command"
	)
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 1,
		"manual toggle-off before readiness should suppress ready-time visible overlay restore"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 0,
		"manual toggle-off before readiness should not resume playback when readiness arrives"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(
		recorded ~= nil,
		"plugin failed to load for repeated ready restore suppression scenario: " .. tostring(err)
	)
	fire_event(recorded, "file-loaded")
	assert_true(recorded.script_messages["subminer-toggle"] ~= nil, "subminer-toggle script message not registered")
	recorded.script_messages["subminer-toggle"]()
	recorded.script_messages["subminer-autoplay-ready"]()
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 1,
		"manual toggle-off should suppress repeated ready-time visible overlay restores for the same session"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		path = "/media/jellyfin-stream.m3u8",
		media_title = "Jellyfin Episode",
		paused = true,
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for manual hide duplicate auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	assert_true(recorded.script_messages["subminer-toggle"] ~= nil, "subminer-toggle script message not registered")
	recorded.script_messages["subminer-toggle"]()
	fire_event(recorded, "file-loaded")
	recorded.script_messages["subminer-autoplay-ready"]()
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 1,
		"manual toggle-off should suppress duplicate auto-start visible overlay reassertion"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 0,
		"manual toggle-off followed by duplicate auto-start should keep paused playback paused"
	)
end

do
	local media_path = "/media/jellyfin-redirect.m3u8"
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		path = media_path,
		media_title = "Jellyfin Redirect",
		paused = true,
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for manual hide same-media reload scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	recorded.script_messages["subminer-autoplay-ready"]()
	recorded.script_messages["subminer-toggle"]()
	fire_event(recorded, "end-file", { reason = "redirect" })
	fire_event(recorded, "file-loaded")
	assert_true(
		count_control_calls(recorded.async_calls, "--show-visible-overlay") == 2,
		"manual toggle-off should suppress same-media reload visible overlay reassertion"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 0,
		"manual toggle-off followed by same-media reload should keep paused playback paused"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "no",
		},
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for manual toggle command scenario: " .. tostring(err))
	assert_true(recorded.script_messages["subminer-toggle"] ~= nil, "subminer-toggle script message not registered")
	recorded.script_messages["subminer-toggle"]()
	assert_true(
		count_control_calls(recorded.async_calls, "--toggle-visible-overlay") == 1,
		"script-message toggle should issue explicit visible-overlay toggle command"
	)
	assert_true(
		count_control_calls(recorded.async_calls, "--toggle") == 0,
		"script-message toggle should not issue legacy generic toggle command"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			auto_start_pause_until_ready = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for pause cleanup scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	fire_event(recorded, "end-file")
	assert_true(
		count_property_set(recorded.property_sets, "pause", true) == 1,
		"pause cleanup scenario should force pause while waiting for tokenization"
	)
	assert_true(
		count_property_set(recorded.property_sets, "pause", false) == 1,
		"ending file while gate is armed should clear forced pause state"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for shutdown-managed-background scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	fire_event(recorded, "end-file", { reason = "quit" })
	assert_true(
		find_control_call(recorded.async_calls, "--hide-visible-overlay") == nil,
		"mpv quit end-file should not spawn hide-visible-overlay helper commands"
	)
	fire_event(recorded, "shutdown")
	assert_true(
		find_control_call(recorded.async_calls, "--stop") == nil,
		"mpv shutdown should leave managed-playback ownership to the app process"
	)
	assert_true(
		find_control_call(recorded.async_calls, "--hide-visible-overlay") == nil,
		"mpv shutdown should not spawn hide-visible-overlay helper commands"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "/opt/SubMiner/subminer --background\n",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for shutdown-borrowed-background scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(start_call ~= nil, "auto-start should attach playback to the existing app")
	assert_true(
		not call_has_arg(start_call, "--background"),
		"borrowed app auto-start should not use the background launch wrapper"
	)
	assert_true(
		call_has_arg(start_call, "--managed-playback"),
		"borrowed app auto-start should still attach managed playback to the existing app"
	)
	fire_event(recorded, "end-file", { reason = "quit" })
	fire_event(recorded, "shutdown")
	assert_true(
		find_control_call(recorded.async_calls, "--stop") == nil,
		"mpv shutdown should leave a pre-existing background SubMiner process running"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "no",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for hidden auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(start_call ~= nil, "auto-start should issue --start command")
	assert_true(
		call_has_arg(start_call, "--background"),
		"auto-start should launch SubMiner in background mode"
	)
	assert_true(
		call_has_arg(start_call, "--managed-playback"),
		"auto-start should mark SubMiner as managed playback"
	)
	assert_true(
		call_has_arg(start_call, "--hide-visible-overlay"),
		"auto-start with visible overlay disabled should include --hide-visible-overlay on --start"
	)
	assert_true(
		not call_has_arg(start_call, "--show-visible-overlay"),
		"auto-start with visible overlay disabled should not include --show-visible-overlay on --start"
	)
	assert_true(
		find_control_call(recorded.async_calls, "--hide-visible-overlay") ~= nil,
		"auto-start with visible overlay disabled should issue a separate --hide-visible-overlay command"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "/tmp/other.sock",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for mismatched socket auto-start scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(
		start_call == nil,
		"auto-start should be skipped when mpv input-ipc-server does not match configured socket_path"
	)
	assert_true(
		not has_property_set(recorded.property_sets, "sid", "auto"),
		"subtitle rearm should not run when mpv input-ipc-server does not match configured socket_path"
	)
	assert_true(
		not has_property_set(recorded.property_sets, "pause", true),
		"pause-until-ready gate should not arm when socket_path does not match"
	)
end

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		option_overrides = {
			binary_path = binary_path,
		},
		input_ipc_server = "/tmp/subminer-socket",
		media_title = "Random Movie",
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for default config scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(
		start_call == nil,
		"plugin should not auto-start from built-in defaults without managed config script opts"
	)
end

do
	local recorded, err = run_plugin_scenario({
		platform = "windows",
		process_list = "",
		option_overrides = {
			binary_path = "C:/Users/test/AppData/Local/Programs/SubMiner/SubMiner.exe",
			auto_start = "yes",
			auto_start_visible_overlay = "yes",
			socket_path = "/tmp/subminer-socket",
		},
		input_ipc_server = "\\\\.\\pipe\\subminer-socket",
		media_title = "Random Movie",
		files = {
			["C:/Users/test/AppData/Local/Programs/SubMiner/SubMiner.exe"] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for Windows legacy socket config scenario: " .. tostring(err))
	fire_event(recorded, "file-loaded")
	local start_call = find_start_call(recorded.async_calls)
	assert_true(
		start_call ~= nil,
		"Windows plugin should normalize legacy /tmp socket_path values to the named pipe default"
	)
end

print("plugin start gate regression tests: OK")
