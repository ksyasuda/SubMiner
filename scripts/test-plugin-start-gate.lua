local function run_plugin_scenario(config)
	config = config or {}

	local recorded = {
		async_calls = {},
		sync_calls = {},
		script_messages = {},
		osd = {},
		logs = {},
	}

	local function make_mp_stub()
		local mp = {}

		function mp.get_property(name)
			if name == "platform" then
				return config.platform or "linux"
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

		function mp.get_property_native(_name)
			return config.chapter_list or {}
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
				return { status = 0, stdout = "{}", stderr = "" }
			end
			return { status = 0, stdout = "", stderr = "" }
		end

		function mp.command_native_async(command, callback)
			recorded.async_calls[#recorded.async_calls + 1] = command
			if callback then
				callback(true, { status = 0, stdout = "", stderr = "" }, nil)
			end
		end

		function mp.add_timeout(_seconds, callback)
			if callback then
				callback()
			end
		end

		function mp.register_script_message(name, fn)
			recorded.script_messages[name] = fn
		end

		function mp.add_key_binding(_keys, _name, _fn) end
		function mp.register_event(_name, _fn) end
		function mp.add_hook(_name, _prio, _fn) end
		function mp.observe_property(_name, _kind, _fn) end
		function mp.osd_message(message, _duration)
			recorded.osd[#recorded.osd + 1] = message
		end
		function mp.get_time()
			return 0
		end
		function mp.commandv(...) end
		function mp.set_property_native(...) end
		function mp.set_property(...) end
		function mp.set_osd_ass(...) end
		function mp.get_property_number(_name, default)
			return default
		end
		function mp.get_property_bool(_name, default)
			return default
		end
		function mp.get_script_name()
			return "subminer"
		end
		function mp.get_script_directory()
			return "plugin/subminer"
		end

		return mp
	end

	local mp = make_mp_stub()
	local options = {}
	local utils = {}

	function options.read_options(target, _name)
		if config.socket_path then
			target.socket_path = config.socket_path
		end
		if config.binary_path then
			target.binary_path = config.binary_path
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

	function utils.parse_json(_json)
		return {}, nil
	end

	package.loaded["mp"] = nil
	package.loaded["mp.input"] = nil
	package.loaded["mp.msg"] = nil
	package.loaded["mp.options"] = nil
	package.loaded["mp.utils"] = nil
	for key, _ in pairs(package.loaded) do
		if key:match("^subminer") then
			package.loaded[key] = nil
		end
	end
	local plugin_modules = {
		"init",
		"bootstrap",
		"options",
		"state",
		"log",
		"binary",
		"environment",
		"process",
		"aniskip",
		"aniskip_match",
		"hover",
		"ui",
		"messages",
		"lifecycle",
	}
	for _, module_name in ipairs(plugin_modules) do
		package.loaded[module_name] = nil
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

local function has_command_flag(calls, flag)
	for _, call in ipairs(calls) do
		local args = call.args or {}
		for i = 1, #args do
			if args[i] == flag then
				return true
			end
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

local binary_path = "/tmp/subminer-binary"

do
	local recorded, err = run_plugin_scenario({
		process_list = "",
		binary_path = binary_path,
		files = {
			[binary_path] = true,
		},
	})
	assert_true(recorded ~= nil, "plugin failed to load for cold-start scenario: " .. tostring(err))
	assert_true(recorded.script_messages["subminer-start"] ~= nil, "subminer-start script message not registered")
	assert_true(recorded.script_messages["subminer-stop"] ~= nil, "subminer-stop script message not registered")
	recorded.script_messages["subminer-start"]("texthooker=no")
	assert_true(find_start_call(recorded.async_calls) ~= nil, "expected cold-start to invoke --start command when process is absent")
	recorded.script_messages["subminer-stop"]()
	assert_true(has_command_flag(recorded.sync_calls, "--stop"), "expected stop message to invoke --stop command")
	assert_true(
		not has_sync_command(recorded.sync_calls, "ps"),
		"expected cold-start start command to avoid synchronous process list scan"
	)
end

print("plugin start gate regression tests: OK")
