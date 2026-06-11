package.path = "plugin/subminer/?.lua;" .. package.path

local process_module = require("process")
local options_helper = require("options")

local function assert_true(condition, message)
	if condition then
		return
	end
	error(message or "assert_true failed")
end

local function has_arg(args, target)
	for _, value in ipairs(args or {}) do
		if value == target then
			return true
		end
	end
	return false
end

local function count_feedback(recorded, target)
	local count = 0
	for _, message in ipairs(recorded.feedback) do
		if message == target then
			count = count + 1
		end
	end
	return count
end

local function create_restart_runtime(config)
	config = config or {}
	local recorded = {
		async_calls = {},
		feedback = {},
		osd = {},
		periodic_timers = {},
	}
	local app_ping_index = 0
	local opts = {
		binary_path = "/tmp/SubMiner",
		socket_path = "/tmp/subminer-socket",
		backend = "x11",
		osd_messages = config.osd_messages == true,
		texthooker_enabled = false,
		log_level = "info",
	}
	local state = {
		binary_path = opts.binary_path,
		overlay_running = true,
		texthooker_running = false,
	}

	local mp = {}

	function mp.command_native_async(command, callback)
		recorded.async_calls[#recorded.async_calls + 1] = command
		local args = command.args or {}
		if has_arg(args, "--playback-feedback") then
			recorded.feedback[#recorded.feedback + 1] = args[#args]
			callback(true, { status = 0, stdout = "", stderr = "" }, nil)
			return
		end
		if has_arg(args, "--app-ping") then
			app_ping_index = app_ping_index + 1
			local statuses = config.app_ping_statuses or { 1, 0 }
			local status = statuses[app_ping_index] or statuses[#statuses]
			callback(status == 0, { status = status, stdout = "", stderr = "" }, nil)
			return
		end
		if has_arg(args, "--show-visible-overlay") and not has_arg(args, "--start") then
			local status = config.show_visible_overlay_status or 0
			callback(status == 0, { status = status, stdout = "", stderr = "" }, nil)
			return
		end
		callback(true, { status = 0, stdout = "", stderr = "" }, nil)
	end

	function mp.add_timeout(_, callback)
		if config.run_timeouts_immediately and callback then
			callback()
		end
		return {
			killed = false,
			kill = function(self)
				self.killed = true
			end,
			callback = callback,
		}
	end

	function mp.add_periodic_timer()
		local timer = {
			killed = false,
			kill = function(self)
				self.killed = true
			end,
		}
		recorded.periodic_timers[#recorded.periodic_timers + 1] = timer
		return timer
	end

	function mp.get_property(name)
		if name == "input-ipc-server" then
			return opts.socket_path
		end
		return ""
	end

	function mp.get_time()
		return 1
	end

	function mp.set_property_native() end

	local process = process_module.create({
		mp = mp,
		utils = {},
		opts = opts,
		state = state,
		binary = {
			ensure_binary_available = function()
				return true
			end,
		},
		environment = {
			is_linux = function()
				return false
			end,
			detect_backend = function()
				return "x11"
			end,
			resolve_subminer_config_dir = function()
				return "/tmp"
			end,
			join_path = function(...)
				return table.concat({ ... }, "/")
			end,
		},
		options_helper = options_helper,
		log = {
			normalize_log_level = function(level)
				return level or "info"
			end,
			subminer_log = function() end,
			show_osd = function(message, options)
				if opts.osd_messages or (options and options.force == true) then
					recorded.osd[#recorded.osd + 1] = message
				end
			end,
		},
	})

	return {
		process = process,
		recorded = recorded,
	}
end

do
	local runtime = create_restart_runtime({ osd_messages = false })

	runtime.process.restart_overlay()

	assert_true(
		runtime.recorded.osd[1] == "Overlay loading |",
		"restart should show the forced overlay loading OSD while the overlay reloads"
	)
	assert_true(
		#runtime.recorded.periodic_timers == 1,
		"restart should refresh the forced overlay loading OSD while the overlay reloads"
	)
	assert_true(
		runtime.recorded.feedback[1] == "Restarting...",
		"restart should route progress through playback feedback"
	)
	assert_true(
		runtime.recorded.feedback[#runtime.recorded.feedback] == "Restarted successfully",
		"restart should route success through playback feedback"
	)
	assert_true(
		runtime.recorded.periodic_timers[1].killed ~= true,
		"restart should keep the loading OSD alive until the overlay reports ready"
	)
end

do
	local runtime = create_restart_runtime({
		osd_messages = false,
		show_visible_overlay_status = 1,
	})

	runtime.process.restart_overlay()

	assert_true(
		count_feedback(runtime.recorded, "Restarted successfully") == 0,
		"restart should not show success feedback when show-visible-overlay fails after ready ping"
	)
	assert_true(
		runtime.recorded.feedback[#runtime.recorded.feedback] == "Restart failed",
		"restart should show failure feedback when show-visible-overlay fails after ready ping"
	)
end

do
	local statuses = { 1 }
	for _ = 1, 20 do
		statuses[#statuses + 1] = 1
	end
	local runtime = create_restart_runtime({
		app_ping_statuses = statuses,
		osd_messages = false,
		run_timeouts_immediately = true,
		show_visible_overlay_status = 1,
	})

	runtime.process.restart_overlay()

	assert_true(
		count_feedback(runtime.recorded, "Restarted successfully") == 0,
		"restart should not show success feedback when fallback show-visible-overlay fails after ping timeout"
	)
	assert_true(
		runtime.recorded.feedback[#runtime.recorded.feedback] == "Restart failed",
		"restart should show failure feedback when fallback show-visible-overlay fails after ping timeout"
	)
end

print("plugin restart feedback tests: OK")
