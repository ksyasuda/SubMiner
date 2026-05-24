local function assert_true(condition, message)
	if condition then
		return
	end
	error(message)
end

local function has_flag(call, flag)
	local args = call.args or {}
	for _, arg in ipairs(args) do
		if arg == flag then
			return true
		end
	end
	return false
end

local function has_timeout(timeouts, target)
	for _, value in ipairs(timeouts) do
		if math.abs(value - target) < 0.0001 then
			return true
		end
	end
	return false
end

local recorded = {
	async_calls = {},
	timeouts = {},
	logs = {},
}

local start_attempts = 0

local mp = {}

function mp.command_native_async(command, callback)
	recorded.async_calls[#recorded.async_calls + 1] = command

	local success = true
	local result = { status = 0, stdout = "", stderr = "" }
	local err = nil

	if has_flag(command, "--start") then
		start_attempts = start_attempts + 1
		if start_attempts == 1 then
			success = false
			result = { status = 1, stdout = "", stderr = "startup-not-ready" }
			err = "startup-not-ready"
		end
	end

	if callback then
		callback(success, result, err)
	end
end

function mp.add_timeout(seconds, callback)
	recorded.timeouts[#recorded.timeouts + 1] = seconds
	if callback then
		callback()
	end
end

local process_module = dofile("plugin/subminer/process.lua")
local process = process_module.create({
	mp = mp,
	opts = {
		backend = "x11",
		socket_path = "/tmp/subminer.sock",
		log_level = "debug",
		texthooker_enabled = true,
		texthooker_port = 5174,
		auto_start_visible_overlay = false,
	},
	state = {
		binary_path = "/tmp/subminer",
		overlay_running = false,
		texthooker_running = false,
	},
	binary = {
		ensure_binary_available = function()
			return true
		end,
	},
		environment = {
			detect_backend = function()
				return "x11"
			end,
			is_linux = function()
				return false
			end,
			is_subminer_app_running_async = function(callback)
				callback(false)
			end,
	},
	options_helper = {
		coerce_bool = function(value, default_value)
			if value == true or value == "yes" or value == "true" then
				return true
			end
			if value == false or value == "no" or value == "false" then
				return false
			end
			return default_value
		end,
	},
	log = {
		subminer_log = function(_level, _scope, line)
			recorded.logs[#recorded.logs + 1] = line
		end,
		show_osd = function(_) end,
		normalize_log_level = function(value)
			return value or "info"
		end,
	},
})

process.start_overlay()

assert_true(start_attempts == 2, "expected start overlay command retry after readiness failure")
assert_true(not has_timeout(recorded.timeouts, 0.35), "fixed texthooker wait (0.35s) should be removed")
assert_true(not has_timeout(recorded.timeouts, 0.6), "fixed startup visibility delay (0.6s) should be removed")

local retry_timeout_seen = false
for _, timeout_seconds in ipairs(recorded.timeouts) do
	if timeout_seconds > 0 and timeout_seconds <= 0.25 then
		retry_timeout_seen = true
		break
	end
end
assert_true(retry_timeout_seen, "expected shorter bounded retry timeout")

do
	local visibility_state = {
		binary_path = "/tmp/subminer",
		overlay_running = true,
		texthooker_running = false,
		visible_overlay_requested = false,
	}
	local visibility_calls = {}
	local visibility_mp = {}

	function visibility_mp.command_native_async(command, callback)
		visibility_calls[#visibility_calls + 1] = command
		if callback then
			callback(false, { status = 1, stdout = "", stderr = "failed" }, "failed")
		end
	end

	local visibility_process = process_module.create({
		mp = visibility_mp,
		opts = {
			backend = "x11",
			socket_path = "/tmp/subminer.sock",
			log_level = "debug",
			texthooker_enabled = true,
			texthooker_port = 5174,
			auto_start_visible_overlay = false,
		},
		state = visibility_state,
		binary = {
			ensure_binary_available = function()
				return true
			end,
		},
		environment = {
			detect_backend = function()
				return "x11"
			end,
			is_linux = function()
				return false
			end,
			is_subminer_app_running_async = function(callback)
				callback(true)
			end,
		},
		options_helper = {
			coerce_bool = function(value, default_value)
				if value == true or value == "yes" or value == "true" then
					return true
				end
				if value == false or value == "no" or value == "false" then
					return false
				end
				return default_value
			end,
		},
		log = {
			subminer_log = function(_level, _scope, line)
				recorded.logs[#recorded.logs + 1] = line
			end,
			show_osd = function(_) end,
			normalize_log_level = function(value)
				return value or "info"
			end,
		},
	})

	visibility_process.run_control_command_async("show-visible-overlay")

	assert_true(#visibility_calls == 1, "expected visible overlay command to run")
	assert_true(
		visibility_state.visible_overlay_requested == false,
		"failed visible-overlay command should not update requested visibility state"
	)
end

print("plugin process retry regression tests: OK")
