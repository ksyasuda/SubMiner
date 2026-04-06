local function assert_equals(expected, actual, message)
	if expected == actual then
		return
	end
	error(message .. " (expected " .. tostring(expected) .. ", got " .. tostring(actual) .. ")")
end

local process_module = dofile("plugin/subminer/process.lua")
local process = process_module.create({
	mp = {},
	opts = {},
	state = {},
	binary = {},
	environment = {},
	options_helper = {
		coerce_bool = function(value, default)
			if value == nil then
				return default
			end
			return value == true or value == "true" or value == "yes" or value == "1"
		end,
	},
	log = {
		subminer_log = function() end,
		show_osd = function() end,
		normalize_log_level = function(value)
			return value or "info"
		end,
	},
})

local overrides = process.parse_start_script_message_overrides("backend=windows")
assert_equals("windows", overrides.backend, "expected backend=windows override to be accepted")

print("plugin process override tests: OK")
