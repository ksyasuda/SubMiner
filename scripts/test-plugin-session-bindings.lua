package.path = "plugin/subminer/?.lua;" .. package.path

local session_bindings = require("session_bindings")

local function assert_true(condition, message)
	if condition then
		return
	end
	error(message)
end

local artifact_path = ".tmp/test-plugin-session-bindings.json"
local is_windows = package.config:sub(1, 1) == "\\"
local mkdir_cmd = is_windows and "mkdir .tmp >NUL 2>NUL" or "mkdir -p .tmp"
os.execute(mkdir_cmd)
local handle = assert(io.open(artifact_path, "w"))
handle:write("__SESSION_BINDINGS__")
handle:close()

local recorded = {
	bindings = {},
	removed = {},
	async_calls = {},
	osd = {},
}

local mp = {}

function mp.add_forced_key_binding(keys, name, fn)
	recorded.bindings[#recorded.bindings + 1] = {
		keys = keys,
		name = name,
		fn = fn,
	}
end

function mp.remove_key_binding(name)
	recorded.removed[#recorded.removed + 1] = name
end

function mp.add_timeout(seconds, callback)
	return {
		seconds = seconds,
		callback = callback,
		killed = false,
		kill = function(self)
			self.killed = true
		end,
	}
end

function mp.osd_message(message)
	recorded.osd[#recorded.osd + 1] = message
end

local ctx = {
	mp = mp,
	utils = {
		parse_json = function(raw)
			if raw ~= "__SESSION_BINDINGS__" then
				return nil, "unexpected artifact"
			end
			return {
				numericSelectionTimeoutMs = 3000,
				bindings = {
					{
						key = {
							code = "KeyS",
							modifiers = { "ctrl", "shift" },
						},
						actionType = "session-action",
						actionId = "mineSentenceMultiple",
					},
				},
			}, nil
		end,
	},
	state = {
		binary_path = "/tmp/subminer",
		session_binding_names = {},
		session_numeric_binding_names = {},
		session_numeric_selection = nil,
	},
	process = {
		check_binary_available = function()
			return true
		end,
		run_binary_command_async = function(args)
			recorded.async_calls[#recorded.async_calls + 1] = args
		end,
	},
	environment = {
		resolve_session_bindings_artifact_path = function()
			return artifact_path
		end,
	},
	log = {
		subminer_log = function() end,
		show_osd = function(message)
			recorded.osd[#recorded.osd + 1] = message
		end,
	},
}

local bindings = session_bindings.create(ctx)
assert_true(bindings.register_bindings(), "session bindings should register")

local starter = nil
for _, binding in ipairs(recorded.bindings) do
	if binding.keys == "Ctrl+Shift+s" then
		starter = binding
		break
	end
end
assert_true(starter ~= nil, "multi-mine starter binding should be registered")

starter.fn()

local modified_digit = nil
for _, binding in ipairs(recorded.bindings) do
	if binding.keys == "Ctrl+Shift+3" then
		modified_digit = binding
		break
	end
end
assert_true(modified_digit ~= nil, "numeric selection should bind Ctrl+Shift+3")

modified_digit.fn()

local call = recorded.async_calls[#recorded.async_calls]
assert_true(call ~= nil, "modified digit should invoke CLI action")
assert_true(call[1] == "/tmp/subminer", "CLI action should use configured binary")
assert_true(call[2] == "--mine-sentence-count", "CLI action should mine sentence count")
assert_true(call[3] == "3", "CLI action should pass selected count")

print("plugin session binding regression tests: OK")
