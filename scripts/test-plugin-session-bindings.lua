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
	mpv_commands = {},
	osd = {},
	overlay_toggles = 0,
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

function mp.commandv(...)
	recorded.mpv_commands[#recorded.mpv_commands + 1] = { ... }
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
							code = "KeyO",
							modifiers = { "alt", "shift" },
						},
						actionType = "session-action",
						actionId = "toggleVisibleOverlay",
					},
					{
						key = {
							code = "KeyS",
							modifiers = { "ctrl", "shift" },
						},
						actionType = "session-action",
						actionId = "mineSentenceMultiple",
					},
					{
						key = {
							code = "Space",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "cycle", "pause" },
					},
					{
						key = {
							code = "KeyF",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "cycle", "fullscreen" },
					},
					{
						key = {
							code = "KeyJ",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "cycle", "sid" },
					},
					{
						key = {
							code = "KeyJ",
							modifiers = { "shift" },
						},
						actionType = "mpv-command",
						command = { "cycle", "secondary-sid" },
					},
					{
						key = {
							code = "ArrowRight",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "seek", 5 },
					},
					{
						key = {
							code = "ArrowLeft",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "seek", -5 },
					},
					{
						key = {
							code = "ArrowUp",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "seek", 60 },
					},
					{
						key = {
							code = "ArrowDown",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "seek", -60 },
					},
					{
						key = {
							code = "KeyH",
							modifiers = { "shift" },
						},
						actionType = "mpv-command",
						command = { "sub-seek", -1 },
					},
					{
						key = {
							code = "KeyL",
							modifiers = { "shift" },
						},
						actionType = "mpv-command",
						command = { "sub-seek", 1 },
					},
					{
						key = {
							code = "BracketRight",
							modifiers = { "shift" },
						},
						actionType = "session-action",
						actionId = "shiftSubDelayNextLine",
					},
					{
						key = {
							code = "BracketLeft",
							modifiers = { "shift" },
						},
						actionType = "session-action",
						actionId = "shiftSubDelayPrevLine",
					},
					{
						key = {
							code = "KeyC",
							modifiers = { "ctrl", "alt" },
						},
						actionType = "session-action",
						actionId = "openYoutubePicker",
					},
					{
						key = {
							code = "KeyP",
							modifiers = { "ctrl", "alt" },
						},
						actionType = "session-action",
						actionId = "openPlaylistBrowser",
					},
					{
						key = {
							code = "KeyH",
							modifiers = { "ctrl", "shift" },
						},
						actionType = "session-action",
						actionId = "replayCurrentSubtitle",
					},
					{
						key = {
							code = "KeyL",
							modifiers = { "ctrl", "shift" },
						},
						actionType = "session-action",
						actionId = "playNextSubtitle",
					},
					{
						key = {
							code = "KeyQ",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "quit" },
					},
					{
						key = {
							code = "KeyW",
							modifiers = { "ctrl" },
						},
						actionType = "mpv-command",
						command = { "quit" },
					},
					{
						key = {
							code = "MBTN_BACK",
							modifiers = {},
						},
						actionType = "mpv-command",
						command = { "sub-seek", -1 },
					},
					{
						key = {
							code = "KeyW",
							modifiers = {},
						},
						actionType = "session-action",
						actionId = "markWatched",
					},
					{
						key = {
							code = "KeyD",
							modifiers = { "ctrl" },
						},
						actionType = "session-action",
						actionId = "openCharacterDictionaryManager",
						cliArgs = { "--session-action", '{"actionId":"openCharacterDictionaryManager"}' },
					},
					{
						key = {
							code = "F12",
							modifiers = { "ctrl", "alt" },
						},
						actionType = "session-action",
						actionId = "openFuturePanel",
						cliArgs = { "--session-action", '{"actionId":"openFuturePanel"}' },
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
		toggle_overlay = function()
			recorded.overlay_toggles = recorded.overlay_toggles + 1
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

local function find_binding(keys)
	for _, binding in ipairs(recorded.bindings) do
		if binding.keys == keys then
			return binding
		end
	end
	return nil
end

local starter = find_binding("Ctrl+S")
assert_true(starter ~= nil, "multi-mine starter binding should be registered")

local expected_mpv_bindings = {
	{ keys = "SPACE", command = { "cycle", "pause" } },
	{ keys = "f", command = { "cycle", "fullscreen" } },
	{ keys = "j", command = { "cycle", "sid" } },
	{ keys = "J", command = { "cycle", "secondary-sid" } },
	{ keys = "RIGHT", command = { "seek", 5 } },
	{ keys = "LEFT", command = { "seek", -5 } },
	{ keys = "UP", command = { "seek", 60 } },
	{ keys = "DOWN", command = { "seek", -60 } },
	{ keys = "H", command = { "sub-seek", -1 } },
	{ keys = "L", command = { "sub-seek", 1 } },
	{ keys = "q", command = { "quit" } },
	{ keys = "Ctrl+w", command = { "quit" } },
	{ keys = "MBTN_BACK", command = { "sub-seek", -1 } },
}

for _, expected in ipairs(expected_mpv_bindings) do
	local binding = find_binding(expected.keys)
	assert_true(binding ~= nil, "default mpv binding should register " .. expected.keys)
	binding.fn()
	local command = recorded.mpv_commands[#recorded.mpv_commands]
	assert_true(command ~= nil, "default mpv binding should invoke mpv command " .. expected.keys)
	for index, value in ipairs(expected.command) do
		assert_true(command[index] == value, "default mpv command mismatch for " .. expected.keys)
	end
end

local expected_cli_bindings = {
	{ keys = "Shift+]", flag = "--shift-sub-delay-next-line" },
	{ keys = "}", flag = "--shift-sub-delay-next-line" },
	{ keys = "Shift+[", flag = "--shift-sub-delay-prev-line" },
	{ keys = "{", flag = "--shift-sub-delay-prev-line" },
	{ keys = "Ctrl+Alt+c", flag = "--open-youtube-picker" },
	{ keys = "Ctrl+Alt+p", flag = "--open-playlist-browser" },
	{ keys = "Ctrl+H", flag = "--replay-current-subtitle" },
	{ keys = "Ctrl+L", flag = "--play-next-subtitle" },
	{ keys = "w", flag = "--mark-watched" },
}

local visible_overlay_toggle = find_binding("Alt+O")
assert_true(visible_overlay_toggle ~= nil, "visible overlay session binding should register")
visible_overlay_toggle.fn()
assert_true(recorded.overlay_toggles == 1, "visible overlay session binding should use plugin toggle")

for _, expected in ipairs(expected_cli_bindings) do
	local binding = find_binding(expected.keys)
	assert_true(binding ~= nil, "default session action should register " .. expected.keys)
	binding.fn()
	local cli_call = recorded.async_calls[#recorded.async_calls]
	assert_true(cli_call ~= nil, "default session action should invoke CLI " .. expected.keys)
	assert_true(cli_call[2] == expected.flag, "default session action should pass " .. expected.flag)
end

local play_next = find_binding("Ctrl+L")
assert_true(play_next ~= nil, "play-next subtitle binding should use mpv shifted-letter form")

local subtitle_jump = find_binding("L")
assert_true(subtitle_jump ~= nil, "shifted subtitle jump binding should use mpv uppercase letter form")

play_next.fn()
local play_next_call = recorded.async_calls[#recorded.async_calls]
assert_true(play_next_call ~= nil, "play-next binding should invoke CLI action")
assert_true(play_next_call[2] == "--play-next-subtitle", "play-next binding should pass CLI flag")

local character_dictionary_manager = find_binding("Ctrl+d")
assert_true(
	character_dictionary_manager ~= nil,
	"character dictionary manager binding should be registered"
)

character_dictionary_manager.fn()
local character_dictionary_manager_call = recorded.async_calls[#recorded.async_calls]
assert_true(
	character_dictionary_manager_call ~= nil,
	"character dictionary manager binding should invoke CLI action"
)
assert_true(
	character_dictionary_manager_call[2] == "--session-action",
	"character dictionary manager binding should use generic session action CLI flag"
)
assert_true(
	character_dictionary_manager_call[3] == '{"actionId":"openCharacterDictionaryManager"}',
	"character dictionary manager binding should pass generated session action payload"
)

local future_panel = find_binding("Ctrl+Alt+F12")
assert_true(future_panel ~= nil, "artifact CLI binding should be registered without plugin mapping")

future_panel.fn()
local future_panel_call = recorded.async_calls[#recorded.async_calls]
assert_true(future_panel_call ~= nil, "artifact CLI binding should invoke CLI action")
assert_true(
	future_panel_call[2] == "--session-action",
	"artifact CLI binding should pass generic session action CLI flag"
)
assert_true(
	future_panel_call[3] == '{"actionId":"openFuturePanel"}',
	"artifact CLI binding should pass generated session action payload"
)

starter.fn()

local call = recorded.async_calls[#recorded.async_calls]
assert_true(call ~= nil, "multi-line shortcut should invoke CLI action")
assert_true(call[1] == "/tmp/subminer", "CLI action should use configured binary")
assert_true(call[2] == "--mine-sentence-multiple", "CLI action should enter mine sentence count selector")
assert_true(call[3] == nil, "CLI action should not bind a plugin-side digit count")

print("plugin session binding regression tests: OK")
