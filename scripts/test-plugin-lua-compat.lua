local MODULE_PATHS = {
	"plugin/subminer/hover.lua",
	"plugin/subminer/environment.lua",
	"plugin/subminer/version.lua",
}

local LEGACY_PARSER_CANDIDATES = {
	"luajit",
	"lua5.1",
	"lua51",
}

local function assert_true(condition, message)
	if condition then
		return
	end
	error(message or "assert_true failed")
end

local function read_file(path)
	local file = assert(io.open(path, "r"), "failed to open " .. path)
	local content = file:read("*a")
	file:close()
	return content
end

local function find_legacy_incompatible_continue(source)
	local goto_start, goto_end = source:find("%f[%a]goto%s+continue%f[%A]")
	if goto_start then
		return "goto continue", goto_start, goto_end
	end

	local label_start, label_end = source:find("::continue::", 1, true)
	if label_start then
		return "::continue::", label_start, label_end
	end

	return nil
end

local function assert_no_legacy_incompatible_continue(path)
	local source = read_file(path)
	local match = find_legacy_incompatible_continue(source)
	assert_true(match == nil, path .. " still contains legacy-incompatible continue control flow: " .. tostring(match))
end

local function assert_loadfile_ok(path)
	local chunk, err = loadfile(path)
	assert_true(chunk ~= nil, "loadfile failed for " .. path .. ": " .. tostring(err))
end

local function normalize_execute_result(ok, why, code)
	if type(ok) == "number" then
		return ok == 0, ok
	end
	if type(ok) == "boolean" then
		if ok then
			return true, code or 0
		end
		return false, code or 1
	end
	return false, code or 1
end

local function command_succeeds(command)
	local ok, why, code = os.execute(command)
	return normalize_execute_result(ok, why, code)
end

local function command_exists(command)
	local shell = package.config:sub(1, 1) == "\\" and "where " or "command -v "
	local redirect = package.config:sub(1, 1) == "\\" and " >NUL 2>NUL" or " >/dev/null 2>&1"
	local escaped = command
	local success = command_succeeds(shell .. escaped .. redirect)
	return success
end

local function find_legacy_parser()
	for _, command in ipairs(LEGACY_PARSER_CANDIDATES) do
		if command_exists(command) then
			return command
		end
	end
	return nil
end

local function shell_redirect()
	if package.config:sub(1, 1) == "\\" then
		return " >NUL 2>NUL"
	end
	return " >/dev/null 2>&1"
end

local function assert_parser_accepts_file(parser, path)
	local command = string.format("%s -e %q%s", parser, "assert(loadfile(" .. string.format("%q", path) .. "))", shell_redirect())
	local success = command_succeeds(command)
	assert_true(success, parser .. " failed to parse " .. path)
end

local function assert_parser_rejects_legacy_fixture(parser)
	local legacy_fixture = [[
local tokens = {}
for _, token in ipairs(tokens or {}) do
	if type(token) ~= "table" then
		goto continue
	end
	::continue::
end
]]
	local command = string.format("%s -e %q%s", parser, legacy_fixture, shell_redirect())
	local success = command_succeeds(command)
	assert_true(not success, parser .. " unexpectedly accepted legacy goto/label continue fixture")
end

do
	local legacy_fixture = [[
for _, token in ipairs(tokens or {}) do
	if type(token) ~= "table" then
		goto continue
	end
	::continue::
end
]]
	local match = find_legacy_incompatible_continue(legacy_fixture)
	assert_true(match ~= nil, "legacy fixture should trigger incompatible continue detector")
end

for _, path in ipairs(MODULE_PATHS) do
	assert_no_legacy_incompatible_continue(path)
	assert_loadfile_ok(path)
end

local parser = find_legacy_parser()
if parser then
	assert_parser_rejects_legacy_fixture(parser)
	for _, path in ipairs(MODULE_PATHS) do
		assert_parser_accepts_file(parser, path)
	end
	print("plugin lua compatibility regression tests: OK (" .. parser .. ")")
else
	print("plugin lua compatibility regression tests: OK (legacy parser unavailable; structural checks only)")
end
