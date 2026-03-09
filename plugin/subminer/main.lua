local mp = require("mp")

local function current_script_dir()
	if type(mp.get_script_directory) == "function" then
		local from_mpv = mp.get_script_directory()
		if type(from_mpv) == "string" and from_mpv ~= "" then
			return from_mpv
		end
	end

	local source = debug.getinfo(1, "S").source or ""
	if source:sub(1, 1) == "@" then
		local full = source:sub(2)
		return full:match("^(.*)[/\\][^/\\]+$") or "."
	end
	return "."
end

local script_dir = current_script_dir()
local module_patterns = script_dir .. "/?.lua;" .. script_dir .. "/?/init.lua;"
if not package.path:find(module_patterns, 1, true) then
	package.path = module_patterns .. package.path
end

local init_module = assert(loadfile(script_dir .. "/init.lua"))()
if type(init_module) == "table" and type(init_module.init) == "function" then
	init_module.init()
elseif type(init_module) == "function" then
	init_module()
end
