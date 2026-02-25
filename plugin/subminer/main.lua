local mp = require("mp")

local script_dir = mp.get_script_directory() or "."
local module_patterns = script_dir .. "/?.lua;" .. script_dir .. "/?/init.lua;"
if not package.path:find(module_patterns, 1, true) then
	package.path = module_patterns .. package.path
end

require("bootstrap").init()
