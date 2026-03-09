local M = {}
local BOOTSTRAP_GUARD_KEY = "__subminer_plugin_bootstrapped"

function M.init()
	if rawget(_G, BOOTSTRAP_GUARD_KEY) == true then
		return
	end
	rawset(_G, BOOTSTRAP_GUARD_KEY, true)

	local input = require("mp.input")
	local mp = require("mp")
	local msg = require("mp.msg")
	local options_lib = require("mp.options")
	local utils = require("mp.utils")

	local options_helper = require("options")
	local environment = require("environment").create({ mp = mp })
	local opts = options_helper.load(options_lib, environment.default_socket_path())
	local state = require("state").new()

	local ctx = {
		input = input,
		mp = mp,
		msg = msg,
		utils = utils,
		opts = opts,
		state = state,
		options_helper = options_helper,
		environment = environment,
	}

	local instances = {}

	local function lazy_instance(key, factory)
		if instances[key] == nil then
			instances[key] = factory()
		end
		return instances[key]
	end

	local function make_lazy_proxy(key, factory)
		return setmetatable({}, {
			__index = function(_, member)
				return lazy_instance(key, factory)[member]
			end,
		})
	end

	ctx.log = make_lazy_proxy("log", function()
		return require("log").create(ctx)
	end)
	ctx.binary = make_lazy_proxy("binary", function()
		return require("binary").create(ctx)
	end)
	ctx.aniskip = make_lazy_proxy("aniskip", function()
		return require("aniskip").create(ctx)
	end)
	ctx.hover = make_lazy_proxy("hover", function()
		return require("hover").create(ctx)
	end)
	ctx.process = make_lazy_proxy("process", function()
		return require("process").create(ctx)
	end)
	ctx.ui = make_lazy_proxy("ui", function()
		return require("ui").create(ctx)
	end)
	ctx.messages = make_lazy_proxy("messages", function()
		return require("messages").create(ctx)
	end)
	ctx.lifecycle = make_lazy_proxy("lifecycle", function()
		return require("lifecycle").create(ctx)
	end)

	ctx.ui.register_keybindings()
	ctx.messages.register_script_messages()
	ctx.lifecycle.register_lifecycle_hooks()
	ctx.log.subminer_log("info", "lifecycle", "SubMiner plugin loaded")
end

return M
