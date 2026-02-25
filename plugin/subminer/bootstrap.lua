local M = {}

function M.init()
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

	ctx.log = require("log").create(ctx)
	ctx.binary = require("binary").create(ctx)
	ctx.aniskip = require("aniskip").create(ctx)
	ctx.hover = require("hover").create(ctx)
	ctx.process = require("process").create(ctx)
	ctx.ui = require("ui").create(ctx)
	ctx.messages = require("messages").create(ctx)
	ctx.lifecycle = require("lifecycle").create(ctx)

	ctx.ui.register_keybindings()
	ctx.messages.register_script_messages()
	ctx.lifecycle.register_lifecycle_hooks()
	ctx.log.subminer_log("info", "lifecycle", "SubMiner plugin loaded")
end

return M
