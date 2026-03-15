local M = {}
local DEFAULT_ANISKIP_BUTTON_KEY = "TAB"
local LEGACY_ANISKIP_BUTTON_KEY = "y-k"

function M.create(ctx)
	local mp = ctx.mp
	local input = ctx.input
	local opts = ctx.opts
	local process = ctx.process
	local aniskip = ctx.aniskip
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd

	local function ensure_binary_for_menu()
		if process.check_binary_available() then
			return true
		end
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return false
	end

	local function show_menu()
		if not ensure_binary_for_menu() then
			return
		end

		local items = {
			"Start overlay",
			"Stop overlay",
			"Toggle overlay",
			"Open options",
			"Restart overlay",
			"Check status",
			"Stats",
		}

		local actions = {
			function()
				process.start_overlay()
			end,
			function()
				process.stop_overlay()
			end,
			function()
				process.toggle_overlay()
			end,
			function()
				process.open_options()
			end,
			function()
				process.restart_overlay()
			end,
			function()
				process.check_status()
			end,
			function()
				mp.commandv("script-message", "subminer-stats-toggle")
			end,
		}

		input.select({
			prompt = "SubMiner: ",
			items = items,
			submit = function(index)
				if index and actions[index] then
					actions[index]()
				end
			end,
		})
	end

	local function register_keybindings()
		mp.add_key_binding("y-s", "subminer-start", function()
			process.start_overlay()
		end)
		mp.add_key_binding("y-S", "subminer-stop", function()
			process.stop_overlay()
		end)
		mp.add_key_binding("y-t", "subminer-toggle", function()
			process.toggle_overlay()
		end)
		mp.add_key_binding("y-y", "subminer-menu", show_menu)
		mp.add_key_binding("y-o", "subminer-options", function()
			process.open_options()
		end)
		mp.add_key_binding("y-r", "subminer-restart", function()
			process.restart_overlay()
		end)
		mp.add_key_binding("y-c", "subminer-status", function()
			process.check_status()
		end)
		if type(opts.aniskip_button_key) == "string" and opts.aniskip_button_key ~= "" then
			mp.add_key_binding(opts.aniskip_button_key, "subminer-skip-intro", function()
				aniskip.skip_intro_now()
			end)
		end
		if
			opts.aniskip_button_key ~= LEGACY_ANISKIP_BUTTON_KEY
			and opts.aniskip_button_key ~= DEFAULT_ANISKIP_BUTTON_KEY
		then
			mp.add_key_binding(LEGACY_ANISKIP_BUTTON_KEY, "subminer-skip-intro-fallback", function()
				aniskip.skip_intro_now()
			end)
		end
	end

	return {
		show_menu = show_menu,
		register_keybindings = register_keybindings,
	}
end

return M
