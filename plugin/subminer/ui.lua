local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local input = ctx.input
	local opts = ctx.opts
	local state = ctx.state
	local process = ctx.process
	local aniskip = ctx.aniskip
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd

	local function show_menu()
		if not state.binary_available then
			subminer_log("error", "binary", "SubMiner binary not found")
			show_osd("Error: binary not found")
			return
		end

		local items = {
			"Start overlay",
			"Stop overlay",
			"Toggle overlay",
			"Toggle invisible overlay",
			"Open options",
			"Restart overlay",
			"Check status",
		}

		local actions = {
			process.start_overlay,
			process.stop_overlay,
			process.toggle_overlay,
			process.toggle_invisible_overlay,
			process.open_options,
			process.restart_overlay,
			process.check_status,
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
		mp.add_key_binding("y-s", "subminer-start", process.start_overlay)
		mp.add_key_binding("y-S", "subminer-stop", process.stop_overlay)
		mp.add_key_binding("y-t", "subminer-toggle", process.toggle_overlay)
		mp.add_key_binding("y-i", "subminer-toggle-invisible", process.toggle_invisible_overlay)
		mp.add_key_binding("y-I", "subminer-show-invisible", process.show_invisible_overlay)
		mp.add_key_binding("y-u", "subminer-hide-invisible", process.hide_invisible_overlay)
		mp.add_key_binding("y-y", "subminer-menu", show_menu)
		mp.add_key_binding("y-o", "subminer-options", process.open_options)
		mp.add_key_binding("y-r", "subminer-restart", process.restart_overlay)
		mp.add_key_binding("y-c", "subminer-status", process.check_status)
		if type(opts.aniskip_button_key) == "string" and opts.aniskip_button_key ~= "" then
			mp.add_key_binding(opts.aniskip_button_key, "subminer-skip-intro", aniskip.skip_intro_now)
		end
		if opts.aniskip_button_key ~= "y-k" then
			mp.add_key_binding("y-k", "subminer-skip-intro-fallback", aniskip.skip_intro_now)
		end
	end

	return {
		show_menu = show_menu,
		register_keybindings = register_keybindings,
	}
end

return M
