local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local opts = ctx.opts
	local state = ctx.state
	local environment = ctx.environment
	local binary = ctx.binary
	local options_helper = ctx.options_helper
	local process = ctx.process
	local aniskip = ctx.aniskip
	local hover = ctx.hover
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd

	local function on_file_loaded()
		if not environment.is_subminer_app_running() then
			aniskip.clear_aniskip_state()
			subminer_log("debug", "lifecycle", "Skipping file load hooks: SubMiner app not running")
			return true
		end

		aniskip.clear_aniskip_state()
		aniskip.fetch_aniskip_for_current_media()
		state.binary_path = binary.find_binary()
		if state.binary_path then
			state.binary_available = true
			subminer_log("info", "lifecycle", "SubMiner ready (binary: " .. state.binary_path .. ")")
			local should_auto_start = options_helper.coerce_bool(opts.auto_start, false)
			if should_auto_start then
				process.start_overlay()
			end
		else
			state.binary_available = false
			subminer_log("warn", "binary", "SubMiner binary not found - overlay features disabled")
			if opts.binary_path ~= "" then
				subminer_log("warn", "binary", "Configured path '" .. opts.binary_path .. "' does not exist")
			end
		end
	end

	local function on_shutdown()
		aniskip.clear_aniskip_state()
		hover.clear_hover_overlay()
		if (state.overlay_running or state.texthooker_running) and state.binary_available then
			subminer_log("info", "lifecycle", "mpv shutting down, stopping SubMiner process")
			show_osd("Shutting down...")
			process.stop_overlay()
		end
	end

	local function register_lifecycle_hooks()
		mp.register_event("file-loaded", on_file_loaded)
		mp.register_event("shutdown", on_shutdown)
		mp.register_event("file-loaded", hover.clear_hover_overlay)
		mp.register_event("end-file", hover.clear_hover_overlay)
		mp.register_event("shutdown", hover.clear_hover_overlay)
		mp.register_event("end-file", aniskip.clear_aniskip_state)
		mp.register_event("shutdown", aniskip.clear_aniskip_state)
		mp.add_hook("on_unload", 10, function()
			hover.clear_hover_overlay()
			aniskip.clear_aniskip_state()
		end)
		mp.observe_property("sub-start", "native", function()
			hover.clear_hover_overlay()
		end)
		mp.observe_property("time-pos", "number", function()
			aniskip.update_intro_button_visibility()
		end)
	end

	return {
		on_file_loaded = on_file_loaded,
		on_shutdown = on_shutdown,
		register_lifecycle_hooks = register_lifecycle_hooks,
	}
end

return M
