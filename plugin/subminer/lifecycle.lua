local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local opts = ctx.opts
	local state = ctx.state
	local options_helper = ctx.options_helper
	local process = ctx.process
	local aniskip = ctx.aniskip
	local hover = ctx.hover
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd

	local function schedule_aniskip_fetch(trigger_source, delay_seconds)
		local delay = tonumber(delay_seconds) or 0
		mp.add_timeout(delay, function()
			aniskip.fetch_aniskip_for_current_media(trigger_source)
		end)
	end

	local function resolve_auto_start_enabled()
		local raw_auto_start = opts.auto_start
		if raw_auto_start == nil then
			raw_auto_start = opts.auto_start_overlay
		end
		if raw_auto_start == nil then
			raw_auto_start = opts["auto-start"]
		end
		return options_helper.coerce_bool(raw_auto_start, false)
	end

	local function on_file_loaded()
		aniskip.clear_aniskip_state()
		process.disarm_auto_play_ready_gate()

		local should_auto_start = resolve_auto_start_enabled()
		if should_auto_start then
			if not process.has_matching_mpv_ipc_socket(opts.socket_path) then
				subminer_log(
					"info",
					"lifecycle",
					"Skipping auto-start: input-ipc-server does not match configured socket_path"
				)
				schedule_aniskip_fetch("file-loaded", 0)
				return
			end

			process.start_overlay({
				auto_start_trigger = true,
				socket_path = opts.socket_path,
			})
			-- Give the overlay process a moment to initialize before querying AniSkip.
			schedule_aniskip_fetch("overlay-start", 0.8)
			return
		end

		schedule_aniskip_fetch("file-loaded", 0)
	end

	local function on_shutdown()
		aniskip.clear_aniskip_state()
		hover.clear_hover_overlay()
		process.disarm_auto_play_ready_gate()
		if state.overlay_running then
			subminer_log("info", "lifecycle", "mpv shutting down, hiding SubMiner overlay")
			process.hide_visible_overlay()
		end
	end

	local function register_lifecycle_hooks()
		mp.register_event("file-loaded", on_file_loaded)
		mp.register_event("shutdown", on_shutdown)
		mp.register_event("file-loaded", function()
			hover.clear_hover_overlay()
		end)
		mp.register_event("end-file", function()
			process.disarm_auto_play_ready_gate()
			hover.clear_hover_overlay()
			if state.overlay_running then
				process.hide_visible_overlay()
			end
		end)
		mp.register_event("shutdown", function()
			hover.clear_hover_overlay()
		end)
		mp.register_event("end-file", function()
			aniskip.clear_aniskip_state()
		end)
		mp.register_event("shutdown", function()
			aniskip.clear_aniskip_state()
		end)
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
