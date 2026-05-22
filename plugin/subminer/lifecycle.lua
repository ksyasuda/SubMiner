local M = {}

local AUTO_START_SOCKET_RETRY_DELAY_SECONDS = 0.2
local AUTO_START_SOCKET_RETRY_MAX_ATTEMPTS = 25

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

	local function resolve_media_identity()
		local path = mp.get_property("path")
		if type(path) == "string" and path ~= "" then
			return path
		end

		local filename = mp.get_property("filename")
		if type(filename) == "string" and filename ~= "" then
			return filename
		end

		local media_title = mp.get_property("media-title")
		if type(media_title) == "string" and media_title ~= "" then
			return media_title
		end

		return nil
	end

	local function is_reload_end_file(reason)
		return reason == "reload" or reason == "redirect"
	end

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

	local function next_auto_start_retry_generation()
		state.auto_start_retry_generation = (state.auto_start_retry_generation or 0) + 1
		return state.auto_start_retry_generation
	end

	local function has_matching_subminer_socket()
		if not process.has_matching_mpv_ipc_socket(opts.socket_path) then
			return false
		end
		return true
	end

	local function rearm_managed_subtitle_load_defaults()
		if not has_matching_subminer_socket() then
			return false
		end
		mp.set_property_native("sub-auto", "fuzzy")
		mp.set_property_native("sid", "auto")
		mp.set_property_native("secondary-sid", "auto")
		return true
	end

	local function refresh_managed_subtitle_autoloading()
		if not has_matching_subminer_socket() then
			return false
		end
		mp.set_property_native("sub-auto", "fuzzy")
		return true
	end

	local function start_overlay_when_socket_ready(generation, media_identity, same_media_loaded, attempt)
		if generation ~= state.auto_start_retry_generation then
			return
		end
		if media_identity ~= nil and state.current_media_identity ~= media_identity then
			return
		end
		if not resolve_auto_start_enabled() then
			schedule_aniskip_fetch("file-loaded", 0)
			return
		end

		local has_matching_socket = refresh_managed_subtitle_autoloading()
		if not has_matching_socket then
			if attempt < AUTO_START_SOCKET_RETRY_MAX_ATTEMPTS then
				mp.add_timeout(AUTO_START_SOCKET_RETRY_DELAY_SECONDS, function()
					start_overlay_when_socket_ready(generation, media_identity, same_media_loaded, attempt + 1)
				end)
				return
			end
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
			rearm_pause_until_ready = not same_media_loaded,
		})
		-- Give the overlay process a moment to initialize before querying AniSkip.
		schedule_aniskip_fetch("overlay-start", 0.8)
	end

	local function on_start_file()
		if state.pending_reload_media_identity ~= nil then
			return
		end
		rearm_managed_subtitle_load_defaults()
	end

	local function on_file_loaded()
		local media_identity = resolve_media_identity()
		local retry_generation = next_auto_start_retry_generation()
		local previous_media_identity = state.current_media_identity
		local same_media_reload = (
			media_identity ~= nil
			and state.pending_reload_media_identity ~= nil
			and media_identity == state.pending_reload_media_identity
		)
		local same_media_loaded = (
			media_identity ~= nil
			and previous_media_identity ~= nil
			and media_identity == previous_media_identity
		)
		local new_media_loaded = media_identity ~= nil and not same_media_reload and not same_media_loaded
		state.pending_reload_media_identity = nil
		state.current_media_identity = media_identity
		if new_media_loaded then
			state.suppress_ready_overlay_restore = false
		end

		if same_media_reload then
			subminer_log("debug", "lifecycle", "Skipping startup lifecycle for same-media mpv reload")
			if
				state.overlay_running
				and not state.suppress_ready_overlay_restore
				and resolve_auto_start_enabled()
				and process.has_matching_mpv_ipc_socket(opts.socket_path)
			then
				process.run_control_command_async("show-visible-overlay", {
					socket_path = opts.socket_path,
				})
			end
			return
		end

		local should_auto_start = resolve_auto_start_enabled()
		local has_matching_socket = process.has_matching_mpv_ipc_socket(opts.socket_path)
		local preserve_active_auto_start_gate = (
			state.overlay_running and state.auto_play_ready_gate_armed and should_auto_start and has_matching_socket
		)
		aniskip.clear_aniskip_state()
		if not preserve_active_auto_start_gate then
			process.disarm_auto_play_ready_gate()
		end

		if should_auto_start then
			start_overlay_when_socket_ready(retry_generation, media_identity, same_media_loaded, 1)
			return
		end

		refresh_managed_subtitle_autoloading()
		schedule_aniskip_fetch("file-loaded", 0)
	end

	local function on_shutdown()
		next_auto_start_retry_generation()
		aniskip.clear_aniskip_state()
		hover.clear_hover_overlay()
		process.disarm_auto_play_ready_gate()
		state.current_media_identity = nil
		state.pending_reload_media_identity = nil
	end

	local function register_lifecycle_hooks()
		mp.register_event("start-file", on_start_file)
		mp.register_event("file-loaded", on_file_loaded)
		mp.register_event("shutdown", on_shutdown)
		mp.register_event("file-loaded", function()
			hover.clear_hover_overlay()
		end)
		mp.register_event("end-file", function(event)
			process.disarm_auto_play_ready_gate()
			hover.clear_hover_overlay()
			local reason = type(event) == "table" and event.reason or nil
			if is_reload_end_file(reason) then
				state.pending_reload_media_identity = state.current_media_identity or resolve_media_identity()
				return
			end
			next_auto_start_retry_generation()
			state.current_media_identity = nil
			state.pending_reload_media_identity = nil
			if state.overlay_running and reason ~= "quit" then
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
