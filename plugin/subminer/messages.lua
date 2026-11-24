local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local opts = ctx.opts
	local process = ctx.process
	local ui = ctx.ui
	local state = ctx.state

	local function register_script_messages()
		mp.register_script_message("subminer-start", function(...)
			process.start_overlay_from_script_message(...)
		end)
		mp.register_script_message("subminer-stop", function()
			process.stop_overlay()
		end)
		mp.register_script_message("subminer-toggle", function()
			process.toggle_overlay()
		end)
		mp.register_script_message("subminer-visible-overlay-hidden", function()
			process.record_visible_overlay_visibility(false)
		end)
		mp.register_script_message("subminer-visible-overlay-shown", function()
			process.record_visible_overlay_visibility(true)
		end)
		mp.register_script_message("subminer-managed-subtitles-loading", function()
			state.skip_managed_subtitle_rearm_once = true
			state.app_managed_playback_pending = true
		end)
		-- Playlist entries the app queued (Jellyfin neighbor episodes). mpv starts them on its own,
		-- so start-file recognizes them by path instead of waiting for subminer-managed-subtitles-loading.
		mp.register_script_message("subminer-managed-playlist-entries", function(...)
			local paths = {}
			for _, path in ipairs({ ... }) do
				paths[path] = true
			end
			state.app_managed_playlist_paths = paths
			-- Options belong to queued entries; drop those the app no longer queues.
			for path in pairs(state.app_playlist_entry_options) do
				if not paths[path] then
					state.app_playlist_entry_options[path] = nil
				end
			end
		end)
		-- `<path> <name> <value> ...`: options for an entry the app loads with loadlist, where mpv
		-- (before 0.42) cannot take both a playlist title and per-file options.
		mp.register_script_message("subminer-playlist-entry-options", function(path, ...)
			if type(path) ~= "string" or path == "" then
				return
			end
			local args = { ... }
			local options = {}
			for index = 1, #args - 1, 2 do
				options[#options + 1] = { args[index], args[index + 1] }
			end
			state.app_playlist_entry_options[path] = options
		end)
		mp.register_script_message("subminer-menu", function()
			ui.show_menu()
		end)
		mp.register_script_message("subminer-options", function()
			process.open_options()
		end)
		mp.register_script_message("subminer-restart", function()
			process.restart_overlay()
		end)
		mp.register_script_message("subminer-status", function()
			process.check_status()
		end)
		mp.register_script_message("subminer-autoplay-ready", function()
			process.notify_auto_play_ready()
		end)
		mp.register_script_message("subminer-autoplay-hold", function()
			process.hold_auto_play_ready_gate()
		end)
		mp.register_script_message("subminer-overlay-loading-ready", function()
			process.mark_overlay_content_ready()
		end)
		-- Sent when the app dismisses loading while its overlay window has no content (hidden mid-load,
		-- or a replaced window).
		mp.register_script_message("subminer-overlay-loading-dismissed", function()
			process.mark_overlay_content_not_ready()
		end)
		mp.register_script_message("subminer-stats-toggle", function()
			if opts.osd_messages then
				mp.osd_message("Stats: press ` (backtick) in overlay", 3)
			end
		end)
		mp.register_script_message("subminer-reload-session-bindings", function()
			ctx.session_bindings.reload_bindings()
		end)
	end

	return {
		register_script_messages = register_script_messages,
	}
end

return M
