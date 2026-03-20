local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local process = ctx.process
	local aniskip = ctx.aniskip
	local hover = ctx.hover
	local ui = ctx.ui

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
		mp.register_script_message("subminer-aniskip-refresh", function()
			aniskip.fetch_aniskip_for_current_media("script-message")
		end)
		mp.register_script_message("subminer-skip-intro", function()
			aniskip.skip_intro_now()
		end)
		mp.register_script_message(hover.HOVER_MESSAGE_NAME, function(payload_json)
			hover.handle_hover_message(payload_json)
		end)
		mp.register_script_message(hover.HOVER_MESSAGE_NAME_LEGACY, function(payload_json)
			hover.handle_hover_message(payload_json)
		end)
		mp.register_script_message("subminer-stats-toggle", function()
			mp.osd_message("Stats: press ` (backtick) in overlay", 3)
		end)
	end

	return {
		register_script_messages = register_script_messages,
	}
end

return M
