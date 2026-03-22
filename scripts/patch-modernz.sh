#!/bin/bash

set -euo pipefail

TARGET="${HOME}/.config/mpv/scripts/modernz.lua"

usage() {
    cat <<'EOF'
Usage: patch-modernz.sh [--target /path/to/modernz.lua]

Applies the local ModernZ OSC sidebar-resize patch to an existing modernz.lua.
If the target file does not exist, the script exits without changing anything.
EOF
}

while [[ $# -gt 0 ]]; do
    case "$1" in
        --target)
            if [[ $# -lt 2 || -z "${2:-}" || "$2" == -* ]]; then
                echo "patch-modernz: --target requires a non-empty file path" >&2
                usage >&2
                exit 1
            fi
            TARGET="$2"
            shift 2
            ;;
        --help|-h)
            usage
            exit 0
            ;;
        *)
            echo "Unknown argument: $1" >&2
            exit 1
            ;;
    esac
done

if [[ ! -f "$TARGET" ]]; then
    echo "patch-modernz: target missing, skipped: $TARGET"
    exit 0
fi

if grep -q 'get_external_video_margin_ratio' "$TARGET" \
    && grep -q 'observe_cached("video-margin-ratio-right"' "$TARGET"; then
    echo "patch-modernz: already patched: $TARGET"
    exit 0
fi

if ! patch --forward --quiet "$TARGET" <<'PATCH'
--- a/modernz.lua
+++ b/modernz.lua
@@ -931,6 +931,26 @@ local function reset_margins()
     set_margin_offset("osd-margin-y", 0)
 end
 
+local function get_external_video_margin_ratio(prop)
+    local value = mp.get_property_number(prop, 0) or 0
+    if value < 0 then return 0 end
+    if value > 0.95 then return 0.95 end
+    return value
+end
+
+local function get_layout_horizontal_bounds()
+    local margin_l = get_external_video_margin_ratio("video-margin-ratio-left")
+    local margin_r = get_external_video_margin_ratio("video-margin-ratio-right")
+    local width_ratio = math.max(0.05, 1 - margin_l - margin_r)
+    local pos_x = osc_param.playresx * margin_l
+    local width = osc_param.playresx * width_ratio
+
+    osc_param.video_margins.l = margin_l
+    osc_param.video_margins.r = margin_r
+
+    return pos_x, width
+end
+
 local function update_margins()
     local use_margins = get_hidetimeout() < 0 or user_opts.dynamic_margins
     local top_vis    = state.wc_visible
@@ -1965,8 +1985,9 @@ layouts["modern"] = function ()
     local chapter_index = user_opts.show_chapter_title and mp.get_property_number("chapter", -1) >= 0
     local osc_height_offset = (no_title and user_opts.notitle_osc_h_offset or 0) + ((no_chapter or not chapter_index) and user_opts.nochapter_osc_h_offset or 0)
 
+    local posX, layout_width = get_layout_horizontal_bounds()
     local osc_geo = {
-        w = osc_param.playresx,
+        w = layout_width,
         h = user_opts.osc_height - osc_height_offset
     }
 
@@ -1974,7 +1995,6 @@ layouts["modern"] = function ()
     osc_param.video_margins.b = math.max(user_opts.osc_height, user_opts.fade_alpha) / osc_param.playresy
 
     -- origin of the controllers, left/bottom corner
-    local posX = 0
     local posY = osc_param.playresy
 
     osc_param.areas = {} -- delete areas
@@ -2191,8 +2211,9 @@ layouts["modern-compact"] = function ()
         ((user_opts.title_mbtn_left_command == "" and user_opts.title_mbtn_right_command == "") and 25 or 0) +
         (((user_opts.chapter_title_mbtn_left_command == "" and user_opts.chapter_title_mbtn_right_command == "") or not chapter_index) and 10 or 0)
 
+    local posX, layout_width = get_layout_horizontal_bounds()
     local osc_geo = {
-        w = osc_param.playresx,
+        w = layout_width,
         h = 145 - osc_height_offset
     }
 
@@ -2200,7 +2221,6 @@ layouts["modern-compact"] = function ()
     osc_param.video_margins.b = math.max(osc_geo.h, user_opts.fade_alpha) / osc_param.playresy
 
     -- origin of the controllers, left/bottom corner
-    local posX = 0
     local posY = osc_param.playresy
 
     osc_param.areas = {} -- delete areas
@@ -2370,8 +2390,9 @@ layouts["modern-compact"] = function ()
 end
 
 layouts["modern-image"] = function ()
+    local posX, layout_width = get_layout_horizontal_bounds()
     local osc_geo = {
-        w = osc_param.playresx,
+        w = layout_width,
         h = 50
     }
 
@@ -2379,7 +2400,6 @@ layouts["modern-image"] = function ()
     osc_param.video_margins.b = math.max(50, user_opts.fade_alpha) / osc_param.playresy
 
     -- origin of the controllers, left/bottom corner
-    local posX = 0
     local posY = osc_param.playresy
 
     osc_param.areas = {} -- delete areas
@@ -3718,6 +3738,14 @@ observe_cached("border", request_init_resize)
 observe_cached("title-bar", request_init_resize)
 observe_cached("window-maximized", request_init_resize)
 observe_cached("idle-active", request_tick)
+observe_cached("video-margin-ratio-left", function ()
+    state.marginsREQ = true
+    request_init_resize()
+end)
+observe_cached("video-margin-ratio-right", function ()
+    state.marginsREQ = true
+    request_init_resize()
+end)
 mp.observe_property("user-data/mpv/console/open", "bool", function(_, val)
     if val and user_opts.visibility == "auto" and not user_opts.showonselect then
         osc_visible(false)
PATCH
then
    echo "patch-modernz: failed to apply patch to $TARGET" >&2
    exit 1
fi

echo "patch-modernz: patched $TARGET"
