import {
  MPV_REQUEST_ID_AID,
  MPV_REQUEST_ID_OSD_DIMENSIONS,
  MPV_REQUEST_ID_OSD_HEIGHT,
  MPV_REQUEST_ID_PATH,
  MPV_REQUEST_ID_SECONDARY_SUBTEXT,
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
  MPV_REQUEST_ID_SUB_ASS_OVERRIDE,
  MPV_REQUEST_ID_SUB_BOLD,
  MPV_REQUEST_ID_SUB_BORDER_SIZE,
  MPV_REQUEST_ID_SUB_FONT,
  MPV_REQUEST_ID_SUB_FONT_SIZE,
  MPV_REQUEST_ID_SUB_ITALIC,
  MPV_REQUEST_ID_SUB_MARGIN_X,
  MPV_REQUEST_ID_SUB_MARGIN_Y,
  MPV_REQUEST_ID_SUB_POS,
  MPV_REQUEST_ID_SUB_SCALE,
  MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW,
  MPV_REQUEST_ID_SUB_SHADOW_OFFSET,
  MPV_REQUEST_ID_SUB_SPACING,
  MPV_REQUEST_ID_SUBTEXT,
  MPV_REQUEST_ID_SUBTEXT_ASS,
  MPV_REQUEST_ID_SUB_USE_MARGINS,
} from "./mpv-protocol";

type MpvProtocolCommand = {
  command: unknown[];
  request_id?: number;
};

export interface MpvSendCommand {
  (command: MpvProtocolCommand): boolean;
}

const MPV_SUBTITLE_PROPERTY_OBSERVATIONS: string[] = [
  "sub-text",
  "path",
  "sub-start",
  "sub-end",
  "time-pos",
  "secondary-sub-text",
  "aid",
  "sub-pos",
  "sub-font-size",
  "sub-scale",
  "sub-margin-y",
  "sub-margin-x",
  "sub-font",
  "sub-spacing",
  "sub-bold",
  "sub-italic",
  "sub-scale-by-window",
  "osd-height",
  "osd-dimensions",
  "sub-text-ass",
  "sub-border-size",
  "sub-shadow-offset",
  "sub-ass-override",
  "sub-use-margins",
  "media-title",
];

const MPV_INITIAL_PROPERTY_REQUESTS: Array<MpvProtocolCommand> = [
  {
    command: ["get_property", "sub-text"],
    request_id: MPV_REQUEST_ID_SUBTEXT,
  },
  {
    command: ["get_property", "sub-text-ass"],
    request_id: MPV_REQUEST_ID_SUBTEXT_ASS,
  },
  {
    command: ["get_property", "path"],
    request_id: MPV_REQUEST_ID_PATH,
  },
  {
    command: ["get_property", "media-title"],
  },
  {
    command: ["get_property", "secondary-sub-text"],
    request_id: MPV_REQUEST_ID_SECONDARY_SUBTEXT,
  },
  {
    command: ["get_property", "secondary-sub-visibility"],
    request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
  },
  {
    command: ["get_property", "aid"],
    request_id: MPV_REQUEST_ID_AID,
  },
  {
    command: ["get_property", "sub-pos"],
    request_id: MPV_REQUEST_ID_SUB_POS,
  },
  {
    command: ["get_property", "sub-font-size"],
    request_id: MPV_REQUEST_ID_SUB_FONT_SIZE,
  },
  {
    command: ["get_property", "sub-scale"],
    request_id: MPV_REQUEST_ID_SUB_SCALE,
  },
  {
    command: ["get_property", "sub-margin-y"],
    request_id: MPV_REQUEST_ID_SUB_MARGIN_Y,
  },
  {
    command: ["get_property", "sub-margin-x"],
    request_id: MPV_REQUEST_ID_SUB_MARGIN_X,
  },
  {
    command: ["get_property", "sub-font"],
    request_id: MPV_REQUEST_ID_SUB_FONT,
  },
  {
    command: ["get_property", "sub-spacing"],
    request_id: MPV_REQUEST_ID_SUB_SPACING,
  },
  {
    command: ["get_property", "sub-bold"],
    request_id: MPV_REQUEST_ID_SUB_BOLD,
  },
  {
    command: ["get_property", "sub-italic"],
    request_id: MPV_REQUEST_ID_SUB_ITALIC,
  },
  {
    command: ["get_property", "sub-scale-by-window"],
    request_id: MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW,
  },
  {
    command: ["get_property", "osd-height"],
    request_id: MPV_REQUEST_ID_OSD_HEIGHT,
  },
  {
    command: ["get_property", "osd-dimensions"],
    request_id: MPV_REQUEST_ID_OSD_DIMENSIONS,
  },
  {
    command: ["get_property", "sub-border-size"],
    request_id: MPV_REQUEST_ID_SUB_BORDER_SIZE,
  },
  {
    command: ["get_property", "sub-shadow-offset"],
    request_id: MPV_REQUEST_ID_SUB_SHADOW_OFFSET,
  },
  {
    command: ["get_property", "sub-ass-override"],
    request_id: MPV_REQUEST_ID_SUB_ASS_OVERRIDE,
  },
  {
    command: ["get_property", "sub-use-margins"],
    request_id: MPV_REQUEST_ID_SUB_USE_MARGINS,
  },
];

export function subscribeToMpvProperties(send: MpvSendCommand): void {
  MPV_SUBTITLE_PROPERTY_OBSERVATIONS.forEach((property, index) => {
    send({ command: ["observe_property", index + 1, property] });
  });
}

export function requestMpvInitialState(send: MpvSendCommand): void {
  MPV_INITIAL_PROPERTY_REQUESTS.forEach((payload) => {
    send(payload);
  });
}
