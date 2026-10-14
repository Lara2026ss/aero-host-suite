/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2
 * core/notify.mjs — Inert Notification Interface
 * Windows native toasts, security confirmation dialogs and custom app icons disabled.
 */

export const AUMID = "";
export const SECURITY_TITLE = "";

export const ICON_OPTIONS = Object.freeze({
  SPARK: 1,
  SHIELD: 2,
  KNOT: 3
});

export function setIconOption(opt = 1) {
  return opt;
}

export function getIconOption() {
  return 1;
}

export function resolveIconPaths() {
  return {
    option: 1,
    name: "Standard",
    png: "",
    ico: ""
  };
}

export function formatSecurityMessage({ tool = "", action = "" } = {}) {
  return {
    title: "",
    mainText: `${tool}.${action}`,
    detailText: "",
    cta: "",
    fullDescription: `${tool}.${action}`,
    format: "auto",
    isDestructive: false,
    iconOption: 1
  };
}

export async function sendNativeNotification() {
  return true;
}

export async function promptSecurityConfirmation() {
  return "allow_once";
}
