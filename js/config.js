/*
 * Deployment config for vTiger CRM Auto Update.
 *
 * userApiUrl: the Web app URL of the Apps Script in apps-script/Code.gs
 * (Apps Script → Deploy → Manage deployments → Web app URL, ends in /exec).
 * Sign-in and the Users list both go through it — see README.md →
 * "User database (Google Sheet)". This URL is not a secret.
 */
window.CRM_AUTO_UPDATE_CONFIG = {
  userApiUrl: "https://script.google.com/macros/s/AKfycbyYpAvLzoU5ggHDhKU1jiAN7HSlQn5_3EvtrLgd-qkG2mmfvrcyUTMsWAtua8VF_DXcbA/exec",
};
