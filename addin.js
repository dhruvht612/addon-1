/**
 * Device Status Checker — Geotab Add-In
 * --------------------------------------
 * Lets the user type a device serial number, then:
 *   1. Searches MyGeotab for a Device with that serial number.
 *   2. Fetches that device's DeviceStatusInfo (last ping time, driver, GPS).
 *   3. Renders a result card with an Online / Idle / Offline badge.
 *
 * MyGeotab looks for the add-in under the global `geotab.addin` object,
 * so everything is registered there (see the bottom of this file).
 */

"use strict";

// Make sure the geotab.addin namespace exists. Inside MyGeotab it always does,
// but this guard lets you open index.html directly in a browser without errors.
window.geotab = window.geotab || {};
geotab.addin = geotab.addin || {};

geotab.addin.deviceStatusChecker = function () {

    // The authenticated MyGeotab API object. MyGeotab hands it to us in
    // initialize() and focus(); every server call goes through it.
    let api;

    // Cached references to the DOM elements we update. Filled in once during
    // initialize() so we don't query the document on every search.
    let el = {};

    /**
     * Look up every element the add-in touches, by id, and store it in `el`.
     * Ids are prefixed with "dsc-" to avoid clashing with MyGeotab's own DOM.
     */
    function cacheElements() {
        el.input = document.getElementById("dsc-serial-input");
        el.button = document.getElementById("dsc-check-button");
        el.buttonSpinner = document.getElementById("dsc-button-spinner");
        el.buttonLabel = document.getElementById("dsc-button-label");
        el.hint = document.getElementById("dsc-hint");
        el.message = document.getElementById("dsc-message");
        el.card = document.getElementById("dsc-result");
        el.badge = document.getElementById("dsc-badge");
        el.vehicleName = document.getElementById("dsc-vehicle-name");
        el.serialNumber = document.getElementById("dsc-serial-number");
        el.driverName = document.getElementById("dsc-driver-name");
        el.lastPing = document.getElementById("dsc-last-ping");
        el.gps = document.getElementById("dsc-gps");
    }

    // ---------------------------------------------------------------------
    // UI state helpers
    // Each function below flips the page into one of the add-in's states:
    // loading, message (not found / error), or showing the result card.
    // ---------------------------------------------------------------------

    /**
     * Toggle the loading state. While loading, the button is disabled and
     * shows a spinner with "Checking..." so the user knows a call is running.
     */
    function setLoading(isLoading) {
        el.button.disabled = isLoading;
        el.buttonSpinner.hidden = !isLoading;
        el.buttonLabel.textContent = isLoading ? "Checking..." : "Check Status";
    }

    /** Show a message (used for "not found" and error states) and hide the card. */
    function showMessage(text) {
        el.message.textContent = text;
        el.message.hidden = false;
        el.card.hidden = true;
    }

    /** Hide the message and result card before starting a fresh search. */
    function clearResults() {
        el.message.hidden = true;
        el.card.hidden = true;
        // The empty-state hint has done its job once the user searches.
        el.hint.hidden = true;
    }

    // ---------------------------------------------------------------------
    // Formatting helpers
    // ---------------------------------------------------------------------

    /**
     * Turn a Date into friendly text like "2 minutes ago", "3 hours ago",
     * "Yesterday at 4:32 PM", or "Aug 10 at 4:32 PM" for anything older.
     */
    function formatLastPing(date) {
        const now = new Date();
        const diffMinutes = Math.floor((now - date) / 60000);

        // Format the clock time once, e.g. "4:32 PM" — used by the older cases.
        const clockTime = date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });

        // Within the last minute
        if (diffMinutes < 1) {
            return "Just now";
        }

        // Within the last hour → "X minutes ago"
        if (diffMinutes < 60) {
            return diffMinutes === 1 ? "1 minute ago" : diffMinutes + " minutes ago";
        }

        // Earlier today → "X hours ago"
        const isToday = date.toDateString() === now.toDateString();
        if (isToday) {
            const diffHours = Math.floor(diffMinutes / 60);
            return diffHours === 1 ? "1 hour ago" : diffHours + " hours ago";
        }

        // Yesterday → "Yesterday at 4:32 PM"
        const yesterday = new Date(now);
        yesterday.setDate(now.getDate() - 1);
        if (date.toDateString() === yesterday.toDateString()) {
            return "Yesterday at " + clockTime;
        }

        // Anything older → "Aug 10 at 4:32 PM"
        const calendarDate = date.toLocaleDateString([], { month: "short", day: "numeric" });
        return calendarDate + " at " + clockTime;
    }

    /**
     * Decide the badge from how long ago the device last pinged:
     *   - Online  (green)  → within the last 5 minutes
     *   - Idle    (yellow) → within the last hour
     *   - Offline (red)    → more than an hour ago (or never)
     * Returns the label plus a CSS suffix used to colour the badge and card.
     */
    function getStatus(lastPingDate) {
        // No ping date at all means the device has never reported in.
        if (!lastPingDate) {
            return { label: "Offline", cssSuffix: "offline" };
        }

        const minutesAgo = (Date.now() - lastPingDate.getTime()) / 60000;

        if (minutesAgo <= 5) {
            return { label: "Online", cssSuffix: "online" };
        }
        if (minutesAgo <= 60) {
            return { label: "Idle", cssSuffix: "idle" };
        }
        return { label: "Offline", cssSuffix: "offline" };
    }

    /**
     * Pull a display name out of DeviceStatusInfo's driver property.
     * Geotab uses a special "UnknownDriverId" entity when nobody is assigned,
     * so that case (and a missing driver) returns null → "No driver assigned".
     */
    function getDriverName(driver) {
        if (!driver || driver === "UnknownDriverId" || driver.id === "UnknownDriverId") {
            return null;
        }

        // Prefer "First Last"; fall back to the account's display name.
        const fullName = [driver.firstName, driver.lastName].filter(Boolean).join(" ");
        return fullName || driver.name || null;
    }

    // ---------------------------------------------------------------------
    // Rendering
    // ---------------------------------------------------------------------

    /**
     * Fill the result card with the device + status info and show it.
     * Also swaps the colour classes on the badge and card so the glow
     * matches the status (online / idle / offline).
     */
    function renderResult(device, statusInfo) {
        // DeviceStatusInfo.dateTime is the timestamp of the last GPS record —
        // this is the "last ping". Guard against it being missing.
        const lastPingDate = statusInfo.dateTime ? new Date(statusInfo.dateTime) : null;
        const status = getStatus(lastPingDate);

        // Text fields
        el.vehicleName.textContent = device.name || "Unnamed device";
        el.serialNumber.textContent = device.serialNumber || "—";
        el.driverName.textContent = getDriverName(statusInfo.driver) || "No driver assigned";
        el.lastPing.textContent = lastPingDate ? formatLastPing(lastPingDate) : "Never";

        // Status badge + matching card glow. Setting className from scratch
        // clears whichever colour class the previous search left behind.
        el.badge.textContent = status.label;
        el.badge.className = "dsc-badge dsc-badge--" + status.cssSuffix;
        el.card.className = "dsc-card dsc-card--" + status.cssSuffix;

        // GPS footer — only shown when the status info includes a position.
        if (typeof statusInfo.latitude === "number" && typeof statusInfo.longitude === "number") {
            const position = statusInfo.latitude.toFixed(5) + ", " + statusInfo.longitude.toFixed(5);
            const speed = typeof statusInfo.speed === "number"
                ? " · " + Math.round(statusInfo.speed) + " km/h"
                : "";
            el.gps.textContent = "Last position: " + position + speed;
            el.gps.hidden = false;
        } else {
            el.gps.hidden = true;
        }

        el.card.hidden = false;
    }

    // ---------------------------------------------------------------------
    // The main search flow
    // ---------------------------------------------------------------------

    /**
     * Runs when the user clicks "Check Status" (or presses Enter).
     * Two chained API calls:
     *   1. Get Device      — find the device whose serial number matches.
     *   2. Get DeviceStatusInfo — fetch its last ping time, driver, and GPS.
     */
    function checkStatus() {
        // Normalise the input: Geotab stores serial numbers in upper case
        // without dashes or spaces, so strip those before searching.
        const serial = el.input.value.trim().replace(/[\s-]/g, "").toUpperCase();

        // Nothing typed — nudge the user instead of calling the API.
        if (!serial) {
            showMessage("Please enter a serial number first.");
            el.hint.hidden = true;
            return;
        }

        clearResults();
        setLoading(true);

        // Shared error handler: any failed API call lands here.
        const onApiError = function () {
            setLoading(false);
            showMessage("Something went wrong. Please try again.");
        };

        // Step 1: search for the device by serial number.
        api.call("Get", {
            typeName: "Device",
            search: { serialNumber: serial },
            resultsLimit: 1
        }, function (devices) {

            // Not-found state: the search returned an empty list.
            if (!devices || devices.length === 0) {
                setLoading(false);
                showMessage("No device found with that serial number.");
                return;
            }

            const device = devices[0];

            // Step 2: fetch the live status info for that one device.
            api.call("Get", {
                typeName: "DeviceStatusInfo",
                search: { deviceSearch: { id: device.id } }
            }, function (statusInfos) {
                setLoading(false);

                // A device that has never communicated may have no status yet.
                if (!statusInfos || statusInfos.length === 0) {
                    showMessage("No status information is available for this device yet.");
                    return;
                }

                renderResult(device, statusInfos[0]);
            }, onApiError);

        }, onApiError);
    }

    // ---------------------------------------------------------------------
    // Add-in lifecycle
    // MyGeotab calls these three methods to drive the add-in:
    //   initialize — once, when the add-in first loads
    //   focus      — every time the user navigates to the page
    //   blur       — every time the user navigates away
    // ---------------------------------------------------------------------
    return {

        initialize: function (freshApi, freshState, initializeCallback) {
            api = freshApi;

            cacheElements();

            // Wire up the two ways to start a search: clicking the button,
            // or pressing Enter while typing in the input.
            el.button.addEventListener("click", checkStatus);
            el.input.addEventListener("keydown", function (event) {
                if (event.key === "Enter") {
                    checkStatus();
                }
            });

            // Tell MyGeotab we're ready — it won't show the page until this runs.
            initializeCallback();
        },

        focus: function (freshApi) {
            // MyGeotab may hand us a fresh API object on each visit; keep it.
            api = freshApi;

            // Put the cursor in the input so the user can type immediately.
            el.input.focus();
        },

        blur: function () {
            // Nothing to clean up — no timers or subscriptions are running.
        }
    };
};
