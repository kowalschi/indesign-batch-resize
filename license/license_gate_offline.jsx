/*
 * Copyright (c) 2026 Dan Ichimescu. constantindan@gmail.com
 *
 * OFFLINE license gate - 5 day trial, no internet, no server.
 * On the first run it writes the date into a folder in %APPDATA%
 * (Windows) / Application Support (Mac). After 5 days the script stops.
 *
 * Same API as the server gate: every script calls LICENSE.check().
 */

var LICENSE = (function () {

    var TRIAL_DAYS = 5;                 // <-- the limit
    var CONTACT = "constantindan@gmail.com";
    var DAY_MS = 24 * 60 * 60 * 1000;

    // Folder.userData = %APPDATA% on Windows, ~/Library/Application Support on Mac
    var dir = Folder(Folder.userData.fsName + "/IBR");   // the folder created in %APPDATA%
    var marker = File(dir.fsName + "/state.dat");

    function nowMs() { return (new Date()).getTime(); }

    // light obfuscation + checksum, so the file can't just be edited to extend the date
    function encode(first, last) {
        var sum = (first + last) % 100000;
        return first.toString(36) + "-" + last.toString(36) + "-" + sum.toString(36);
    }
    function decode(txt) {
        try {
            var p = String(txt).replace(/^\s+|\s+$/g, "").split("-");
            if (p.length !== 3) return null;
            var first = parseInt(p[0], 36);
            var last = parseInt(p[1], 36);
            var chk = parseInt(p[2], 36);
            if (isNaN(first) || isNaN(last) || isNaN(chk)) return null;
            if (((first + last) % 100000) !== chk) return null;   // tampered
            return { first: first, last: last };
        } catch (e) { return null; }
    }

    function readState() {
        if (!marker.exists) return "missing";
        try {
            marker.encoding = "UTF-8";
            marker.open("r");
            var txt = marker.read();
            marker.close();
            var st = decode(txt);
            return st ? st : "tampered";
        } catch (e) { return "tampered"; }
    }
    function writeState(first, last) {
        try {
            if (!dir.exists) dir.create();
            marker.encoding = "UTF-8";
            marker.open("w");
            marker.write(encode(first, last));
            marker.close();
            try { marker.hidden = true; } catch (e) {}
        } catch (e) {}
    }

    function stop(reason) {
        try { app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL; } catch (e) {}
        alert("Trial period ended.\n\n(" + reason + ")\n\nContact: " + CONTACT);
        exit();
        throw new Error("license");
    }

    return {
        check: function () {
            var now = nowMs();
            var st = readState();

            if (st === "missing") {        // first run ever -> start the 5 day clock
                writeState(now, now);
                return true;
            }
            if (st === "tampered") {       // someone edited the file
                stop("tampered");
            }

            // clock set back? (allow 1h skew) -> treat as cheating
            if (now < st.last - 60 * 60 * 1000) {
                stop("clock");
            }

            var elapsed = now - st.first;
            if (elapsed > TRIAL_DAYS * DAY_MS) {
                stop("expired");
            }

            // remember the latest run time
            writeState(st.first, (now > st.last ? now : st.last));

            var daysLeft = Math.ceil((TRIAL_DAYS * DAY_MS - elapsed) / DAY_MS);
            if (daysLeft <= 2) {
                alert("Trial license: " + daysLeft + " day(s) left.");
            }
            return true;
        }
    };

})();
