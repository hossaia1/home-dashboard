/**
 * Home Dashboard — Google Calendar bridge
 * ---------------------------------------
 * Deployed as a Web App on the family Google account, this gives the dashboard
 * live read AND write access to the calendar without any OAuth in the browser.
 *
 * Why this rather than the browser talking to Google directly:
 *   - No access tokens, so nothing expires on an unattended wall tablet.
 *   - No third-party CORS proxy, so the calendar never leaves Google.
 *   - Reads are live, instead of waiting hours for Google's .ics export.
 *
 * SECURITY: the deployment URL is the only credential. Anyone holding it can
 * read and add events. Treat it like a password — keep it out of git.
 *
 * ---------------------------------------------------------------------------
 * SETUP (about 10 minutes, all in the browser)
 *
 *  1. Sign in as hossainayubfamily@gmail.com
 *  2. Go to https://script.google.com  →  New project
 *  3. Delete the placeholder code, paste this whole file in
 *  4. Set CALENDAR_ID below (see the note on it)
 *  5. Click Deploy → New deployment
 *       - Type:            Web app
 *       - Execute as:      Me (hossainayubfamily@gmail.com)
 *       - Who has access:  Anyone
 *     "Anyone" sounds alarming but means "anyone with the unguessable URL" —
 *     it is what lets the tablet call it without signing in.
 *  6. Authorise when prompted. Google will warn the app is unverified; it's
 *     your own script, so continue via "Advanced → Go to project".
 *  7. Copy the Web app URL (ends in /exec)
 *  8. Paste it into the dashboard: Settings → Calendar bridge URL → Save
 *
 * If you later edit this script, you must Deploy → Manage deployments → edit →
 * Version: New version, or the tablet keeps calling the old code.
 * ---------------------------------------------------------------------------
 */

// 'primary' = the account's own calendar. To target the separate
// "Hossain_Ayub Family" calendar instead, use its Calendar ID from
// Google Calendar → Settings → that calendar → Integrate calendar.
// It looks like: abc123...@group.calendar.google.com
const CALENDAR_ID = 'primary';

// How far ahead to return events when the dashboard asks for the agenda.
const DEFAULT_DAYS = 90;

// Bump this whenever you change this file. ?action=ping reports it, which is the
// only reliable way to tell whether a new version actually published — the
// deployment URL keeps serving a frozen snapshot until you republish, and
// list/ping alone look identical before and after.
const SCRIPT_VERSION = 3;   // v2 = update + delete, v3 = masjid time verification

// Public prayer-times page for the masjid the household follows. Fetched
// server-side here so the tablet never scrapes anything directly and no
// third-party CORS proxy is involved.
const MASJID_URL = 'https://masjidbox.com/prayer-times/damic';


function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  try {
    if (action === 'list') {
      var days = Number((e.parameter && e.parameter.days) || DEFAULT_DAYS);
      return json({ ok: true, events: listEvents(days) });
    }
    if (action === 'masjid') return json(fetchMasjidTimes());
    if (action === 'ping') {
      return json({
        ok: true,
        version: SCRIPT_VERSION,
        calendar: cal().getName(),
        supports: ['list', 'create', 'update', 'delete']
      });
    }
    return json({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/**
 * The dashboard posts as text/plain on purpose: that counts as a "simple"
 * CORS request, so the browser skips the preflight OPTIONS call, which Apps
 * Script web apps cannot answer.
 */
function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'create') return json({ ok: true, event: createEvent(body) });
    if (body.action === 'update') return json({ ok: true, event: updateEvent(body) });
    if (body.action === 'delete') return json({ ok: true, deleted: deleteEvent(body.id) });
    return json({ ok: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}


function cal() {
  var c = CALENDAR_ID === 'primary'
    ? CalendarApp.getDefaultCalendar()
    : CalendarApp.getCalendarById(CALENDAR_ID);
  if (!c) throw new Error('Calendar not found: ' + CALENDAR_ID);
  return c;
}

function listEvents(days) {
  var start = new Date();
  start.setHours(0, 0, 0, 0);
  var end = new Date(start.getTime() + days * 86400000);

  return cal().getEvents(start, end).map(function (ev) {
    return {
      id:       ev.getId(),
      title:    ev.getTitle() || '(No title)',
      start:    ev.getStartTime().toISOString(),
      end:      ev.getEndTime().toISOString(),
      allDay:   ev.isAllDayEvent(),
      location: ev.getLocation() || ''
    };
  });
}

function createEvent(b) {
  if (!b.title) throw new Error('title is required');
  var c = cal();
  var ev;

  if (b.allDay) {
    ev = c.createAllDayEvent(b.title, new Date(b.start));
  } else {
    if (!b.end) throw new Error('end is required for timed events');
    ev = c.createEvent(b.title, new Date(b.start), new Date(b.end));
  }
  if (b.location) ev.setLocation(b.location);

  return { id: ev.getId(), title: ev.getTitle(), start: ev.getStartTime().toISOString() };
}

/**
 * Switching an event between timed and all-day isn't something CalendarApp can
 * do in place, so in that case we delete and recreate. The event ID changes as
 * a result — the dashboard re-fetches after every write, so it picks up the new
 * one rather than holding a stale reference.
 */
function updateEvent(b) {
  if (!b.id) throw new Error('id is required');
  var ev = cal().getEventById(b.id);
  if (!ev) throw new Error('Event not found');

  var wasAllDay = ev.isAllDayEvent();
  if (!!b.allDay !== wasAllDay) {
    ev.deleteEvent();
    return createEvent(b);
  }

  if (b.title) ev.setTitle(b.title);
  if (b.allDay) {
    ev.setAllDayDate(new Date(b.start));
  } else {
    ev.setTime(new Date(b.start), new Date(b.end));
  }
  if (b.location !== undefined) ev.setLocation(b.location || '');

  return { id: ev.getId(), title: ev.getTitle(), start: ev.getStartTime().toISOString() };
}

function deleteEvent(id) {
  var ev = cal().getEventById(id);
  if (!ev) throw new Error('Event not found');
  ev.deleteEvent();
  return true;
}

/**
 * Reads the masjid's published times so the dashboard can check its own
 * calculation against them.
 *
 * Deliberately anchored on the label text plus the stable `time mono` wrapper,
 * NOT on masjidbox's CSS class names — those are build-generated hashes
 * (styles__Wrapper-sc-1rm9q09-0) and change on every deploy of theirs.
 *
 * Always reports ok:false with a reason rather than partial data. A verifier
 * that quietly claims success while broken is worse than no verifier.
 */
function fetchMasjidTimes() {
  var res;
  try {
    res = UrlFetchApp.fetch(MASJID_URL, { muteHttpExceptions: true, followRedirects: true });
  } catch (err) {
    return { ok: false, error: 'fetch failed: ' + err };
  }
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: 'HTTP ' + res.getResponseCode() };
  }

  var html = res.getContentText();
  var want = { Fajr: 'Fajr', Shuruq: 'Shuruq', Dhuhr: 'Dhuhr',
               Asr: 'Asr', Maghrib: 'Maghrib', Isha: 'Isha' };
  var times = {}, missing = [];

  Object.keys(want).forEach(function (key) {
    var i = html.indexOf('>' + want[key] + '<');
    if (i === -1) { missing.push(key); return; }
    var seg = html.substring(i, i + 400);
    var m = seg.match(/<div class="time mono">(\d{1,2})<div[^>]*><\/div>(\d{2})<\/div>/);
    if (!m) { missing.push(key); return; }
    times[key] = ('0' + m[1]).slice(-2) + ':' + m[2];
  });

  if (missing.length) {
    return { ok: false, error: 'could not parse: ' + missing.join(', ') +
             ' — the page layout has probably changed' };
  }

  // The page states the date it is showing. Returning it lets the dashboard
  // refuse to compare against a stale or cached page.
  var d = html.match(/([A-Z][a-z]+day, [A-Z][a-z]+ \d+, \d{4})/);

  return { ok: true, source: 'masjidbox/damic', pageDate: d ? d[1] : null,
           fetchedAt: new Date().toISOString(), times: times };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
