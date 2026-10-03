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
// v2 = update + delete, v3 = masjid verification, v4 = JSON-LD parsing + Jumu'ah,
// v5 = writes via GET (POST was being downgraded by the 302) + action echo,
// v6 = location + description read/write,
// v7 = ownership flag; invitations can't be edited/deleted, declined ones hidden
// v8 = find events whose ID getEventById can't resolve (imported from Outlook)
const SCRIPT_VERSION = 8;

// Public prayer-times page for the masjid the household follows. Fetched
// server-side here so the tablet never scrapes anything directly and no
// third-party CORS proxy is involved.
const MASJID_URL = 'https://masjidbox.com/prayer-times/damic';


/**
 * Writes arrive as GET, not POST.
 *
 * /exec answers with a 302 to script.googleusercontent.com. Per the fetch spec a
 * 302 on a POST is re-issued as a GET with the body discarded — so the write
 * silently became a read, landed in the default 'list' branch, and returned
 * {ok:true}. The dashboard saw success while nothing had been written.
 *
 * GET has no body to lose and needs no CORS preflight. Every response now echoes
 * the action it performed so the caller can verify it got an answer to the
 * question it actually asked.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'list';
  try {
    if (action === 'list') {
      return json({ ok: true, action: action, events: listEvents(Number(p.days || DEFAULT_DAYS)) });
    }
    if (action === 'masjid') {
      var m = fetchMasjidTimes();
      m.action = action;
      return json(m);
    }
    if (action === 'ping') {
      return json({ ok: true, action: action, version: SCRIPT_VERSION,
                    calendar: cal().getName(),
                    supports: ['list', 'create', 'update', 'delete', 'masjid'] });
    }
    if (action === 'create') {
      return json({ ok: true, action: action, event: createEvent(fromParams(p)) });
    }
    if (action === 'update') {
      return json({ ok: true, action: action, event: updateEvent(fromParams(p)) });
    }
    if (action === 'delete') {
      return json({ ok: true, action: action, deleted: deleteEvent(p.id, p.origStart || p.start) });
    }
    return json({ ok: false, action: action, error: 'Unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, action: action, error: String(err) });
  }
}

// Query params arrive as strings; normalise into the shape create/update expect.
// A whitespace-only value means "clear this field" — the dashboard sends a single
// space to distinguish "cleared" from "not supplied", since an empty query param
// can't survive the round trip.
function fromParams(p) {
  function text(v) { return v === undefined ? undefined : String(v).trim(); }
  return {
    id: p.id,
    origStart: p.origStart,
    title: text(p.title),
    start: p.start,
    end: p.end,
    location: text(p.location),
    description: text(p.description),
    allDay: p.allDay === 'true' || p.allDay === '1'
  };
}

/**
 * Kept for compatibility. If a POST does survive without being downgraded, this
 * still works — but the dashboard no longer relies on it.
 */
function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'create') return json({ ok: true, action: 'create', event: createEvent(body) });
    if (body.action === 'update') return json({ ok: true, action: 'update', event: updateEvent(body) });
    if (body.action === 'delete') return json({ ok: true, action: 'delete', deleted: deleteEvent(body.id, body.origStart) });
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

  return cal().getEvents(start, end).filter(function (ev) {
    // Invitations you've declined shouldn't clutter the family dashboard.
    try { return ev.getMyStatus() !== CalendarApp.GuestStatus.NO; } catch (e) { return true; }
  }).map(function (ev) {
    var owned = true;
    try { owned = ev.isOwnedByMe(); } catch (e) {}
    return {
      // Events sent by someone else (e.g. school invitations from Outlook) are
      // owned by the organiser. They can't be edited or deleted from here, so
      // the dashboard needs to know which is which.
      owned:    owned,
      id:       ev.getId(),
      title:    ev.getTitle() || '(No title)',
      start:    ev.getStartTime().toISOString(),
      end:      ev.getEndTime().toISOString(),
      allDay:   ev.isAllDayEvent(),
      location: ev.getLocation() || '',
      // Trimmed: Google stores meeting boilerplate and HTML here, and the
      // dashboard only ever shows a short excerpt.
      description: (ev.getDescription() || '').slice(0, 1000)
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
  if (b.location)    ev.setLocation(b.location);
  if (b.description) ev.setDescription(b.description);

  return { id: ev.getId(), title: ev.getTitle(), start: ev.getStartTime().toISOString() };
}

/**
 * Switching an event between timed and all-day isn't something CalendarApp can
 * do in place, so in that case we delete and recreate. The event ID changes as
 * a result — the dashboard re-fetches after every write, so it picks up the new
 * one rather than holding a stale reference.
 */
/**
 * Events that arrived from Outlook/Exchange report an Outlook UID from getId()
 * (040000008200E000…), and getEventById() cannot resolve that same value — it
 * returns null for an event that plainly exists. That made edit and delete fail
 * with "Event not found" for most of the family's events.
 *
 * So: try the direct lookup, then fall back to scanning events around the
 * event's original start time and matching by ID. The window is small, so it's
 * one cheap query rather than a crawl of the whole calendar.
 */
function findEvent(id, startIso) {
  var c = cal(), ev = null;
  try { ev = c.getEventById(id); } catch (e) {}
  if (ev) return ev;

  var from, to;
  if (startIso) {
    var s = new Date(startIso);
    from = new Date(s.getTime() - 36 * 3600000);
    to   = new Date(s.getTime() + 36 * 3600000);
  } else {
    from = new Date(Date.now() - 30 * 86400000);
    to   = new Date(Date.now() + 400 * 86400000);
  }
  var list = c.getEvents(from, to);
  for (var i = 0; i < list.length; i++) {
    if (list[i].getId() === id) return list[i];
  }
  return null;
}

function updateEvent(b) {
  if (!b.id) throw new Error('id is required');
  // origStart: where the event was, so it can be found even if this edit moves it.
  var ev = findEvent(b.id, b.origStart || b.start);
  if (!ev) throw new Error('Event not found');
  if (!ev.isOwnedByMe()) throw new Error(NOT_OWNED_MSG);

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
  // Only touch these when the caller actually sent them, so an edit from a
  // client that doesn't know about a field can't blank it out.
  if (b.location    !== undefined) ev.setLocation(b.location || '');
  if (b.description !== undefined) ev.setDescription(b.description || '');

  return { id: ev.getId(), title: ev.getTitle(), start: ev.getStartTime().toISOString() };
}

// The dashboard recognises this text and offers to hide the event instead.
var NOT_OWNED_MSG = 'NOT_OWNED: this is an invitation from someone else, so it can only be ' +
                    'removed by its organiser (or declined in Google Calendar).';

/**
 * Deliberately does NOT decline invitations automatically. Declining sends a
 * reply to the organiser — for school events that would mean an email to the
 * school on the family's behalf, which is not what "delete" should do silently.
 */
function deleteEvent(id, startIso) {
  var ev = findEvent(id, startIso);
  if (!ev) throw new Error('Event not found');
  if (!ev.isOwnedByMe()) throw new Error(NOT_OWNED_MSG);
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
  var out = parseJsonLd(html);
  if (out.ok) return out;

  // Fall back to reading the rendered markup if the structured data ever goes
  // away. Less trustworthy, so the response says which route produced it.
  var alt = parseVisibleTimes(html);
  if (alt.ok) { alt.degraded = out.error; return alt; }
  return { ok: false, error: 'JSON-LD: ' + out.error + ' | markup: ' + alt.error };
}

/**
 * Preferred route. The page embeds schema.org Event objects — one per prayer per
 * day for the coming week, each with a full ISO datetime including the Helsinki
 * offset. Far more dependable than reading styled <div>s, and it's the only
 * place the Jumu'ah time appears in machine-readable form.
 *
 * Note Fridays carry a "Jumuah Prayer" event *instead of* a Dhuhr one, which is
 * what tells us Jumu'ah replaces Dhuhr rather than sitting alongside it.
 */
function parseJsonLd(html) {
  var re = /<script[^>]*application\/ld\+json[^>]*>([\s\S]*?)<\/script>/g;
  var events = [], m;
  while ((m = re.exec(html)) !== null) {
    try {
      var parsed = JSON.parse(m[1]);
      var arr = (parsed instanceof Array) ? parsed : [parsed];
      for (var i = 0; i < arr.length; i++) {
        if (arr[i] && arr[i]['@type'] === 'Event' && arr[i].startDate) events.push(arr[i]);
      }
    } catch (e) { /* one bad block shouldn't sink the rest */ }
  }
  if (!events.length) return { ok: false, error: 'no schema.org events found' };

  // "2026-08-07T13:31:00+03:00" -> date and HH:MM taken literally. Parsing to a
  // Date would reinterpret this in the script's timezone and shift the clock.
  function parts(iso) {
    var mm = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/.exec(iso);
    return mm ? { date: mm[1], time: mm[2] } : null;
  }

  var todayStr = Utilities.formatDate(new Date(), 'Europe/Helsinki', 'yyyy-MM-dd');
  var names = ['Fajr', 'Shuruq', 'Dhuhr', 'Asr', 'Maghrib', 'Isha'];
  var times = {}, jumuah = null;

  events.forEach(function (ev) {
    var p = parts(ev.startDate);
    if (!p) return;
    // Deliberately IGNORE the schema.org "Jumuah Prayer" event. On this page it
    // reports the Dhuhr time (13:31) while both the visible page and masjidbox's
    // own data payload say 12:32 — their generator appears to derive it from
    // Dhuhr. Jumu'ah is read from the data payload instead, see parseJumuah().
    if (/Jumuah/i.test(ev.name)) return;
    if (p.date !== todayStr) return;
    for (var i = 0; i < names.length; i++) {
      if (ev.name.indexOf(names[i]) !== -1) { times[names[i]] = p.time; break; }
    }
  });

  var missing = names.filter(function (n) { return !times[n]; });
  if (missing.length) return { ok: false, error: 'missing today: ' + missing.join(', ') };

  return { ok: true, via: 'json-ld', source: 'masjidbox/damic',
           pageDate: todayStr, fetchedAt: new Date().toISOString(),
           times: times, jumuah: parseJumuah(html),
           daysPublished: countDays(events, parts) };
}

/**
 * Jumu'ah comes from masjidbox's own serialised data payload, which is embedded
 * URL-encoded in the page as "jumuah":["2026-08-07T12:32:00+03:00"].
 *
 * This is the value the masjid actually configured — it matches what a human
 * reads on the page. The schema.org event on the same page says 13:31 (the
 * Dhuhr time), so the two disagree and this one is right.
 *
 * It's an array because a masjid can hold more than one sitting; we take the
 * earliest and report how many were published.
 */
function parseJumuah(html) {
  var re = /%22jumuah%22%3A%5B%22([\d-]{10})T(\d{2})%3A(\d{2})/g;   // encoded form
  var alt = /"jumuah"\s*:\s*\[\s*"([\d-]{10})T(\d{2}):(\d{2})/g;    // if ever plain
  var found = [], m;
  while ((m = re.exec(html))  !== null) found.push({ date: m[1], time: m[2] + ':' + m[3] });
  while ((m = alt.exec(html)) !== null) found.push({ date: m[1], time: m[2] + ':' + m[3] });
  if (!found.length) return null;

  found.sort(function (a, b) { return (a.date + a.time) < (b.date + b.time) ? -1 : 1; });
  return { date: found[0].date, time: found[0].time,
           sittings: found.filter(function (f) { return f.date === found[0].date; }).length };
}

function countDays(events, parts) {
  var seen = {};
  events.forEach(function (e) { var p = parts(e.startDate); if (p) seen[p.date] = 1; });
  return Object.keys(seen).length;
}

// Legacy fallback: read the visible time cells. Anchored on the label text and
// the "time mono" wrapper rather than masjidbox's generated class hashes, which
// change whenever they deploy.
function parseVisibleTimes(html) {
  var names = ['Fajr', 'Shuruq', 'Dhuhr', 'Asr', 'Maghrib', 'Isha'];
  var times = {}, missing = [];
  names.forEach(function (key) {
    var i = html.indexOf('>' + key + '<');
    if (i === -1) { missing.push(key); return; }
    var seg = html.substring(i, i + 400);
    var m = seg.match(/<div class="time mono">(\d{1,2})<div[^>]*><\/div>(\d{2})<\/div>/);
    if (!m) { missing.push(key); return; }
    times[key] = ('0' + m[1]).slice(-2) + ':' + m[2];
  });
  if (missing.length) return { ok: false, error: 'could not parse ' + missing.join(', ') };
  var d = html.match(/([A-Z][a-z]+day, [A-Z][a-z]+ \d+, \d{4})/);
  return { ok: true, via: 'markup', source: 'masjidbox/damic',
           pageDate: d ? d[1] : null, fetchedAt: new Date().toISOString(),
           times: times, jumuah: parseJumuah(html) };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
