/**
 * ISE CRM <-> Google bridge (Google Apps Script). Runs every 5 minutes in your Google account.
 *
 *  - Puts CRM follow-ups on Google Calendar, and moves them in the CRM when you move them in Calendar.
 *  - Shows your meetings in the CRM (Today screen, recaps), linked to clients by attendee or name.
 *  - Logs Google Voice texts, voicemails and missed calls from Gmail onto the matching contact.
 *    Unknown numbers land in the CRM Inbox as possible new leads.
 *
 * SETUP (about 3 minutes)
 *  1. script.google.com -> New project -> paste this file over Code.gs. Name it "ISE CRM bridge".
 *  2. Services (+, left sidebar) -> Google Calendar API -> Add.
 *  3. Project Settings (gear) -> Script Properties -> add:
 *       CRM_URL    https://qpokiahksvtbcwsjlhzn.supabase.co/functions/v1/crm
 *       CRM_TOKEN  (ISE CRM app -> More -> Connections & setup -> Show token)
 *     Optional:
 *       CALENDARS  calendar ids to show in the CRM, comma-separated (default: primary)
 *       TARGET_CALENDAR  where CRM follow-ups go (default: primary)
 *       GV_NUMBER  your Google Voice number, e.g. +19085550100 (helps read text senders)
 *       SYNC_VOICE  false to skip Gmail / Google Voice (default: true)
 *  4. Select the function "setup" and press Run. Approve the permissions (your own account).
 *     That installs the 5-minute trigger and runs the first sync.
 */

const PROPS = PropertiesService.getScriptProperties()

function cfg_() {
  const list = (k, d) => (PROPS.getProperty(k) || d).split(',').map((s) => s.trim()).filter(Boolean)
  return {
    url: (PROPS.getProperty('CRM_URL') || '').replace(/\/$/, ''),
    token: PROPS.getProperty('CRM_TOKEN') || '',
    calendars: list('CALENDARS', 'primary'),
    target: PROPS.getProperty('TARGET_CALENDAR') || 'primary',
    gv: (PROPS.getProperty('GV_NUMBER') || '').replace(/\D/g, ''),
    voice: (PROPS.getProperty('SYNC_VOICE') || 'true') !== 'false',
  }
}

/** Run once by hand. */
function setup() {
  const c = cfg_()
  if (!c.url || !c.token) throw new Error('Add CRM_URL and CRM_TOKEN in Project Settings -> Script Properties first.')
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t))
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(5).create()
  sync()
  Logger.log('ISE CRM bridge is running every 5 minutes.')
}

/** The 5-minute job. Each part is independent so one failure does not block the others. */
function sync() {
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(2000)) return
  try {
    for (const step of [pushFollowUps_, pullCalendar_, scanVoice_]) {
      try {
        step()
      } catch (e) {
        console.error(step.name + ': ' + e)
      }
    }
  } finally {
    lock.releaseLock()
  }
}

function api_(method, path, body) {
  const c = cfg_()
  const res = UrlFetchApp.fetch(c.url + path, {
    method: method,
    contentType: 'application/json',
    headers: { 'x-crm-token': c.token },
    payload: body ? JSON.stringify(body) : undefined,
    muteHttpExceptions: true,
  })
  const code = res.getResponseCode()
  if (code >= 300) throw new Error(path + ' -> ' + code + ': ' + res.getContentText().slice(0, 300))
  return JSON.parse(res.getContentText() || '{}')
}

function calendarId_(id) {
  return id === 'primary' ? Calendar.Calendars.get('primary').id : id
}

// ------------------------------------------------------------------ CRM follow-ups -> Calendar

function pushFollowUps_() {
  const out = api_('get', '/bridge/outbox')
  const results = []
  for (const op of out.ops || []) {
    const calId = calendarId_(op.calendar_id === 'primary' ? cfg_().target : op.calendar_id)
    try {
      if (op.op === 'delete') {
        try {
          Calendar.Events.remove(calId, op.event_id)
        } catch (e) {
          if (!/not found|deleted|410|404/i.test(String(e))) throw e
        }
        results.push({ outbox_ids: op.outbox_ids, task_id: op.task_id, ok: true, op: 'delete' })
        continue
      }
      const resource = {
        summary: op.summary,
        description: op.description,
        start: op.all_day ? { date: op.start_date } : { dateTime: op.start },
        end: op.all_day ? { date: nextDay_(op.start_date) } : { dateTime: op.end },
        colorId: '6', // tangerine, so CRM follow-ups stand out
        extendedProperties: { private: { crmTaskId: op.task_id } },
      }
      let ev
      if (op.event_id) {
        try {
          ev = Calendar.Events.patch(resource, calId, op.event_id)
        } catch (e) {
          ev = Calendar.Events.insert(resource, calId)
        }
      } else {
        ev = Calendar.Events.insert(resource, calId)
      }
      results.push({ outbox_ids: op.outbox_ids, task_id: op.task_id, ok: true, event_id: ev.id, calendar_id: calId, op: 'upsert' })
    } catch (e) {
      results.push({ outbox_ids: op.outbox_ids, task_id: op.task_id, ok: false, error: String(e).slice(0, 300) })
    }
  }
  if (results.length) api_('post', '/bridge/ack', { results: results })
}

function nextDay_(d) {
  const x = new Date(d + 'T12:00:00Z')
  x.setUTCDate(x.getUTCDate() + 1)
  return x.toISOString().slice(0, 10)
}

// ------------------------------------------------------------------ Calendar -> CRM

function pullCalendar_() {
  const c = cfg_()
  const from = new Date(Date.now() - 3 * 86400000)
  const to = new Date(Date.now() + 30 * 86400000)
  const events = []
  const calIds = []
  for (const raw of c.calendars) {
    const calId = calendarId_(raw)
    calIds.push(calId)
    let pageToken
    do {
      const r = Calendar.Events.list(calId, {
        timeMin: from.toISOString(), timeMax: to.toISOString(), singleEvents: true, orderBy: 'startTime',
        maxResults: 250, pageToken: pageToken,
      })
      for (const e of r.items || []) {
        if (e.status === 'cancelled') continue
        if (['workingLocation', 'outOfOffice', 'birthday'].indexOf(e.eventType) >= 0) continue
        const conf = (e.conferenceData && (e.conferenceData.entryPoints || []).find((p) => p.entryPointType === 'video')) || null
        events.push({
          calendar_id: calId,
          event_id: e.id,
          title: e.summary || '(no title)',
          description: (e.description || '').slice(0, 2000),
          location: e.location || null,
          conference_url: e.hangoutLink || (conf ? conf.uri : null),
          start: e.start.dateTime || e.start.date,
          end: e.end ? e.end.dateTime || e.end.date : null,
          all_day: !e.start.dateTime,
          status: e.status || null,
          attendees: (e.attendees || []).map((a) => ({ email: a.email, name: a.displayName || null, self: !!a.self, response: a.responseStatus })),
          html_link: e.htmlLink || null,
          crm_task_id: (e.extendedProperties && e.extendedProperties.private && e.extendedProperties.private.crmTaskId) || null,
        })
      }
      pageToken = r.nextPageToken
    } while (pageToken)
  }
  api_('post', '/bridge/calendar', { from: from.toISOString(), to: to.toISOString(), calendars: calIds, events: events })
}

// ------------------------------------------------------------------ Google Voice (via Gmail) -> CRM

function scanVoice_() {
  const c = cfg_()
  if (!c.voice) return
  const label = GmailApp.getUserLabelByName('ISE CRM/logged') || GmailApp.createLabel('ISE CRM/logged')
  const threads = GmailApp.search('(from:voice-noreply@google.com OR from:txt.voice.google.com) newer_than:14d -label:"ISE CRM/logged"', 0, 40)
  if (!threads.length) {
    api_('post', '/bridge/phone', { events: [] })
    return
  }
  const events = []
  for (const t of threads) {
    for (const m of t.getMessages()) {
      const e = parseVoice_(m, c.gv)
      if (e) events.push(e)
    }
  }
  api_('post', '/bridge/phone', { events: events })
  for (const t of threads) t.addLabel(label)
}

function parseVoice_(m, gvDigits) {
  const subject = m.getSubject() || ''
  const from = m.getFrom() || ''
  let kind = null
  if (/text message/i.test(subject) || /txt\.voice\.google\.com/i.test(from)) kind = 'text'
  else if (/voicemail/i.test(subject)) kind = 'voicemail'
  else if (/missed call/i.test(subject)) kind = 'missed_call'
  if (!kind) return null

  let phone = (subject.match(/\+?1?\s*\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}/) || [])[0] || null
  const addr = (from.match(/([^<>\s"]+@txt\.voice\.google\.com)/i) || [])[1]
  if (!phone && addr) {
    const nums = addr.split('@')[0].split('.').filter((s) => /^\d{10,11}$/.test(s))
    const other = nums.find((n) => !gvDigits || gvDigits.slice(-10) !== n.slice(-10))
    if (other) phone = other
  }
  const nameMatch = subject.match(/from\s+(.+?)(?:\s+at\s+\d|$)/i)
  const name = nameMatch && !/\d{3}/.test(nameMatch[1]) ? nameMatch[1].replace(/[()"]/g, '').trim() : null

  let text = (m.getPlainBody() || '').replace(/\r/g, '')
  text = text.split(/\n\s*(?:To respond to this text message|YOUR ACCOUNT|HELP CENTER|HELP FORUM|Google LLC|Play message|Listen to message)/i)[0].trim()
  if (kind === 'missed_call') text = ''
  return {
    external_id: m.getId(),
    kind: kind,
    direction: 'in',
    phone: phone,
    name: name,
    body: text ? text.slice(0, 2000) : null,
    occurred_at: m.getDate().toISOString(),
  }
}
