/**
 * =====================================================
 * AUTOGLANCIS - Google Calendar Booking Handler
 * =====================================================
 * 
 * This Google Apps Script receives booking requests from
 * your static booking-modal.html and creates real events
 * in your Google Calendar (with conflict detection + real availability).
 *
 * SETUP INSTRUCTIONS (read carefully):
 *
 * 1. Go to https://script.google.com and create a NEW PROJECT
 * 2. Delete the default myFunction code
 * 3. Paste this entire file content
 * 4. Click Save (give it a name like "AutoGlancis Booking")
 * 5. Click "Deploy" → "New deployment"
 *    - Type: Web app
 *    - Description: AutoGlancis public booking form
 *    - Execute as: Me
 *    - Who has access: Anyone (or "Anyone, even anonymous")
 * 6. Click "Deploy" and AUTHORIZE when prompted (important!)
 * 7. Copy the "Web app URL" (it ends with /exec)
 * 8. Open booking-modal.html and replace the placeholder URL
 *    with your real URL (look for BOOKING_SCRIPT_URL)
 * 9. Deploy your website (git push) and test
 *
 * IMPORTANT NOTES:
 * - The first time someone books it may take 5-10 seconds
 * - All times are handled in Europe/Riga timezone
 * - If you change your Google Calendar settings, re-deploy the script
 *
 * Created for autoglancis.lv - Mobile car detailing (Ventspils)
 */

// =====================================================
// CONFIGURATION
// =====================================================

const CONFIG = {
  BUSINESS_NAME: 'AutoGlancis',
  BUSINESS_EMAIL: 'autoglancis@gmail.com',
  CALENDAR: CalendarApp.getDefaultCalendar(),
  EVENT_COLOR: 7,
  TIMEZONE: 'Europe/Riga',
  BUFFER_MINUTES: 30
};

const SERVICE_DURATIONS = {
  'standarta-salona-atra-tirisana': 2.5,
  'auto-salona-tirisana-60': 4,
  'detalizeta-kimiska-apstrade-120': 5,
  'standarta-auto-arpuses-tirisana': 1,
  'auto-arpuses-detalizeta-tirisana-45': 2.5
};

// Required notice period (in hours) before the service start time
const NOTICE_HOURS = {
  'standarta-salona-atra-tirisana': 2,      // Ātrā salona tīrīšana
  'standarta-auto-arpuses-tirisana': 2,     // Standarta virsbūves tīrīšana
  // All other services default to 12 hours
};

// =====================================================
// MAIN ENTRY POINTS
// =====================================================

function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var params = JSON.parse(e.postData.contents || "{}");

    if (params.action === 'getBusyTimes') {
      var cache = CacheService.getScriptCache();
      var cached = cache.get("busy_times_cache");
      if (cached != null) {
        return createJsonOutputWithHeaders(cached, headers);
      }

      var start = new Date(params.start);
      var end   = new Date(params.end);

      var events = CalendarApp.getDefaultCalendar().getEvents(start, end);
      var busy = [];

      for (var i = 0; i < events.length; i++) {
        if (events[i].isAllDayEvent()) continue;

        var endWithBuffer = new Date(events[i].getEndTime().getTime() + CONFIG.BUFFER_MINUTES * 60 * 1000);

        busy.push({
          start: events[i].getStartTime().toISOString(),
          end:   endWithBuffer.toISOString()
        });
      }

      var payload = JSON.stringify({ success: true, busy: busy });
      cache.put("busy_times_cache", payload, 60);

      return createJsonOutputWithHeaders(payload, headers);
    }

    if (params.action === 'createBooking') {
      var cal = CalendarApp.getDefaultCalendar();

      // === PROPER EUROPE/RIGA TIME HANDLING ===
      // We prefer date + time + durationHours (sent by frontend) over the UTC ISO strings
      // to avoid the common UTC vs Riga timezone shift bug.
      var startTime;
      var endTime;

      if (params.date && params.time) {
        // Best path: build dates correctly in Europe/Riga
        var durationH = parseFloat(params.durationHours) || 3;
        startTime = parseDateTime(params.date, params.time);
        endTime   = new Date(startTime.getTime() + durationH * 60 * 60 * 1000);
      } else if (params.startTime && params.endTime) {
        // Fallback for old clients (still better than nothing)
        startTime = new Date(params.startTime);
        endTime   = new Date(params.endTime);
      } else {
        // Last resort – should never happen
        throw new Error("Missing date/time information for booking");
      }

      // === CONFLICT CHECK (prevents double bookings) ===
      if (hasConflict(startTime, endTime)) {
        return createJsonOutputWithHeaders({ 
          success: false, 
          error: "Šis laiks jau ir aizņemts. Lūdzu izvēlieties citu laiku." 
        }, headers);
      }

      var title = params.serviceName + " - " + params.name;
      var desc  = "Klients: " + params.name + "\n" +
                  "Tālrunis: " + params.phone + "\n" +
                  "Vieta: " + params.location + "\n" +
                  "Piezīmes: " + (params.notes || "");

      cal.createEvent(title, startTime, endTime, {
        description: desc,
        location: params.location
      });

      CacheService.getScriptCache().remove("busy_times_cache");

      var emailData = {
        serviceName: params.serviceName,
        name: params.name,
        phone: params.phone,
        email: params.email,
        location: params.location,
        notes: params.notes,
        date: params.date,
        time: params.time
      };

      try {
        sendBookingConfirmationEmail(emailData, params.durationHours || 3, startTime);
      } catch (e) {
        console.error('Email failed after booking:', e);
      }

      return createJsonOutputWithHeaders({ success: true }, headers);
    }

    return createJsonOutputWithHeaders({ success: false, error: "Unknown action" }, headers);

  } catch (err) {
    return createJsonOutputWithHeaders({ success: false, error: err.toString() }, headers);
  }
}

// Safe helper - setHeaders only works in real Web App context, not when running functions directly in the editor
function createJsonOutputWithHeaders(data, headers) {
  var out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  if (headers) {
    try {
      out.setHeaders(headers);
    } catch (e) {
      // Expected when running test functions directly in the Apps Script editor.
      // Headers will be applied correctly when the code runs as a deployed Web App.
    }
  }
  return out;
}

function doOptions(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
  var out = ContentService.createTextOutput("");
  try {
    out.setHeaders(headers);
  } catch (e) {}
  return out;
}

function doGet(e) {
  if (e.parameter.action === 'getBusyTimes') {
    const start = e.parameter.start || '';
    const end = e.parameter.end || '';
    const busy = getBusyTimes(start, end);
    return jsonResponse({ success: true, busy: busy });
  }

  return jsonResponse({
    status: 'AutoGlancis Booking API is running',
    instructions: 'Use action=getBusyTimes or POST a booking.'
  });
}

// =====================================================
// HELPER FUNCTIONS
// =====================================================

function parseDateTime(dateStr, timeStr) {
  const [year, month, day] = dateStr.split('-').map(Number);
  const [hours, minutes] = timeStr.split(':').map(Number);
  const date = new Date(year, month - 1, day, hours, minutes, 0);
  return new Date(date.toLocaleString('en-US', { timeZone: CONFIG.TIMEZONE }));
}

/**
 * Checks if the proposed booking overlaps with any existing event,
 * taking into account the BUFFER_MINUTES after each existing job.
 * This logic is now consistent with how busy periods are shown on the frontend (getBusyTimes).
 */
function hasConflict(proposedStart, proposedEnd) {
  var bufferMs = CONFIG.BUFFER_MINUTES * 60 * 1000;

  // Query a generous window around the proposed time
  var queryStart = new Date(proposedStart.getTime() - bufferMs - 6 * 60 * 60 * 1000);
  var queryEnd   = new Date(proposedEnd.getTime()   + bufferMs + 6 * 60 * 60 * 1000);

  var events = CONFIG.CALENDAR.getEvents(queryStart, queryEnd);

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (ev.isAllDayEvent()) continue;

    var evStart = ev.getStartTime();
    // Existing job blocks until its end + buffer (same rule as getBusyTimes)
    var evBlockedUntil = new Date(ev.getEndTime().getTime() + bufferMs);

    // Overlap if the proposed booking touches the blocked window of an existing job
    if (proposedStart < evBlockedUntil && proposedEnd > evStart) {
      return true;
    }
  }
  return false;
}

function buildEventDescription(data, durationHours) {
  let desc = `PAKALPOJUMS: ${data.serviceName}\n`;
  desc += `ILGUMS: ${durationHours} stundas\n\n`;
  desc += `KLIENTS:\n`;
  desc += `Vārds: ${data.name}\n`;
  desc += `Tālrunis: ${data.phone}\n`;
  
  if (data.email) {
    desc += `E-pasts: ${data.email}\n`;
  }
  
  desc += `Atrašanās vieta: ${data.location}\n\n`;

  if (data.notes) {
    desc += `PIEZIĻMES NO KLIENTA:\n${data.notes}\n\n`;
  }

  desc += `---\n`;
  desc += `Rezervēts caur autoglancis.lv booking sistēmu\n`;
  desc += `Laiks: ${data.date} ${data.time}`;

  return desc;
}

/**
 * Returns busy time ranges from Google Calendar (for real availability on the website)
 */
function getBusyTimes(startDateStr, endDateStr) {
  try {
    const today = new Date();

    // Use safe Riga-time date construction for the query range
    // This fixes cases where naive new Date("YYYY-MM-DD") misses events due to timezone boundaries.
    let start, end;

    if (startDateStr) {
      start = parseDateTime(startDateStr, "00:00");
    } else {
      start = today;
    }

    if (endDateStr) {
      end = parseDateTime(endDateStr, "23:59");
    } else {
      end = new Date(today.getTime() + 45 * 24 * 60 * 60 * 1000);
    }

    const searchEnd = new Date(end.getTime() + 2 * 24 * 60 * 60 * 1000);
    const events = CONFIG.CALENDAR.getEvents(start, searchEnd);

    return events.map(event => {
      const endTime = event.getEndTime();
      const endWithBuffer = new Date(endTime.getTime() + CONFIG.BUFFER_MINUTES * 60 * 1000);
      
      return {
        start: event.getStartTime().toISOString(),
        end: endWithBuffer.toISOString()
      };
    });
  } catch (err) {
    console.error('Error fetching busy times:', err);
    return [];
  }
}

/**
 * Sends confirmation email to client + copy to business
 */
function sendBookingConfirmationEmail(data, durationHours, startTime) {

  try {
    const dayNames = ['Svētdiena', 'Pirmdiena', 'Otrdiena', 'Trešdiena', 'Ceturtdiena', 'Piektdiena', 'Sestdiena'];
    
    // Very safe date handling
    if (!startTime && data.date && data.time) {
      startTime = parseDateTime(data.date, data.time);
    }
    
    let dayName = '';
    let dateStr = data.date || '';
    
    if (startTime) {
      dayName = dayNames[startTime.getDay()];
      dateStr = startTime.toLocaleDateString('lv-LV', { day: 'numeric', month: 'long', year: 'numeric' });
    }

    const subject = `Apstiprinājums: ${data.serviceName} – ${dateStr} plkst. ${data.time}`;

    // ============ CLIENT EMAIL (nice design) ============
    const clientHtmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 580px; margin: 0 auto; background-color: #f5f5f5; padding: 20px;">
        <div style="background: white; border-radius: 8px; padding: 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; max-width: 540px; margin: 0 auto;">
          <div style="background: #f8f9fa; padding: 32px 24px; text-align: center; border-bottom: 1px solid #e8eaed;">
            <div style="width: 48px; height: 48px; background: #34a853; border-radius: 50%; margin: 0 auto 16px; display: flex; align-items: center; justify-content: center;">
              <span style="color: white; font-size: 28px; line-height: 1;">✓</span>
            </div>
            <div style="font-size: 18px; font-weight: 500; color: #202124;">Rezervācija ir apstiprināta</div>
          </div>
          <div style="padding: 0 24px 24px;">
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Ko</div>
              <div style="font-size: 14px; color: #202124;">${data.serviceName} (${durationHours} stundas)</div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Kad</div>
              <div style="font-size: 14px; color: #202124;">${dayName}, ${dateStr} plkst. ${data.time} <span style="color:#5f6368;">(Europe/Riga)</span></div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Kurš</div>
              <div style="font-size: 14px; color: #202124; line-height: 1.5;">
                <strong>AutoGlancis</strong> — Organizators<br>
                <span style="color:#5f6368;">autoglancis@gmail.com</span><br><br>
                <strong>${data.name}</strong> — Klients<br>
                <span style="color:#5f6368;">${data.phone}</span>
              </div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Kur</div>
              <div style="font-size: 14px; color: #202124;">${data.location}</div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Papildu piezīmes</div>
              <div style="font-size: 14px; color: #202124;">${data.notes || '—'}</div>
            </div>
            <div style="padding: 16px 0;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 8px; font-weight: 500;">Lietotņu statuss</div>
              <div style="display:flex; align-items:center; gap:8px;">
                <span style="background:#34a853; color:white; font-size:12px; padding:2px 6px; border-radius:3px;">✓</span>
                <span style="font-size:14px; color:#202124;">Google Calendar</span>
              </div>
            </div>
          </div>
        </div>
        <p style="text-align:center; font-size:12px; color:#5f6368; margin-top:16px;">AutoGlancis • Ventspils • +371 2568 9885</p>
      </div>
    `;

    // Send to customer (only if they provided an email)
    if (data.email && data.email.trim() !== '') {
      MailApp.sendEmail({
        to: data.email,
        subject: subject,
        htmlBody: clientHtmlBody,
        name: 'AutoGlancis'
      });
    }

    // ============ BUSINESS COPY (for you) ============
    const businessHtmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 580px; margin: 0 auto; background-color: #f5f5f5; padding: 20px;">
        <div style="background: white; border-radius: 8px; padding: 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; max-width: 540px; margin: 0 auto;">
          <div style="background: #1a73e8; padding: 20px 24px; color: white;">
            <div style="font-size: 16px; font-weight: 500;">Jauna rezervācija no mājaslapas</div>
          </div>
          <div style="padding: 0 24px 24px;">
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Ko</div>
              <div style="font-size: 14px; color: #202124;">${data.serviceName} (${durationHours} stundas)</div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Kad</div>
              <div style="font-size: 14px; color: #202124;">${dayName}, ${dateStr} plkst. ${data.time}</div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Klients</div>
              <div style="font-size: 14px; color: #202124; line-height: 1.5;">
                <strong>${data.name}</strong><br>
                <span style="color:#5f6368;">${data.phone}</span><br>
                <span style="color:#5f6368;">${data.email || '—'}</span>
              </div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Kur</div>
              <div style="font-size: 14px; color: #202124;">${data.location}</div>
            </div>
            <div style="padding: 16px 0; border-bottom: 1px solid #f1f3f4;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 4px; font-weight: 500;">Papildu piezīmes</div>
              <div style="font-size: 14px; color: #202124;">${data.notes || '—'}</div>
            </div>
            <div style="padding: 16px 0;">
              <div style="font-size: 12px; color: #5f6368; margin-bottom: 8px; font-weight: 500;">Lietotņu statuss</div>
              <div style="display:flex; align-items:center; gap:8px;">
                <span style="background:#34a853; color:white; font-size:12px; padding:2px 6px; border-radius:3px;">✓</span>
                <span style="font-size:14px; color:#202124;">Google Calendar</span>
              </div>
            </div>
          </div>
        </div>
        <p style="text-align:center; font-size:12px; color:#5f6368; margin-top:16px;">AutoGlancis • Ventspils • +371 2568 9885</p>
      </div>
    `;

    MailApp.sendEmail({
      to: CONFIG.BUSINESS_EMAIL,
      subject: `[Jauna rezervācija] ${data.serviceName} – ${data.name}`,
      htmlBody: businessHtmlBody,
      name: 'AutoGlancis Booking'
    });

  } catch (err) {
    console.error('Email sending failed:', err);
  }
}

function jsonResponse(obj) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
  var out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    out.setHeaders(headers);
  } catch (e) {}
  return out;
}

/**
 * Test function - run manually in the Apps Script editor
 */
function testCreateBooking() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        serviceKey: 'auto-salona-tirisana-60',
        serviceName: 'Standarta salona tīrīšana',
        date: '2026-06-20',
        time: '09:00',
        durationHours: 4,
        name: 'TEST Klients',
        phone: '+371 1234 5678',
        email: 'autoglancis@gmail.com',  // Using your business email for easier testing (you'll receive both client + business copies)
        location: 'Ventspils, TEST iela 1',
        notes: 'Šī ir TEST rezervācija. Droši dzēsiet.'
      })
    }
  };

  const result = doPost(fakeEvent);
  console.log(result.getContent());
}

/**
 * Simple test to force MailApp authorization
 */
function testSendEmail() {
  MailApp.sendEmail({
    to: "autoglancis@gmail.com",
    subject: "TEST - AutoGlancis Email Check",
    htmlBody: "<b>This is a test email from the booking system.</b><br>If you receive this, email sending works."
  });
  console.log("Test email sent to autoglancis@gmail.com");
}
