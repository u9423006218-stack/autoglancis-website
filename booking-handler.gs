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
 * AFTER PASTING THE CODE (CRITICAL FOR NEW FUNCTIONS):
 * - After step 3 (paste), you **MUST** click the Save icon (floppy disk) at the top of the editor.
 * - Then reload the entire page (F5 / Ctrl+R).
 * - Only then will all functions (including the new cancellation ones) appear in the "Select function" dropdown next to the Run ▶ button.
 * - If a function is still missing after save+reload, there is likely a red syntax error at the bottom of the editor — fix it and save again.
 *
 * IMPORTANT NOTES:
 * - The first time someone books it may take 5-10 seconds
 * - All times are handled in Europe/Riga timezone
 * - If you change your Google Calendar settings, re-deploy the script
 *
 * CANCELLATIONS:
 * - New bookings (regular or consultation) send ONLY "new booking was made" confirmation emails to the client AND to the business.
 * - Clients get a "Atcelt šo rezervāciju" button/link in their confirmation email.
 * - Clicking the link prompts a 2-step confirmation page to prevent automated email security scanners/bots from cancelling bookings on delivery.
 * - If the owner manually deletes an event from Google Calendar *while it is still upcoming or in progress*, the 15-min detector (checkForCancelledBookings) detects the early removal and sends proper cancellation notices.
 * - Bookings that simply reach their natural end time are pruned silently by the detector. NO spurious "canceled by me" emails are sent on completion.
 * - Run setupCancellationDetector() once after deploy to install the trigger.
 *   (After pasting/saving/reloading the editor page, it will appear in the function dropdown.)
 *
 * Created for autoglancis.lv - Mobile car detailing (Ventspils)
 *
 * Hosted Stripe Checkout flow (server creates session with correct unit_amount in cents).
 * Return handling + extension noise suppression improved on client.
 */

// =====================================================
// CONFIGURATION
// =====================================================

/**
 * If you don't see functions in the dropdown after pasting:
 * 1. Click the Save button (floppy disk icon) at the top of the editor.
 * 2. Reload this page (F5 / Ctrl+R).
 * Then the function "setupCancellationDetector" will appear in the dropdown next to the Run button.
 */
function RUN_ME_AFTER_PASTE_save_and_refresh_editor_then_select_setupCancellationDetector() {
  console.log('=== AutoGlancis Cancellation Setup ===');
  console.log('Step 1: Click the SAVE icon (floppy disk) at the very top of this editor.');
  console.log('Step 2: Reload the browser tab (press F5 or Ctrl+R).');
  console.log('Step 3: In the "Select function" dropdown (right next to the blue Run button), choose "setupCancellationDetector".');
  console.log('Step 4: Click the Run button. This installs the every-15-minute job that emails clients when you delete events from Google Calendar.');
  console.log('Step 5: (Optional test) Select "checkForCancelledBookings" and Run it manually.');
  console.log('Step 6: Redeploy your Web App after any changes.');
  console.log('The big setup function is also defined below for reference.');
}

/**
 * Main function to install the periodic cancellation detector trigger.
 * After save + reload, select this in the dropdown and Run it once.
 */
function setupCancellationDetector() {
  // Remove old triggers for this function
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkForCancelledBookings') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkForCancelledBookings')
    .timeBased()
    .everyMinutes(15)
    .create();

  console.log('Cancellation detector trigger installed (every 15 min).');
  console.log('You can also run checkForCancelledBookings() manually anytime from the dropdown to test immediately.');
}

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
  'detalizeta-kimiska-apstrade-120': 8,
  'standarta-auto-arpuses-tirisana': 1,
  'auto-arpuses-detalizeta-tirisana-45': 2.5,
  'majaslapu-izstrade-uznemumiem': 1,
  'virs-uz-stundu': 2
};

// Required notice period (in hours) before the service start time.
// Unified to a minimum of 12 hours for ALL services (owner needs preparation time).
const NOTICE_HOURS = {
  // All services now default to 12h. No shorter-notice exceptions.
};

// STRIPE CONFIG - DO NOT commit real secret keys to git or leave in deployed source!
// Best practice: run setupStripeSecretKey() in the editor (after temp-pasting the key), 
// then immediately replace key with placeholder in BOTH places below, Save, and Redeploy Web App.
// The code falls back to PropertiesService + this placeholder only as last resort.
// IMPORTANT: Never commit real secret keys. Use the setup function below (run in editor), then DELETE the key value from this source file + from setup func, then redeploy Web App.
const STRIPE_SECRET_KEY_FALLBACK = 'YOUR_STRIPE_SECRET_KEY_HERE'; // placeholder - real key belongs ONLY in Script Properties (via setupStripeSecretKey) or temporarily here for one-time setup only.

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

    if (params.action === 'createStripeCheckoutSession') {
      // Stripe Checkout (hosted) - creates session server-side for security
      let stripeSecret = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
      if (!stripeSecret) {
        stripeSecret = STRIPE_SECRET_KEY_FALLBACK;
      }
      if (!stripeSecret || stripeSecret.indexOf('YOUR_') === 0 || !stripeSecret.startsWith('sk_')) {
        return createJsonOutputWithHeaders({ 
          success: false, 
          error: 'Stripe nav konfigurēts pareizi (nepareiza atslēga). Lūdzu: 1) Atveriet Apps Script projektu. 2) Izpildiet setupStripeSecretKey() (pagaidām ielīmējiet reālo sk_live_... atslēgu iekšā, izpildiet, tad atgrieziet placeholder). 3) Pārbaudiet Project settings > Script properties. 4) Save + Redeploy Web App. Pēc tam atjauniniet BOOKING_SCRIPT_URL HTML failā ja URL mainījās.' 
        }, headers);
      }

      console.log('Stripe secret loaded (prefix: ' + stripeSecret.substring(0,8) + '..., source: ' + (PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') ? 'PropertiesService' : 'FALLBACK') + ')');

      // Robust amountCents: always integer cents, never float/NaN/string with comma etc.
      let amountCents = Math.round( parseFloat( String( params.amountCents || 0 ) ) );
      if (isNaN(amountCents)) amountCents = 0;
      console.log('createStripeCheckoutSession: received amountCents=' + amountCents + ' (raw=' + params.amountCents + ') for ' + (params.serviceKey || params.serviceName));
      console.log('Stripe form integers being sent: unit_amount=' + Math.round(amountCents) + ', quantity=1 (as string)');
      if (!amountCents || amountCents < 50) {
        return createJsonOutputWithHeaders({ success: false, error: 'Nederīga summa maksājumam (' + amountCents + ' centi). Minimālā ~0,50 €.' }, headers);
      }

      const serviceName = params.serviceName || 'Pakalpojums';
      const description = `Datums: ${params.date} ${params.time} • ${params.name} • ${params.location}`;

      // Build metadata for later verification/booking creation
      const metadata = {
        serviceKey: params.serviceKey || '',
        date: params.date || '',
        time: params.time || '',
        name: params.name || '',
        phone: params.phone || '',
        email: params.email || '',
        location: params.location || '',
        notes: (params.notes || '').substring(0, 400),
        durationHours: params.durationHours || '',
        m2: params.m2 || ''
      };

      // success/cancel can be passed from client or default
      const origin = params.origin || 'https://localhost:8000'; // client can send window.location.origin
      const successUrl = params.successUrl || (origin + '/booking-modal.html?stripe_success=1&session_id={CHECKOUT_SESSION_ID}');
      const cancelUrl = params.cancelUrl || (origin + '/booking-modal.html?stripe_canceled=1');

      const payload = {
        'payment_method_types[]': 'card',
        'line_items[0][price_data][currency]': 'eur',
        'line_items[0][price_data][product_data][name]': serviceName,
        'line_items[0][price_data][product_data][description]': description,
        'line_items[0][price_data][unit_amount]': String( Math.round(amountCents) ),
        'line_items[0][quantity]': '1',
        mode: 'payment',
        success_url: successUrl,
        cancel_url: cancelUrl,
        'metadata[serviceKey]': metadata.serviceKey,
        'metadata[serviceName]': serviceName,
        'metadata[date]': metadata.date,
        'metadata[time]': metadata.time,
        'metadata[name]': metadata.name,
        'metadata[phone]': metadata.phone,
        'metadata[email]': metadata.email,
        'metadata[location]': metadata.location,
        'metadata[notes]': metadata.notes,
        'metadata[durationHours]': metadata.durationHours,
        'metadata[m2]': metadata.m2,
        customer_email: params.email || ''
      };

      try {
        const sessionResponse = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', {
          method: 'post',
          headers: {
            Authorization: 'Bearer ' + stripeSecret,
            'Content-Type': 'application/x-www-form-urlencoded'
          },
          payload: payload,
          muteHttpExceptions: true   // CRITICAL: prevents "Exception ... (use muteHttpExceptions)" and gives full error body for 4xx/5xx
        });

        const responseCode = sessionResponse.getResponseCode();
        const responseText = sessionResponse.getContentText();
        let session;
        try {
          session = JSON.parse(responseText);
        } catch (p) {
          session = {};
        }

        if (responseCode >= 200 && responseCode < 300 && session.id && session.url) {
          return createJsonOutputWithHeaders({ 
            success: true, 
            url: session.url, 
            sessionId: session.id 
          }, headers);
        } else {
          const stripeErr = (session && session.error && session.error.message) ? session.error.message : (responseText || 'HTTP ' + responseCode);
          return createJsonOutputWithHeaders({ 
            success: false, 
            error: 'Stripe API kļūda: ' + stripeErr + ' (HTTP ' + responseCode + ', amountCents used: ' + amountCents + ')'
          }, headers);
        }
      } catch (e) {
        return createJsonOutputWithHeaders({ 
          success: false, 
          error: 'Stripe API kļūda: ' + e.toString() + ' (amountCents: ' + amountCents + ')'
        }, headers);
      }
    }

    if (params.action === 'verifyStripeAndCreateBooking') {
      // Called from client after redirect back from Stripe Checkout success
      let stripeSecret = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
      if (!stripeSecret) {
        stripeSecret = STRIPE_SECRET_KEY_FALLBACK;
      }
      if (!stripeSecret || stripeSecret.indexOf('YOUR_') === 0 || !stripeSecret.startsWith('sk_')) {
        return createJsonOutputWithHeaders({ 
          success: false, 
          error: 'Stripe nav konfigurēts pareizi (nepareiza atslēga). Lūdzu izpildiet setupStripeSecretKey() Apps Script editorā (pagaidām ielīmējiet reālo atslēgu, izpildiet, noņemiet no koda, redeploy). Pēc tam atjauniniet URL booking-modal.html.' 
        }, headers);
      }

      console.log('Stripe secret loaded for verify (prefix: ' + stripeSecret.substring(0,8) + '..., source: ' + (PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') ? 'PropertiesService' : 'FALLBACK') + ')');

      const sessionId = params.sessionId;
      if (!sessionId) {
        return createJsonOutputWithHeaders({ success: false, error: 'Trūkst session ID.' }, headers);
      }

      try {
        // Retrieve the Checkout session to verify payment
        const getUrl = 'https://api.stripe.com/v1/checkout/sessions/' + sessionId;
        const sessionResp = UrlFetchApp.fetch(getUrl, {
          method: 'get',
          headers: { Authorization: 'Bearer ' + stripeSecret },
          muteHttpExceptions: true
        });
        const responseCode = sessionResp.getResponseCode();
        const responseText = sessionResp.getContentText();
        let session;
        try { session = JSON.parse(responseText); } catch (p) { session = {}; }

        if (responseCode !== 200) {
          return createJsonOutputWithHeaders({ success: false, error: 'Stripe sesijas nolasīšana neizdevās (HTTP ' + responseCode + '): ' + responseText }, headers);
        }
        if (session.payment_status !== 'paid') {
          return createJsonOutputWithHeaders({ success: false, error: 'Maksājums nav apstiprināts.' }, headers);
        }

        // Rebuild booking params from metadata
        const meta = session.metadata || {};
        const bookingParams = {
          serviceKey: meta.serviceKey,
          serviceName: meta.serviceName || 'Pakalpojums',
          date: meta.date,
          time: meta.time,
          durationHours: meta.durationHours || 3,
          name: meta.name,
          phone: meta.phone,
          email: meta.email || (session.customer_details ? session.customer_details.email : ''),
          location: meta.location,
          notes: meta.notes,
          m2: meta.m2 || null
        };

        // Call the existing booking creation logic (we'll simulate by calling internal or duplicate minimal)
        // For simplicity, duplicate the core creation here (or refactor later)
        // Use the same startTime parsing etc.

        var startTime;
        var endTime;
        if (bookingParams.date && bookingParams.time) {
          var durationH = parseFloat(bookingParams.durationHours) || 3;
          startTime = parseDateTime(bookingParams.date, bookingParams.time);
          endTime = new Date(startTime.getTime() + durationH * 60 * 60 * 1000);
        } else {
          return createJsonOutputWithHeaders({ success: false, error: 'Trūkst datuma/laika.' }, headers);
        }

        // Defensive 12h notice check (client should have enforced, but payment already happened so we log+proceed if violated; still check conflict)
        var noticeHrsV = Math.max( (NOTICE_HOURS[bookingParams.serviceKey] || 0), 12 );
        var serverNowV = new Date();
        var rigaNowStrV = serverNowV.toLocaleString('en-US', { timeZone: CONFIG.TIMEZONE });
        var rigaNowV = new Date(rigaNowStrV);
        var minAllowedStartV = new Date( rigaNowV.getTime() + noticeHrsV * 60 * 60 * 1000 );
        if (startTime < minAllowedStartV) {
          console.warn('Stripe paid booking startTime is <12h notice window. Proceeding to create event because payment succeeded. start=' + startTime);
          // Do not block paid booking; owner can handle manually if needed.
        }

        if (hasConflict(startTime, endTime)) {
          return createJsonOutputWithHeaders({ success: false, error: 'Šis laiks vairs nav pieejams.' }, headers);
        }

        var title = (bookingParams.serviceName || 'Pakalpojums') + ' - ' + bookingParams.name;
        var desc = 'Klients: ' + bookingParams.name + '\n' +
                   'Tālrunis: ' + bookingParams.phone + '\n' +
                   'E-pasts: ' + (bookingParams.email || '') + '\n' +
                   'Vieta: ' + bookingParams.location + '\n' +
                   'Piezīmes: ' + (bookingParams.notes || '') + '\n' +
                   'Stripe Session: ' + sessionId + ' (apmaksāts)\n' +
                   'BOOKING_SOURCE: stripe';

        var createdEvent = CONFIG.CALENDAR.createEvent(title, startTime, endTime, {
          description: desc,
          location: bookingParams.location
        });

        var eventId = createdEvent ? createdEvent.getId() : '';

        // Track for deletion detection
        if (eventId) {
          var known = getKnownBookings();
          known[eventId] = {
            eventId: eventId,
            email: bookingParams.email,
            name: bookingParams.name,
            phone: bookingParams.phone,
            serviceName: bookingParams.serviceName || 'Pakalpojums',
            date: bookingParams.date,
            time: bookingParams.time,
            location: bookingParams.location,
            notes: bookingParams.notes || '',
            endTs: endTime ? endTime.getTime() : null
          };
          known = pruneOldBookings(known);
          saveKnownBookings(known);
        }

        CacheService.getScriptCache().remove('busy_times_cache');

        var emailData = {
          serviceName: bookingParams.serviceName || 'Pakalpojums',
          name: bookingParams.name,
          phone: bookingParams.phone,
          email: bookingParams.email,
          location: bookingParams.location,
          notes: bookingParams.notes,
          date: bookingParams.date,
          time: bookingParams.time
        };

        try {
          sendBookingConfirmationEmail(emailData, bookingParams.durationHours || 3, startTime, eventId);
        } catch (e) {
          console.error('Email failed after Stripe booking:', e);
        }

        return createJsonOutputWithHeaders({ success: true, paid: true }, headers);

      } catch (e) {
        return createJsonOutputWithHeaders({ success: false, error: 'Stripe verifikācijas kļūda: ' + e.toString() }, headers);
      }
    }

    if (params.action === 'getBusyTimes') {
      var cache = CacheService.getScriptCache();
      var cached = cache.get("busy_times_cache");
      if (cached != null) {
        return createJsonOutputWithHeaders(cached, headers);
      }

      // Reuse the shared getBusyTimes (safe Riga TZ parsing + buffer + extra search window).
      // This keeps busy data consistent between doGet/doPost and avoids duplication.
      var busy = getBusyTimes(params.start, params.end);

      var responseObj = { success: true, busy: busy };
      var payload = JSON.stringify(responseObj);
      cache.put("busy_times_cache", payload, 60);

      return createJsonOutputWithHeaders(responseObj, headers);
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

      // === NOTICE PERIOD CHECK (at least 12 hours in advance for ALL services) ===
      // Owner needs preparation time. Enforced on server to prevent bypass.
      var noticeHrs = Math.max(NOTICE_HOURS[params.serviceKey] || 0, 12);
      var serverNow = new Date();
      // Align "now" to Europe/Riga so comparison with client-provided startTime (parsed as Riga wall time) is fair.
      var rigaNowStr = serverNow.toLocaleString('en-US', { timeZone: CONFIG.TIMEZONE });
      var rigaNow = new Date(rigaNowStr);
      var minAllowedStart = new Date( rigaNow.getTime() + noticeHrs * 60 * 60 * 1000 );
      if (startTime < minAllowedStart) {
        return createJsonOutputWithHeaders({ 
          success: false, 
          error: "Rezervāciju var veikt tikai vismaz 12 stundas iepriekš (sagatavošanās laikam)." 
        }, headers);
      }

      // Special restriction for the 8-hour chemical service:
      // User works 9:00–18:00, so only 9:00 start time is allowed (ends at 17:00).
      // Any later start time would go past 18:00.
      if (params.serviceKey === 'detalizeta-kimiska-apstrade-120') {
        var startHour = startTime.getHours();
        var startMinute = startTime.getMinutes();
        if (!(startHour === 9 && startMinute === 0)) {
          return createJsonOutputWithHeaders({ 
            success: false, 
            error: "Šim pakalpojumam (8 stundas) ir pieejams tikai sākuma laiks 9:00." 
          }, headers);
        }
      }

      // Special restriction for Standarta salona tīrīšana (4 hours):
      // Block late start times so the job finishes by 18:00 (with 1h break).
      // Same rule every day (user works every day).
      if (params.serviceKey === 'auto-salona-tirisana-60') {
        var startHour = startTime.getHours();
        var startMinute = startTime.getMinutes();
        var isLate = (startHour === 15) || (startHour === 16 && startMinute === 30) || (startHour === 18);
        if (isLate) {
          return createJsonOutputWithHeaders({ 
            success: false, 
            error: "Šim pakalpojumam (4 stundas) nav pieejami vēlie laiki (no 15:00)." 
          }, headers);
        }
      }

      // Special restriction for Standarta virsbūves tīrīšana (1 hour):
      // Same rule every day: only block 18:00 (1h job starting at 18:00 finishes at 19:00).
      if (params.serviceKey === 'standarta-auto-arpuses-tirisana') {
        var startHour = startTime.getHours();
        var startMinute = startTime.getMinutes();
        if (startHour === 18 && startMinute === 0) {
          return createJsonOutputWithHeaders({ 
            success: false, 
            error: "Šim pakalpojumam (1 stunda) nav pieejams laiks 18:00." 
          }, headers);
        }
      }

      // Special restriction for Detalizētā virsbūves tīrīšana (2.5 hours):
      // Same rule every day as the 4h service: block late start times (15:00+) 
      // so the job finishes by 18:00 with 1h break.
      if (params.serviceKey === 'auto-arpuses-detalizeta-tirisana-45') {
        var startHour = startTime.getHours();
        var startMinute = startTime.getMinutes();
        var isLate = (startHour === 15) || (startHour === 16 && startMinute === 30) || (startHour === 18);
        if (isLate) {
          return createJsonOutputWithHeaders({ 
            success: false, 
            error: "Šim pakalpojumam (2.5 stundas) nav pieejami vēlie laiki (no 15:00)." 
          }, headers);
        }
      }

      // Special restriction for Mēbeļu ķīmiskā tīrīšana (2 hours):
      // Same rule every day: only block 18:00 (2h job starting at 18:00 finishes at 20:00).
      if (params.serviceKey === 'mebelu-kimiska-tirisana') {
        var startHour = startTime.getHours();
        var startMinute = startTime.getMinutes();
        if (startHour === 18 && startMinute === 0) {
          return createJsonOutputWithHeaders({ 
            success: false, 
            error: "Šim pakalpojumam (2 stundas) nav pieejams laiks 18:00." 
          }, headers);
        }
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
                  "E-pasts: " + (params.email || '') + "\n" +
                  "Vieta: " + params.location + "\n" +
                  "Piezīmes: " + (params.notes || "") + "\n" +
                  "BOOKING_SOURCE: website";

      var createdEvent = cal.createEvent(title, startTime, endTime, {
        description: desc,
        location: params.location
      });

      var eventId = createdEvent ? createdEvent.getId() : '';

      // Track for deletion detection (owner cancels by deleting in Calendar)
      if (eventId) {
        var known = getKnownBookings();
        known[eventId] = {
          eventId: eventId,
          email: params.email,
          name: params.name,
          phone: params.phone,
          serviceName: params.serviceName,
          date: params.date,
          time: params.time,
          location: params.location,
          notes: params.notes || '',
          endTs: endTime ? endTime.getTime() : null
        };
        known = pruneOldBookings(known);
        saveKnownBookings(known);
      }

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
        sendBookingConfirmationEmail(emailData, params.durationHours || 3, startTime, eventId);
      } catch (e) {
        console.error('Email failed after booking:', e);
      }

      return createJsonOutputWithHeaders({ success: true }, headers);
    }

    if (params.action === 'createConsultation') {
      // Timed consultation booking for "Mājaslapu izstrāde uzņēmumiem" (now uses date+time + real Calendar event like other services)
      const name = params.name || 'Nezināms';
      const email = params.email || '';
      const phone = params.phone || '';
      const location = params.location || '';
      const notes = params.notes || '';
      const serviceName = params.serviceName || 'Mājaslapu izstrāde uzņēmumiem';

      var startTime = null;
      var endTime = null;
      var eventId = '';
      var hasSlot = params.date && params.time;

      if (hasSlot) {
        var cal = CalendarApp.getDefaultCalendar();
        var durationH = parseFloat(params.durationHours) || 1;
        startTime = parseDateTime(params.date, params.time);
        endTime = new Date(startTime.getTime() + durationH * 60 * 60 * 1000);

        // 12h advance notice (same as all services)
        var noticeHrs = 12;
        var serverNow = new Date();
        var rigaNowStr = serverNow.toLocaleString('en-US', { timeZone: CONFIG.TIMEZONE });
        var rigaNow = new Date(rigaNowStr);
        var minAllowedStart = new Date(rigaNow.getTime() + noticeHrs * 60 * 60 * 1000);
        if (startTime < minAllowedStart) {
          return createJsonOutputWithHeaders({
            success: false,
            error: "Rezervāciju var veikt tikai vismaz 12 stundas iepriekš (sagatavošanās laikam)."
          }, headers);
        }

        var title = 'Konsultācija: ' + serviceName + ' - ' + name;
        var desc = 'KONZULTĀCIJA MĀJASLAPU IZSTRĀDEI UZŅĒMUMIEM\n\n' +
                   'Klients: ' + name + '\n' +
                   'Tālrunis: ' + phone + '\n' +
                   'E-pasts: ' + (email || '') + '\n' +
                   'Vieta: ' + location + '\n' +
                   'Piezīmes: ' + (notes || '') + '\n' +
                   'BOOKING_SOURCE: website_consultation';

        var createdEvent = cal.createEvent(title, startTime, endTime, {
          description: desc,
          location: location
        });
        eventId = createdEvent ? createdEvent.getId() : '';

        if (eventId) {
          var known = getKnownBookings();
          known[eventId] = {
            eventId: eventId,
            email: email,
            name: name,
            phone: phone,
            serviceName: 'Konsultācija: ' + serviceName,
            date: params.date,
            time: params.time,
            location: location,
            notes: notes || '',
            endTs: endTime ? endTime.getTime() : null
          };
          known = pruneOldBookings(known);
          saveKnownBookings(known);
        }

        CacheService.getScriptCache().remove('busy_times_cache');
      }

      try {
        // Business notification (includes slot if booked)
        var businessBody = 'Jauna konsultācijas rezervācija par mājaslapu izstrādi.\n\n' +
          'Klients: ' + name + '\n' +
          'E-pasts: ' + email + '\n' +
          'Tālrunis: ' + phone + '\n' +
          'Atrašanās vieta: ' + location + '\n' +
          'Piezīmes: ' + (notes || '-') + '\n' +
          'Pakalpojums: ' + serviceName + '\n';
        if (hasSlot && startTime) {
          businessBody += 'Kad: ' + params.date + ' plkst. ' + params.time + ' (Europe/Riga)\n';
        }
        businessBody += '\nLūdzu sagatavojieties konsultācijai.';

        MailApp.sendEmail({
          to: CONFIG.BUSINESS_EMAIL,
          subject: hasSlot ? ('[Konsultācija rezervēta] ' + serviceName + ' - ' + name) : ('[Konsultācija] ' + serviceName + ' - ' + name),
          body: businessBody,
          name: 'AutoGlancis'
        });

        // Client confirmation email - "consultation is booked"
        if (email) {
          var dayNames = ['Svētdiena', 'Pirmdiena', 'Otrdiena', 'Trešdiena', 'Ceturtdiena', 'Piektdiena', 'Sestdiena'];
          var dayName = '';
          var dateStr = params.date || '';
          if (startTime) {
            dayName = dayNames[startTime.getDay()];
            dateStr = startTime.toLocaleDateString('lv-LV', { day: 'numeric', month: 'long', year: 'numeric' });
          }

          var cancelLinkHtml = '';
          var baseUrl = ScriptApp.getService().getUrl();
          if (baseUrl && email && eventId) {
            var cancelUrl = baseUrl + '?action=cancelBooking&eventId=' + encodeURIComponent(eventId);
            cancelLinkHtml = '\n\nAtcelt rezervāciju (vismaz 12h iepriekš): ' + cancelUrl;
          }

          var clientBody = 'Labdien, ' + name + '!\n\n' +
            (hasSlot
              ? ('Jūsu bezmaksas konsultācija par mājaslapu izstrādi uzņēmumiem ir rezervēta!\n\n' +
                 'Kad: ' + (dayName ? dayName + ', ' : '') + dateStr + ' plkst. ' + params.time + ' (Europe/Riga)\n' +
                 'Kur: ' + location + '\n\n' +
                 'Paldies par rezervāciju. Lūdzu esiet pieejams norādītajā laikā. Ja nepieciešams mainīt vai atcelt, izmantojiet saiti zemāk.')
              : ('Paldies par Jūsu pieprasījumu konsultācijai par mājaslapu izstrādi.\n' +
                 'Mēs sazināsimies ar Jums tuvākajā laikā.\n')) +
            '\nJūsu dati:\n' +
            'Vārds: ' + name + '\n' +
            'E-pasts: ' + email + '\n' +
            'Tālrunis: ' + phone + '\n' +
            'Atrašanās vieta: ' + location + '\n' +
            'Piezīmes: ' + (notes || '-') + '\n\n' +
            'Ar cieņu,\nAutoGlancis komanda' +
            cancelLinkHtml;

          MailApp.sendEmail({
            to: email,
            subject: hasSlot
              ? ('Konsultācija rezervēta: ' + (dayName ? dayName + ', ' : '') + dateStr + ' plkst. ' + params.time)
              : 'Paldies par pieprasījumu - AutoGlancis mājaslapu izstrāde',
            body: clientBody,
            name: 'AutoGlancis'
          });
        }
      } catch (e) {
        console.error('Consultation email failed:', e);
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
  // Avoid double-stringify when data is already a JSON string (e.g. from cache hit path).
  // This was causing "Standarta režīms (reālā pieejamība nav ielādēta)" on every cache hit.
  var jsonStr;
  if (typeof data === 'string') {
    jsonStr = data;
  } else {
    jsonStr = JSON.stringify(data);
  }
  var out = ContentService.createTextOutput(jsonStr);
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
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  if (e.parameter.action === 'getBusyTimes') {
    const start = e.parameter.start || '';
    const end = e.parameter.end || '';
    const busy = getBusyTimes(start, end);
    return jsonResponse({ success: true, busy: busy });
  }

  // Client clicks "Cancel" link from confirmation email
  if (e.parameter.action === 'cancelBooking') {
    const eventId = e.parameter.eventId;
    if (!eventId) {
      const errHtml = '<html><head><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{margin:0;font-family:system-ui,sans-serif;background:#0a0a0a;color:#e5e7eb;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}</style></head><body><div>Kļūda: Trūkst rezervācijas ID.</div></body></html>';
      return HtmlService.createHtmlOutput(errHtml);
    }

    try {
      const event = CONFIG.CALENDAR.getEventById(eventId);
      if (!event) {
        const notFound = '<html><head><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{margin:0;font-family:system-ui,sans-serif;background:#0a0a0a;color:#e5e7eb;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}</style></head><body><div>Rezervācija nav atrasta.<br>Iespējams, tā jau ir atcelta vai pagājusi.</div></body></html>';
        return HtmlService.createHtmlOutput(notFound);
      }

      const startTime = event.getStartTime();
      const now = new Date();
      const hoursUntil = (startTime.getTime() - now.getTime()) / (1000 * 60 * 60);

      if (hoursUntil < 12) {
        const tooLate = '<html><head><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{margin:0;font-family:system-ui,sans-serif;background:#0a0a0a;color:#e5e7eb;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}</style></head><body><div>Atcelšana vairs nav iespējama.<br>Atcelšana iespējama tikai vismaz 12 stundas iepriekš.</div></body></html>';
        return HtmlService.createHtmlOutput(tooLate);
      }

      // === 2-STEP CONFIRMATION (CRITICAL BOT PROTECTION) ===
      // Email security scanners / anti-virus bots automatically crawl links in incoming emails via GET.
      // If confirmed !== 'true', we show a confirmation screen so bots do NOT trigger immediate deletion!
      if (e.parameter.confirmed !== 'true') {
        const serviceName = event.getTitle();
        const dateStr = startTime.toLocaleDateString('lv-LV', { day: 'numeric', month: 'long', year: 'numeric' });
        const timeStr = startTime.toLocaleTimeString('lv-LV', { hour: '2-digit', minute: '2-digit' });
        const baseUrl = ScriptApp.getService().getUrl();
        const confirmUrl = baseUrl + '?action=cancelBooking&eventId=' + encodeURIComponent(eventId) + '&confirmed=true';

        const confirmHtml = `
          <html>
            <head>
              <meta name="viewport" content="width=device-width, initial-scale=1">
              <title>Atcelt rezervāciju • AutoGlancis</title>
              <style>
                body { margin:0; font-family: system-ui, -apple-system, sans-serif; background:#0a0a0a; color:#e5e7eb; display:flex; align-items:center; justify-content:center; min-height:100vh; padding:20px; box-sizing:border-box; }
                .card { background:#18181b; border:1px solid #27272a; border-radius:16px; padding:32px 24px; max-width:440px; width:100%; text-align:center; box-shadow:0 8px 30px rgba(0,0,0,0.5); }
                .title { font-size:20px; font-weight:600; margin-bottom:12px; color:#ffffff; }
                .box { background:#111111; border:1px solid #27272a; border-radius:12px; padding:16px; margin:20px 0; text-align:left; font-size:14px; line-height:1.6; }
                .btn { display:inline-block; background:#dc2626; color:white; padding:12px 24px; border-radius:10px; text-decoration:none; font-size:15px; font-weight:600; transition:opacity 0.2s; width:100%; box-sizing:border-box; text-align:center; }
                .btn:hover { opacity:0.9; }
                .back { display:block; margin-top:16px; color:#9ca3af; text-decoration:none; font-size:13px; }
                .back:hover { color:#ffffff; }
              </style>
            </head>
            <body>
              <div class="card">
                <div class="title">Vai tiešām vēlaties atcelt rezervāciju?</div>
                <div class="box">
                  <div style="color:#9ca3af; font-size:12px; font-weight:500;">PAKALPOJUMS</div>
                  <div style="font-weight:600; margin-bottom:10px; color:#ffffff;">${serviceName}</div>
                  <div style="color:#9ca3af; font-size:12px; font-weight:500;">LAIKS</div>
                  <div style="font-weight:600; color:#ffffff;">${dateStr} plkst. ${timeStr}</div>
                </div>
                <a href="${confirmUrl}" class="btn">✕ Jā, atcelt šo rezervāciju</a>
                <a href="https://autoglancis.lv" class="back">← Nē, saglabāt rezervāciju</a>
              </div>
            </body>
          </html>
        `;
        return HtmlService.createHtmlOutput(confirmHtml);
      }

      // User confirmed via secondary click - proceed with actual deletion
      // Get info before delete (for email + known list)
      const desc = event.getDescription() || '';
      let booking = { serviceName: event.getTitle(), date: '', time: '', email: '', name: '', phone: '', location: event.getLocation() || '' };

      // Try to get from known first (best data)
      const known = getKnownBookings();
      if (known[eventId]) {
        booking = Object.assign(booking, known[eventId]);
      } else {
        // Fallback: parse description
        const lines = desc.split('\n');
        lines.forEach(l => {
          if (l.toLowerCase().indexOf('klients:') === 0) booking.name = l.split(':')[1] ? l.split(':')[1].trim() : '';
          if (l.toLowerCase().indexOf('tālrunis:') === 0) booking.phone = l.split(':')[1] ? l.split(':')[1].trim() : '';
          if (l.toLowerCase().indexOf('e-pasts:') === 0) booking.email = l.split(':')[1] ? l.split(':')[1].trim() : '';
        });
      }

      // Delete the event
      event.deleteEvent();

      // Remove from tracking
      if (known[eventId]) {
        delete known[eventId];
        saveKnownBookings(known);
      }
      CacheService.getScriptCache().remove('busy_times_cache');

      // Notify
      if (booking.email) {
        try { sendCancellationEmail(booking, 'client_cancelled'); } catch (e) {}
      }
      try {
        MailApp.sendEmail({
          to: CONFIG.BUSINESS_EMAIL,
          subject: `[Klienta atcelšana] ${booking.serviceName || ''}`,
          body: `Klients atcēla rezervāciju via saiti.\n\n${booking.name} (${booking.email})\n${booking.date || ''} ${booking.time || ''}\nEventId: ${eventId}`
        });
      } catch (e) {}

      const successHtml = `
        <html>
          <head>
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>Rezervācija atcelta • AutoGlancis</title>
            <style>
              body { margin:0; font-family: system-ui, -apple-system, sans-serif; background:#0a0a0a; color:#e5e7eb; display:flex; align-items:center; justify-content:center; min-height:100vh; padding:20px; box-sizing:border-box; text-align:center; }
              .card { background:#18181b; border:1px solid #27272a; border-radius:16px; padding:36px 24px; max-width:440px; width:100%; box-shadow:0 8px 30px rgba(0,0,0,0.5); }
              .icon { font-size:42px; margin-bottom:12px; color:#34d399; }
              .title { font-size:20px; font-weight:600; margin-bottom:8px; color:#ffffff; }
              .desc { color:#9ca3af; font-size:14px; margin-bottom:20px; line-height:1.5; }
              .btn { display:inline-block; background:#0284c7; color:white; padding:10px 20px; border-radius:8px; text-decoration:none; font-size:14px; font-weight:500; }
            </style>
          </head>
          <body>
            <div class="card">
              <div class="icon">✓</div>
              <div class="title">Jūsu rezervācija ir atcelta</div>
              <div class="desc">Rezervācija ir veiksmīgi dzēsta no kalendāra. Paldies, ka paziņojāt!</div>
              <a href="https://autoglancis.lv" class="btn">Atgriezties uz autoglancis.lv</a>
            </div>
          </body>
        </html>`;
      return HtmlService.createHtmlOutput(successHtml);

    } catch (err) {
      const errHtml = '<html><head><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{margin:0;font-family:system-ui,sans-serif;background:#0a0a0a;color:#e5e7eb;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}</style></head><body><div>Kļūda: ' + err.toString() + '</div></body></html>';
      return HtmlService.createHtmlOutput(errHtml);
    }
  }

  return jsonResponse({
    status: 'AutoGlancis Booking API is running',
    instructions: 'Use action=getBusyTimes or POST a booking. For cancellation use the link from your confirmation email.'
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
 * Includes a "Cancel appointment" link protected by a 2-step confirmation screen.
 */
function sendBookingConfirmationEmail(data, durationHours, startTime, eventId) {

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

    // Cancellation link (shown to client)
    let cancelLinkHtml = '';
    const baseUrl = ScriptApp.getService().getUrl();
    if (baseUrl && data.email && eventId) {
      const cancelUrl = baseUrl + '?action=cancelBooking&eventId=' + encodeURIComponent(eventId);
      cancelLinkHtml = `
        <div style="margin-top: 16px; text-align: center;">
          <a href="${cancelUrl}" 
             style="display:inline-block; background:#dc2626; color:white; padding:10px 18px; border-radius:6px; text-decoration:none; font-size:14px; font-weight:500;">
            ✕ Atcelt šo rezervāciju
          </a>
          <div style="font-size:11px; color:#9aa0a6; margin-top:6px;">
            Atcelšana iespējama vismaz 12 stundas iepriekš.
          </div>
        </div>
      `;
    }

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
            ${cancelLinkHtml}
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
        email: 'autoglancis@gmail.com',
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

// =====================================================
// STRIPE SETUP FUNCTIONS
// =====================================================

function setupStripeSecretKey() {
  const secretKey = 'YOUR_STRIPE_SECRET_KEY_HERE';
  if (!secretKey || secretKey.indexOf('YOUR_') === 0) {
    console.error('Replace YOUR_STRIPE_SECRET_KEY_HERE with the real sk_live_... SECRET key (the long one starting sk_) before running this setup function. Do NOT use the pk_live_ publishable key here.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('STRIPE_SECRET_KEY', secretKey);
  console.log('STRIPE_SECRET_KEY has been set successfully in Script Properties (starts with: ' + secretKey.substring(0,8) + '...).');
  console.log('ACTION REQUIRED: Edit this file, put back the placeholder "YOUR_STRIPE_SECRET_KEY_HERE", Save, then Deploy → Manage deployments → Redeploy the Web App.');
}

function setupStripePublishableKey() {
  const pubKey = 'pk_live_51RFtsFJ5RwEYMru3HTTp01PCdPe5C6xux6Err7DBOGFnIqTKnASdX5SzLzLQZTdoxApGqLtLmWUCrR5C5fISXCbf00K77xDN8Z';
  PropertiesService.getScriptProperties().setProperty('STRIPE_PUBLISHABLE_KEY', pubKey);
  console.log('STRIPE_PUBLISHABLE_KEY set (optional for this integration).');
}

function testStripeSecret() {
  const secret = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (secret) {
    console.log('STRIPE_SECRET_KEY is set in Properties (starts with: ' + secret.substring(0, 7) + '...)');
  } else {
    console.log('STRIPE_SECRET_KEY NOT in Properties. Use setupStripeSecretKey() (temp paste key, run, then remove from source + redeploy). FALLBACK is also placeholder.');
  }
}

// =====================================================
// CANCELLATION SYSTEM
// =====================================================

function getKnownBookings() {
  const prop = PropertiesService.getScriptProperties().getProperty('KNOWN_BOOKINGS');
  if (!prop) return {};
  try {
    return JSON.parse(prop);
  } catch (e) {
    return {};
  }
}

function saveKnownBookings(bookings) {
  PropertiesService.getScriptProperties().setProperty('KNOWN_BOOKINGS', JSON.stringify(bookings));
}

function pruneOldBookings(known) {
  const now = new Date();
  const cutoff = now.getTime() - 2 * 24 * 60 * 60 * 1000; // keep 2 days history
  Object.keys(known).forEach(id => {
    const b = known[id];
    if (!b) {
      delete known[id];
      return;
    }
    try {
      if (b.endTs && b.endTs < cutoff) {
        delete known[id];
        return;
      }
      if (b.date) {
        const eventStart = parseDateTime(b.date, b.time || '09:00');
        if (eventStart.getTime() < cutoff) {
          delete known[id];
          return;
        }
      }
    } catch (e) {}
  });
  return known;
}

function checkForCancelledBookings() {
  console.log('[CancelDetector] Starting check...');
  let known = getKnownBookings();
  if (Object.keys(known).length === 0) {
    console.log('[CancelDetector] No known bookings to check.');
    return;
  }

  known = pruneOldBookings(known);

  const now = new Date();
  const queryStart = new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000); // 3 days back
  const searchEnd = new Date(now.getTime() + 45 * 24 * 60 * 60 * 1000);
  let currentEvents;
  try {
    currentEvents = CONFIG.CALENDAR.getEvents(queryStart, searchEnd);
  } catch (e) {
    console.error('[CancelDetector] Failed to fetch events', e);
    return;
  }

  const currentIds = {};
  currentEvents.forEach(ev => {
    currentIds[ev.getId()] = true;
  });

  const cancelled = [];
  Object.keys(known).forEach(id => {
    if (!currentIds[id]) {
      const booking = known[id];
      if (!booking) {
        delete known[id];
        return;
      }

      const bookingEndTs = (booking.endTs && typeof booking.endTs === 'number') ? booking.endTs : 0;
      let shouldBePast = false;

      if (bookingEndTs > 0) {
        shouldBePast = bookingEndTs < now.getTime();
      } else if (booking.date && booking.time) {
        try {
          const start = parseDateTime(booking.date, booking.time || '09:00');
          const syntheticEnd = start.getTime() + (10 * 60 * 60 * 1000);
          shouldBePast = syntheticEnd < now.getTime();
        } catch (e) {
          shouldBePast = false;
        }
      }

      if (shouldBePast) {
        console.log('[CancelDetector] Pruning naturally completed booking (no email): ' + id);
        delete known[id];
        return;
      }

      console.log('[CancelDetector] Detected real early delete (owner) for event ' + id + ' for ' + booking.email);
      cancelled.push(booking);
      try {
        sendCancellationEmail(booking, 'owner_deleted');
      } catch (e) {
        console.error('[CancelDetector] Failed to send cancel email', e);
      }
      delete known[id];
    }
  });

  saveKnownBookings(known);
  console.log('[CancelDetector] Check complete. Real owner-cancellations this run: ' + cancelled.length);
}

function sendCancellationEmail(booking, reason) {
  if (!booking || !booking.email) return;

  const subject = `Atcelšana: ${booking.serviceName || 'Rezervācija'} – ${booking.date} plkst. ${booking.time}`;

  let reasonText = 'Diemžēl jūsu rezervācija tika atcelta.';
  if (reason === 'client_cancelled') {
    reasonText = 'Jūs veiksmīgi atcēlāt šo rezervāciju.';
  } else {
    reasonText = 'Diemžēl mēs bijām spiesti atcelt šo rezervāciju (dzēsts no Google Calendar).';
  }

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 580px; margin: 0 auto; background-color: #f5f5f5; padding: 20px;">
      <div style="background: white; border-radius: 8px; padding: 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; max-width: 540px; margin: 0 auto;">
        <div style="background: #dc2626; padding: 20px 24px; color: white; text-align: center;">
          <div style="font-size: 18px; font-weight: 500;">Rezervācija atcelta</div>
        </div>
        <div style="padding: 24px;">
          <p style="color: #202124; font-size: 14px;">${reasonText}</p>
          <div style="margin: 16px 0; padding: 12px; background: #fef2f2; border-radius: 6px;">
            <strong>${booking.serviceName || ''}</strong><br>
            ${booking.date} plkst. ${booking.time}<br>
            ${booking.location || ''}
          </div>
          <p style="color: #5f6368; font-size: 13px;">Ja vēlaties, varat veikt jaunu rezervāciju mūsu mājaslapā: <a href="https://autoglancis.lv" style="color:#1a73e8;">autoglancis.lv</a></p>
        </div>
      </div>
      <p style="text-align:center; font-size:12px; color:#5f6368; margin-top:16px;">AutoGlancis • Ventspils • +371 2568 9885</p>
    </div>
  `;

  MailApp.sendEmail({
    to: booking.email,
    subject: subject,
    htmlBody: htmlBody,
    name: 'AutoGlancis'
  });

  try {
    MailApp.sendEmail({
      to: CONFIG.BUSINESS_EMAIL,
      subject: `[Atcelšana] ${booking.serviceName || ''} – ${booking.name}`,
      body: `Klienta rezervācija atcelta.\n\nKlients: ${booking.name} (${booking.email}, ${booking.phone})\nKad: ${booking.date} ${booking.time}\nIemesls: ${reason || 'manual delete'}`
    });
  } catch (e) {}
}
