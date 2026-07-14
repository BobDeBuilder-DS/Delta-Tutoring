// ============================================================
// DELTA TUTORING — Full Web App
// Deploy: Extensions > Apps Script > Deploy > New Deployment
// Type: Web App | Execute as: Me | Access: Anyone
// ============================================================

const SPREADSHEET_ID  = "1ArWmkkaEDSGpjQsVh6DSk2wBxlkqe74vw-QxJPBAXcI";
const STUDENT_TAB     = "Accepted Students";
const TUTOR_TAB       = "Accepted Tutors";
const STUDENT_RESPONSES = "Student Responses 1";
const TUTOR_RESPONSES   = "Tutor Responses 1";
const FRIDAY_RESPONSES  = "Friday Session Registrations";
const STUDENT_PREFIX  = "DT-STU-";
const TUTOR_PREFIX    = "DT-TUT-";
const BUSINESS_NAME   = "Delta Tutoring";
const BUSINESS_EMAIL  = "deltatutoring.pta@gmail.com";
const WHATSAPP        = "27793228286";

const FRIDAY_SESSIONS = {
  "1": { date: "24 July 2026", theme: "Foundation Check",    desc: "Know what you don't know before it's too late." },
  "2": { date: "31 August 2026", theme: "Danger Zone Topics", desc: "The topics that fail most students — let's tackle them." },
  "3": { date: "7 August 2026", theme: "Past Paper Sprint",  desc: "Work through real questions under real conditions." },
  "4": { date: "14 August 2026", theme: "Final Push",        desc: "One week out — close every gap before test week." },
  "5": { date: "21 August 2026", theme: "Wrap up",        desc: "Take a rest day" },
};

// ── ROUTING ───────────────────────────────────────────────
function doGet(e) {
  const page = e.parameter.page || "home";
  const id   = (e.parameter.id  || "").trim().toUpperCase();
  const week = e.parameter.week || "1";

  if (page === "student-apply")    return serve(studentApplyPage());
  if (page === "tutor-apply")      return serve(tutorApplyPage());
  if (page === "friday-register")  return serve(fridayRegisterPage(week));
  if (page === "friday-success")   return serve(fridaySuccessPage(week));
  if (page === "portal") {
    if (id) {
      const client = lookupClient(id);
      return serve(client ? portalDashboard(client) : portalLogin("ID not recognised. Please check and try again."));
    }
    return serve(portalLogin(""));
  }
  if (page === "success") return serve(successPage(e.parameter.type||"student", e.parameter.name||""));
  return serve(homePage());
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || "";

    if (action === "student-submit") return json(handleStudentSubmit(data));
    if (action === "tutor-submit")   return json(handleTutorSubmit(data));
    if (action === "friday-submit")  return json(handleFridaySubmit(data));
    if (action === "portal-login") {
      const client = lookupClient((data.id||"").trim().toUpperCase());
      return json(client ? {success:true, client} : {success:false});
    }
    return json({success:false, error:"Unknown action"});
  } catch(err) {
    return json({success:false, error:err.toString()});
  }
}

function serve(html) {
  return HtmlService.createHtmlOutput(html)
    .setTitle("Delta Tutoring")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── FRIDAY REGISTRATION HANDLER ───────────────────────────
function handleFridaySubmit(d) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(FRIDAY_RESPONSES);
  if (!sheet) {
    sheet = ss.insertSheet(FRIDAY_RESPONSES);
    sheet.getRange(1,1,1,5).setValues([["Timestamp","Name","Email","Session","Theme"]]);
  }
  const session = FRIDAY_SESSIONS[d.week] || FRIDAY_SESSIONS["1"];
  sheet.appendRow([new Date(), d.name, d.email, "Week "+d.week+" — "+session.date, session.theme]);
  sendFridayConfirmation(d.name, d.email, d.week, session);
  return {success:true, name:d.name, week:d.week};
}

function sendFridayConfirmation(name, email, week, session) {
  const subject = "You're registered! Delta Tutoring — " + session.theme;
  const html = `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1A1A1A;">
    <div style="border-bottom:3px solid #1D9E75;padding-bottom:14px;margin-bottom:22px;">
      <span style="font-size:26px;color:#1D9E75;font-weight:bold;">&#916;</span>
      <span style="font-size:16px;font-weight:bold;margin-left:8px;">Delta Tutoring</span>
    </div>
    <p style="font-size:15px;">Hi <strong>${name}</strong>,</p>
    <p style="font-size:14px;line-height:1.6;">You are registered for our free Friday study session!</p>
    <div style="background:#F7F7F7;border-left:4px solid #1D9E75;padding:14px 18px;margin:22px 0;border-radius:0 8px 8px 0;">
      <p style="margin:0;font-size:11px;color:#888;text-transform:uppercase;letter-spacing:0.08em;">Session Details</p>
      <p style="margin:6px 0 4px;font-size:18px;font-weight:bold;color:#1D9E75;">Week ${week} — ${session.theme}</p>
      <p style="margin:0;font-size:14px;color:#555;">${session.date}</p>
      <p style="margin:6px 0 0;font-size:13px;color:#888;">${session.desc}</p>
    </div>
    <p style="font-size:14px;line-height:1.6;">We will send you the Google Meet link closer to the session date. Keep an eye on your inbox.</p>
    <p style="font-size:13px;color:#666;">This session is completely free — no strings attached. See you Friday!</p>
    <div style="border-top:1px solid #E5E5E5;margin-top:28px;padding-top:14px;">
      <p style="margin:0;font-size:12px;color:#888;">Kind regards,</p>
      <p style="margin:4px 0 0;font-size:13px;font-weight:bold;">${BUSINESS_NAME}</p>
      <p style="margin:2px 0 0;font-size:12px;color:#1D9E75;">${BUSINESS_EMAIL}</p>
    </div>
  </div>`;
  GmailApp.sendEmail(email, subject,
    `Hi ${name},\n\nYou are registered for Week ${week} — ${session.theme} on ${session.date}.\n\nWe will send the Google Meet link closer to the date.\n\nKind regards,\n${BUSINESS_NAME}`,
    {htmlBody:html, name:BUSINESS_NAME, replyTo:BUSINESS_EMAIL});
}

// ── FRIDAY REGISTRATION PAGE ───────────────────────────────
function fridayRegisterPage(week) {
  const url     = ScriptApp.getService().getUrl();
  const session = FRIDAY_SESSIONS[week] || FRIDAY_SESSIONS["1"];
  const allWeeks = Object.entries(FRIDAY_SESSIONS);

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Register — Delta Tutoring Friday Sessions</title>${css()}</head><body>
${nav("friday")}
<div class="form-wrap">
  <span style="display:inline-block;background:#1D9E75;color:#E1F5EE;font-size:10px;font-weight:700;padding:3px 12px;border-radius:20px;letter-spacing:.06em;text-transform:uppercase;margin-bottom:16px;">100% Free</span>
  <h1 class="form-title">Friday Study Sessions</h1>
  <p class="form-sub">Register for one or all of our free virtual study sessions. We will send you the Google Meet link closer to each session date.</p>

  <div id="err-box" class="err-box"></div>

  <!-- Session selector -->
  <div class="form-section">
    <p class="form-section-title">Select Your Session</p>
    <div style="display:flex;flex-direction:column;gap:10px;margin-bottom:1.5rem;" id="session-cards">
      ${allWeeks.map(([w, s]) => `
      <label style="display:flex;align-items:center;gap:16px;padding:14px 18px;border:1.5px solid ${w===week?'#1D9E75':'var(--bd)'};border-radius:10px;cursor:pointer;background:${w===week?'var(--tl)':'#fff'};transition:all .15s;" onclick="selectWeek('${w}')">
        <input type="radio" name="week" value="${w}" ${w===week?'checked':''} style="accent-color:#1D9E75;"/>
        <div>
          <div style="font-size:13px;font-weight:600;color:#1D9E75;margin-bottom:2px;">Week ${w} — ${s.theme}</div>
          <div style="font-size:12px;color:var(--md);">${s.date}</div>
          <div style="font-size:11px;color:#aaa;margin-top:2px;">${s.desc}</div>
        </div>
      </label>`).join('')}
    </div>
  </div>

  <!-- Personal details -->
  <div class="form-section">
    <p class="form-section-title">Your Details</p>
    <div class="form-group">
      <label class="form-label">Full Name <span style="color:#DC2626;">*</span></label>
      <input type="text" id="f-name" class="form-input" placeholder="Your name and surname"/>
    </div>
    <div class="form-group">
      <label class="form-label">Email Address <span style="color:#DC2626;">*</span></label>
      <input type="email" id="f-email" class="form-input" placeholder="your@email.com"/>
    </div>
  </div>

  <button class="btn primary full" onclick="fSubmit()">Register for Free →</button>
  <p style="font-size:11px;color:#aaa;text-align:center;margin-top:12px;">We will send the Google Meet link to your email before the session.</p>

  <!-- Success -->
  <div id="f-success" style="display:none;">
    <div class="success-wrap">
      <div class="success-delta">Δ</div>
      <h2 class="success-title">You're registered!</h2>
      <p style="font-size:14px;color:var(--md);margin-bottom:8px;">Check your email for confirmation.</p>
      <p style="font-size:13px;color:var(--md);max-width:380px;margin:0 auto 24px;">We will send you the Google Meet link before the session. See you Friday!</p>
      <div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
        <a href="${url}?page=friday-register&week=${parseInt(week)<4?parseInt(week)+1:1}" class="btn primary">Register for another session →</a>
        <a href="${url}" class="btn">Back to home</a>
      </div>
    </div>
  </div>

</div>
${foot()}
<script>
const WEB_URL = '${url}';
let selectedWeek = '${week}';

function selectWeek(w){
  selectedWeek = w;
  document.querySelectorAll('input[name="week"]').forEach(r=>r.checked=r.value===w);
  document.querySelectorAll('#session-cards label').forEach(l=>{
    const isSelected = l.querySelector('input').value === w;
    l.style.borderColor = isSelected ? '#1D9E75' : 'var(--bd)';
    l.style.background  = isSelected ? 'var(--tl)' : '#fff';
  });
}

function showErr(msg){ const e=document.getElementById('err-box'); e.textContent=msg; e.style.display='block'; e.scrollIntoView({behavior:'smooth'}); }
function clearErr(){ document.getElementById('err-box').style.display='none'; }

function fSubmit(){
  clearErr();
  const name  = document.getElementById('f-name').value.trim();
  const email = document.getElementById('f-email').value.trim();
  if(!name){showErr('Please enter your name.');return;}
  if(!email){showErr('Please enter your email address.');return;}

  const btn = event.target;
  btn.textContent='Registering...'; btn.disabled=true;

  fetch(WEB_URL,{method:'POST',body:JSON.stringify({action:'friday-submit',name,email,week:selectedWeek})})
    .then(r=>r.json())
    .then(res=>{
      if(res.success){
        document.querySelector('.form-wrap > *:not(#f-success)') && null;
        document.querySelectorAll('.form-wrap > *:not(#f-success)').forEach(el=>el.style.display='none');
        document.getElementById('f-success').style.display='block';
        window.scrollTo(0,0);
      } else {
        showErr('Something went wrong. Please try again.');
        btn.textContent='Register for Free →'; btn.disabled=false;
      }
    })
    .catch(()=>{showErr('Connection error. Please try again.');btn.textContent='Register for Free →';btn.disabled=false;});
}
</script>
</body></html>`;}

function fridaySuccessPage(week) {
  const url = ScriptApp.getService().getUrl();
  const session = FRIDAY_SESSIONS[week] || FRIDAY_SESSIONS["1"];
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Registered! — Delta Tutoring</title>${css()}</head><body>
${nav("")}
<div style="max-width:480px;margin:80px auto;padding:0 24px;text-align:center;">
  <div style="font-family:'DM Serif Display',serif;font-size:60px;color:#1D9E75;margin-bottom:16px;">Δ</div>
  <h1 style="font-family:'DM Serif Display',serif;font-size:28px;font-weight:400;margin-bottom:8px;">You're registered!</h1>
  <p style="font-size:14px;color:var(--md);margin-bottom:20px;">Week ${week} — ${session.theme}<br>${session.date}</p>
  <p style="font-size:13px;color:var(--md);line-height:1.6;margin-bottom:28px;">Check your inbox for a confirmation email. We will send the Google Meet link closer to the session date.</p>
  <div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
    <a href="${url}?page=friday-register" class="btn primary">Register for another →</a>
    <a href="${url}" class="btn">Back to home</a>
  </div>
</div>
${foot()}</body></html>`;}

// ── HANDLE STUDENT SUBMIT ──────────────────────────────────
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENT_RESPONSES) || ss.insertSheet(STUDENT_RESPONSES);
  const lastRow = Math.max(sheet.getLastRow(), 1);

  if (lastRow === 1 && sheet.getRange(1,1).getValue() === "") {
    sheet.getRange(1,1,1,12).setValues([[
      "Timestamp","Name","Surname","University","Faculty","Degree","Year","Contact","Email",
      "Maths Modules","Stats Modules","Areas of Difficulty","Free Consultation","Application ID"
    ]]);
  }

  const year = new Date().getFullYear();
  const seq  = String(sheet.getLastRow()).padStart(4,"0");
  const appId = STUDENT_PREFIX + year + "-" + seq;

  sheet.appendRow([
    new Date(), d.name, d.surname, d.university, d.faculty, d.degree,
    d.year, d.contact, d.email,
    (d.mathsModules||[]).join(", "),
    (d.statsModules||[]).join(", "),
    d.difficulty||"", d.consultation||"No", appId
  ]);

  sendEmail(d.name, d.email, appId, true);
  return {success:true, appId, name:d.name};
}

function handleTutorSubmit(d) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TUTOR_RESPONSES) || ss.insertSheet(TUTOR_RESPONSES);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,14).setValues([[
      "Timestamp","Name","Surname","University","Faculty","Degree","Year","Contact","Email",
      "Maths Modules","Stats Modules","Tools","Tutoring Experience","Outside Tutoring","Interview Date","Application ID"
    ]]);
  }

  const year = new Date().getFullYear();
  const seq  = String(sheet.getLastRow()).padStart(4,"0");
  const appId = TUTOR_PREFIX + year + "-" + seq;

  sheet.appendRow([
    new Date(), d.name, d.surname, d.university, d.faculty, d.degree,
    d.year, d.contact, d.email,
    (d.mathsModules||[]).join(", "),
    (d.statsModules||[]).join(", "),
    (d.tools||[]).join(", "),
    d.experience||"No", d.outsideTutoring||"No", d.interviewDate||"", appId
  ]);

  sendEmail(d.name, d.email, appId, false);
  return {success:true, appId, name:d.name};
}

// ── EMAIL ──────────────────────────────────────────────────
function sendEmail(name, email, appId, isStudent) {
  const type    = isStudent ? "tutoring programme" : "tutor programme";
  const subject = "Application Received — " + BUSINESS_NAME;
  const html = `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1A1A1A;">
    <div style="border-bottom:3px solid #1D9E75;padding-bottom:14px;margin-bottom:22px;">
      <span style="font-size:26px;color:#1D9E75;font-weight:bold;">&#916;</span>
      <span style="font-size:16px;font-weight:bold;margin-left:8px;">Delta Tutoring</span>
    </div>
    <p style="font-size:15px;">Hi <strong>${name}</strong>,</p>
    <p style="font-size:14px;line-height:1.6;">Thank you for expressing interest in our ${type}. Your application has been successfully received.</p>
    <div style="background:#F7F7F7;border-left:4px solid #1D9E75;padding:14px 18px;margin:22px 0;border-radius:0 8px 8px 0;">
      <p style="margin:0;font-size:11px;color:#888;text-transform:uppercase;letter-spacing:0.08em;">Application ID</p>
      <p style="margin:6px 0 0;font-size:22px;font-weight:bold;color:#1D9E75;letter-spacing:0.05em;">${appId}</p>
    </div>
    <p style="font-size:14px;line-height:1.6;">Our team will review your application and contact you within <strong>24–48 hours</strong>.</p>
    <p style="font-size:13px;color:#666;">Keep your Application ID safe.</p>
    <div style="border-top:1px solid #E5E5E5;margin-top:28px;padding-top:14px;">
      <p style="margin:0;font-size:12px;color:#888;">Kind regards,</p>
      <p style="margin:4px 0 0;font-size:13px;font-weight:bold;">${BUSINESS_NAME}</p>
      <p style="margin:2px 0 0;font-size:12px;color:#1D9E75;">${BUSINESS_EMAIL}</p>
    </div>
  </div>`;
  GmailApp.sendEmail(email, subject,
    `Hi ${name},\n\nApplication received.\nApplication ID: ${appId}\n\nWe will contact you within 24-48 hours.\n\nKind regards,\n${BUSINESS_NAME}`,
    {htmlBody:html, name:BUSINESS_NAME, replyTo:BUSINESS_EMAIL});
}

// ── DB LOOKUP ─────────────────────────────────────────────
function lookupClient(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sSheet = ss.getSheetByName(STUDENT_TAB);
  if (sSheet) {
    const data = sSheet.getDataRange().getValues();
    const hdrs = data[0].map(h=>h.toString().trim());
    for (let i=1;i<data.length;i++) {
      const r={};hdrs.forEach((h,j)=>r[h]=data[i][j]);
      if ((r["Application ID"]||"").toString().trim().toUpperCase()===id && r["Status"]==="Active") {
        return {type:"student",id:r["Application ID"],name:r["Full Name"],email:r["Email"],
          package:r["Package"],modules:(r["Modules"]||"").toString().split(",").map(m=>m.trim()).filter(Boolean),
          tutor:r["Tutor Assigned"],month:r["Month"],status:r["Status"]};
      }
    }
  }
  const tSheet = ss.getSheetByName(TUTOR_TAB);
  if (tSheet) {
    const data = tSheet.getDataRange().getValues();
    const hdrs = data[0].map(h=>h.toString().trim());
    for (let i=1;i<data.length;i++) {
      const r={};hdrs.forEach((h,j)=>r[h]=data[i][j]);
      if ((r["Application ID"]||"").toString().trim().toUpperCase()===id && r["Status"]==="Active") {
        return {type:"tutor",id:r["Application ID"],name:r["Full Name"],email:r["Email"],
          modules:(r["Modules"]||"").toString().split(",").map(m=>m.trim()).filter(Boolean),
          studentsAssigned:r["Students Assigned"],maxHours:r["Max Hours"],month:r["Month"],status:r["Status"]};
      }
    }
  }
  return null;
}

// ── SHARED CSS ────────────────────────────────────────────
function css() { return `
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=Inter:wght@400;500;600&display=swap" rel="stylesheet"/>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
:root{--t:#1D9E75;--td:#0F6E56;--tl:#E1F5EE;--dk:#1A1A1A;--md:#555;--lt:#F7F7F7;--bd:#E5E5E5;--r:10px;}
body{font-family:'Inter',sans-serif;color:var(--dk);background:#fff;min-height:100vh;}
nav{background:var(--dk);padding:16px 32px;display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid var(--t);position:sticky;top:0;z-index:100;}
.logo{font-family:'DM Serif Display',serif;font-size:20px;color:#fff;text-decoration:none;}
.logo span{color:var(--t);}
.nav-links{display:flex;gap:20px;align-items:center;}
.nav-links a{font-size:13px;color:#aaa;text-decoration:none;}
.nav-links a:hover{color:#fff;}
.nav-btn{font-size:12px;font-weight:600;color:var(--tl);background:var(--t);padding:7px 16px;border-radius:8px;text-decoration:none;}
.wrap{max-width:720px;margin:0 auto;padding:0 24px;}
.section{padding:2.5rem 0;border-bottom:0.5px solid var(--bd);}
.section-title{font-family:'DM Serif Display',serif;font-size:28px;font-weight:400;margin-bottom:1.5rem;}
.btn{font-size:13px;font-weight:500;padding:10px 22px;border-radius:var(--r);border:0.5px solid var(--bd);background:transparent;color:var(--dk);cursor:pointer;text-decoration:none;display:inline-block;}
.btn:hover{background:var(--lt);}
.btn.primary{background:var(--t);color:var(--tl);border-color:var(--t);}
.btn.primary:hover{background:var(--td);}
.btn.full{width:100%;text-align:center;display:block;}
.card{background:#fff;border:0.5px solid var(--bd);border-radius:12px;padding:1.25rem;}
.grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px;}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}
.grid4{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;}
.stat{background:var(--lt);border-radius:var(--r);padding:1rem;text-align:center;}
.stat-num{font-family:'DM Serif Display',serif;font-size:30px;color:var(--t);}
.stat-lbl{font-size:12px;color:var(--md);margin-top:2px;}
.pkg{background:var(--lt);border:0.5px solid var(--bd);border-radius:12px;padding:1.25rem;display:flex;flex-direction:column;}
.pkg.featured{background:#fff;border:2px solid var(--t);}
.pkg.custom{background:#fff;border:2px dashed var(--bd);}
.pkg-badge{display:inline-block;font-size:11px;background:var(--tl);color:var(--td);padding:3px 10px;border-radius:20px;margin-bottom:.75rem;align-self:flex-start;}
.pkg-badge.muted{background:var(--lt);color:var(--md);}
.pkg-hrs{font-family:'DM Serif Display',serif;font-size:34px;line-height:1;margin-bottom:.2rem;}
.pkg-name{font-size:12px;color:var(--md);margin-bottom:.75rem;}
.pkg-price{font-size:22px;font-weight:600;margin-bottom:1rem;}
.pkg-line{font-size:12px;color:var(--md);padding:4px 0;border-top:0.5px solid var(--bd);display:flex;align-items:center;gap:6px;}
.pkg-line::before{content:"✓";color:var(--t);}
.notice{background:var(--lt);border-left:3px solid var(--t);border-radius:0 8px 8px 0;padding:.75rem 1rem;font-size:13px;color:var(--md);line-height:1.65;margin-top:1rem;}
.step{display:flex;gap:16px;padding:1rem 0;border-bottom:0.5px solid var(--bd);}
.step:last-child{border-bottom:none;}
.step-num{font-family:'DM Serif Display',serif;font-size:22px;color:var(--t);min-width:28px;}
.step-title{font-size:14px;font-weight:600;margin-bottom:.25rem;}
.step-desc{font-size:13px;color:var(--md);line-height:1.65;}
footer{padding:2rem 0;text-align:center;border-top:0.5px solid var(--bd);background:var(--lt);}
.footer-name{font-family:'DM Serif Display',serif;font-size:18px;margin-bottom:.5rem;}
.footer-fine{font-size:11px;color:#aaa;margin-top:.5rem;}
/* FORM STYLES */
.form-wrap{max-width:600px;margin:40px auto;padding:0 24px 60px;}
.form-title{font-family:'DM Serif Display',serif;font-size:32px;font-weight:400;margin-bottom:.5rem;}
.form-sub{font-size:13px;color:var(--md);margin-bottom:2rem;line-height:1.6;}
.form-section{margin-bottom:2rem;}
.form-section-title{font-size:11px;font-weight:600;color:var(--md);text-transform:uppercase;letter-spacing:.1em;margin-bottom:1rem;padding-bottom:8px;border-bottom:2px solid var(--t);}
.form-group{margin-bottom:1.25rem;}
.form-label{font-size:13px;font-weight:500;color:var(--dk);margin-bottom:6px;display:block;}
.form-label span{color:#DC2626;}
.form-input{width:100%;padding:10px 14px;border:1.5px solid var(--bd);border-radius:8px;font-size:14px;color:var(--dk);outline:none;font-family:'Inter',sans-serif;background:#fff;transition:border-color .15s;}
.form-input:focus{border-color:var(--t);}
.form-select{width:100%;padding:10px 14px;border:1.5px solid var(--bd);border-radius:8px;font-size:14px;color:var(--dk);outline:none;background:#fff;cursor:pointer;}
.form-select:focus{border-color:var(--t);}
.form-textarea{width:100%;padding:10px 14px;border:1.5px solid var(--bd);border-radius:8px;font-size:14px;color:var(--dk);outline:none;font-family:'Inter',sans-serif;min-height:100px;resize:vertical;}
.form-textarea:focus{border-color:var(--t);}
.checkbox-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:8px;}
.checkbox-item{display:flex;align-items:flex-start;gap:10px;padding:10px 14px;border:1.5px solid var(--bd);border-radius:8px;cursor:pointer;transition:all .15s;background:#fff;}
.checkbox-item:hover{border-color:var(--t);background:var(--tl);}
.checkbox-item input{margin-top:2px;accent-color:var(--t);cursor:pointer;flex-shrink:0;}
.checkbox-item label{font-size:13px;color:var(--dk);cursor:pointer;line-height:1.4;}
.radio-group{display:flex;flex-direction:column;gap:8px;}
.radio-item{display:flex;align-items:center;gap:10px;padding:10px 14px;border:1.5px solid var(--bd);border-radius:8px;cursor:pointer;transition:all .15s;}
.radio-item:hover{border-color:var(--t);background:var(--tl);}
.radio-item input{accent-color:var(--t);cursor:pointer;}
.radio-item label{font-size:13px;color:var(--dk);cursor:pointer;}
.form-step{display:none;}
.form-step.active{display:block;}
.step-indicator{display:flex;gap:8px;margin-bottom:2rem;flex-wrap:wrap;}
.step-dot{width:8px;height:8px;border-radius:50%;background:var(--bd);transition:background .2s;}
.step-dot.done{background:var(--t);}
.step-dot.current{background:var(--t);width:24px;border-radius:4px;}
.nav-btns{display:flex;gap:12px;margin-top:2rem;}
.err-box{background:#FEE2E2;border:1px solid #FECACA;border-radius:8px;padding:10px 14px;font-size:13px;color:#DC2626;margin-bottom:16px;display:none;}
.success-wrap{text-align:center;padding:60px 24px;}
.success-delta{font-family:'DM Serif Display',serif;font-size:60px;color:var(--t);}
.success-title{font-family:'DM Serif Display',serif;font-size:32px;margin:16px 0 8px;}
.success-id{font-family:monospace;font-size:20px;color:var(--t);background:var(--tl);padding:12px 24px;border-radius:8px;display:inline-block;margin:16px 0;}
/* PORTAL */
.portal-wrap{max-width:720px;margin:0 auto;padding:32px 24px;}
.welcome-card{background:var(--t);border-radius:12px;padding:22px 26px;margin-bottom:22px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;}
.welcome-name{font-family:'DM Serif Display',serif;font-size:22px;color:#fff;margin-bottom:4px;}
.welcome-sub{font-size:12px;color:var(--tl);letter-spacing:.06em;}
.welcome-id{font-family:monospace;font-size:13px;color:var(--tl);background:rgba(0,0,0,0.2);padding:6px 12px;border-radius:6px;}
.chip{background:var(--tl);border-radius:20px;padding:6px 14px;font-size:12px;color:var(--t);font-weight:500;}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:22px;}
.docs-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:32px;}
.doc-card{background:#fff;border:0.5px solid var(--bd);border-radius:10px;padding:16px 18px;cursor:pointer;transition:all .15s;}
.doc-card:hover{border-color:var(--t);box-shadow:0 0 0 2px var(--tl);}
.doc-icon{font-size:22px;margin-bottom:8px;}
.doc-title{font-size:13px;font-weight:600;color:var(--dk);margin-bottom:4px;}
.doc-desc{font-size:11px;color:var(--md);line-height:1.5;}
.modal-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:200;align-items:center;justify-content:center;padding:24px;}
.modal-box{background:#fff;border-radius:14px;width:100%;max-width:600px;max-height:80vh;overflow:auto;}
.modal-head{background:var(--t);padding:16px 24px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;border-radius:14px 14px 0 0;}
.modal-title{font-size:15px;font-weight:600;color:#fff;}
.modal-close{background:rgba(255,255,255,0.2);border:none;color:#fff;border-radius:6px;padding:4px 10px;cursor:pointer;}
.modal-body{padding:24px;font-size:13px;line-height:1.7;white-space:pre-wrap;}
@media(max-width:600px){.grid3{grid-template-columns:repeat(2,1fr);}.nav-links{display:none;}.checkbox-grid{grid-template-columns:1fr;}}
</style>`; }

function nav(activePage) {
  const url = ScriptApp.getService().getUrl();
  return `<nav>
  <a href="${url}" class="logo"><span>Δ</span> Delta Tutoring</a>
  <div class="nav-links">
    <a href="${url}#about">About</a>
    <a href="${url}#packages">Packages</a>
    <a href="${url}#process">How it works</a>
    <a href="${url}?page=portal" class="nav-btn">Client Portal</a>
  </div>
</nav>`; }

function foot() {
  return `<footer>
  <div class="wrap">
    <p class="footer-name">Δ Delta Tutoring</p>
    <p style="font-size:13px;color:var(--md);">Pretoria &nbsp;·&nbsp; <a href="https://wa.me/${WHATSAPP}" style="color:var(--t);">WhatsApp us</a></p>
    <p class="footer-fine">Specialist university mathematics & statistics · South Africa</p>
  </div>
</footer>`; }

// ── HOME PAGE ─────────────────────────────────────────────
function homePage() {
  const url = ScriptApp.getService().getUrl();
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Delta Tutoring</title>${css()}</head><body>
${nav("home")}
<div class="wrap">
  <div class="section" style="padding-top:4rem;">
    <span style="font-family:'DM Serif Display',serif;font-size:52px;color:var(--t);display:block;margin-bottom:8px;">Δ</span>
    <p style="font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--md);margin-bottom:1.25rem;">Delta Tutoring · University mathematics & statistics</p>
    <h1 style="font-family:'DM Serif Display',serif;font-size:42px;line-height:1.15;font-weight:400;margin-bottom:1.25rem;">Where <em style="font-style:italic;color:var(--t);">change</em> begins<br>with commitment</h1>
    <p style="font-size:15px;color:var(--md);line-height:1.75;max-width:540px;margin-bottom:2rem;">Specialist one-on-one tutoring for undergraduate quantitative modules. We vet both our tutors and our students,because real progress requires commitment on both sides.</p>
    <div style="display:flex;gap:12px;flex-wrap:wrap;">
      <a href="${url}?page=student-apply" class="btn primary">Apply as a student ↗</a>
      <a href="${url}?page=tutor-apply" class="btn">Join as a tutor ↗</a>
    </div>
  </div>

  <div class="section" id="about">
    <div class="grid3">
      <div class="stat"><div class="stat-num">1:1</div><div class="stat-lbl">Sessions only</div></div>
      <div class="stat"><div class="stat-num">20hrs</div><div class="stat-lbl">Max per tutor monthly</div></div>
      <div class="stat"><div class="stat-num">100%</div><div class="stat-lbl">Quant modules</div></div>
    </div>
  </div>

  <div class="section">
    <h2 class="section-title">What makes Delta different</h2>
    <div class="grid2">
      <div class="card"><div style="font-size:20px;color:var(--t);margin-bottom:.75rem;">✓</div><p style="font-size:14px;font-weight:600;margin-bottom:.4rem;">Two-sided vetting</p><p style="font-size:13px;color:var(--md);line-height:1.65;">We assess both tutors and students before any match is made. Tutors must demonstrate they can teach, not just solve.</p></div>
      <div class="card"><div style="font-size:20px;color:var(--t);margin-bottom:.75rem;">📚</div><p style="font-size:14px;font-weight:600;margin-bottom:.4rem;">Subject-matched pairs</p><p style="font-size:13px;color:var(--md);line-height:1.65;">Your tutor has studied the exact module you need help with. Engineering maths, statistics, calculus. Matched precisely!</p></div>
      <div class="card"><div style="font-size:20px;color:var(--t);margin-bottom:.75rem;">📈</div><p style="font-size:14px;font-weight:600;margin-bottom:.4rem;">Progress over marks</p><p style="font-size:13px;color:var(--md);line-height:1.65;">We track growth from your baseline. Delta measures change (that's the point).</p></div>
      <div class="card"><div style="font-size:20px;color:var(--t);margin-bottom:.75rem;">🛡</div><p style="font-size:14px;font-weight:600;margin-bottom:.4rem;">Tutor protection</p><p style="font-size:13px;color:var(--md);line-height:1.65;">Our tutors are capped at 20 hours per month. They are students too. Their academics come first.</p></div>
    </div>
  </div>

  <div class="section" id="packages">
    <h2 class="section-title">Monthly packages</h2>
    <p style="font-size:13px;color:var(--md);margin:-0.75rem 0 1.5rem;">Valid for the current calendar month only. Unused hours do not carry over.</p>
    <div class="grid4">
      <div class="pkg"><p class="pkg-hrs">5</p><p class="pkg-name">hours · Starter</p><p class="pkg-price">R1,500</p><div class="pkg-line">Ideal for 1 to 2 modules</div><div class="pkg-line">Tutor matching</div><div class="pkg-line">Progress report</div></div>
      <div class="pkg"><p class="pkg-hrs">10</p><p class="pkg-name">hours · Standard</p><p class="pkg-price">R3,000</p><div class="pkg-line">Ideal for 2 to 3 modules</div><div class="pkg-line">Tutor matching</div><div class="pkg-line">Progress report</div></div>
      <div class="pkg"><p class="pkg-hrs">15</p><p class="pkg-name">hours · Committed</p><p class="pkg-price">R4,000</p><div class="pkg-line">Ideal for 3 to 4 modules</div><div class="pkg-line">Tutor matching</div><div class="pkg-line">Progress report</div></div>
      <div class="pkg featured"><span class="pkg-badge">Best value</span><p class="pkg-hrs">20</p><p class="pkg-name">hours · Full month</p><p class="pkg-price">R5,000</p><div class="pkg-line">Ideal for 5 modules</div><div class="pkg-line">Tutor matching</div><div class="pkg-line">Progress report</div><div class="pkg-line">Mid-month check-in</div></div>
    </div>
    <div class="pkg custom" style="margin-top:12px;"><span class="pkg-badge muted">By application only</span><p class="pkg-hrs">20+</p><p class="pkg-name">hours · Custom package</p><p class="pkg-price" style="font-size:18px;">Fee on request</p><div class="pkg-line">For students with 6 to 9 modules</div><div class="pkg-line">Hours and fee agreed at intake</div><div class="pkg-line">Bi-weekly payment plan available</div></div>
    <div class="notice">All packages are month-to-month. Bi-weekly payment plans available — first instalment before session one, second by the 15th.</div>
  </div>

  <div class="section" id="process">
    <h2 class="section-title">How it works</h2>
    <div>
      <div class="step"><div class="step-num">01</div><div><p class="step-title">You apply</p><p class="step-desc">Fill in our application form. Tell us your module, faculty, year level, and where you are struggling.</p></div></div>
      <div class="step"><div class="step-num">02</div><div><p class="step-title">We assess fit</p><p class="step-desc">A brief intake conversation to understand your needs and commitment. Not every applicant is accepted, and that is by design.</p></div></div>
      <div class="step"><div class="step-num">03</div><div><p class="step-title">We match you</p><p class="step-desc">You are paired with a tutor who has studied your exact module.</p></div></div>
      <div class="step"><div class="step-num">04</div><div><p class="step-title">You commit and grow</p><p class="step-desc">Sessions begin. We track your Δ, your change from baseline, monthly. Accountability runs both ways.</p></div></div>
    </div>
  </div>

  <div class="section" id="portal" style="background:var(--lt);margin:0 -24px;padding:2.5rem 24px;">
    <h2 class="section-title">Client Portal</h2>
    <p style="font-size:13px;color:var(--md);margin:-0.75rem 0 1.5rem;">Already a student or tutor? Access your package documents with your Application ID.</p>
    <div style="max-width:440px;">
      <div id="portal-err" class="err-box"></div>
      <label style="font-size:12px;font-weight:500;color:var(--md);margin-bottom:6px;display:block;text-transform:uppercase;letter-spacing:.06em;">Application ID</label>
      <input type="text" id="portalId" class="form-input" placeholder="e.g. DT-STU-2026-0001" onkeydown="if(event.key==='Enter')portalGo()" style="margin-bottom:12px;"/>
      <a href="#" onclick="portalGo();return false;" class="btn primary">Access My Package →</a>
      <p style="font-size:11px;color:#aaa;margin-top:12px;">Your ID was emailed when your application was accepted. Contact hello@deltatutoring.co.za for help.</p>
    </div>
  </div>
</div>
${foot()}
<script>
const URL = '${url}';
function portalGo(){
  const id = document.getElementById('portalId').value.trim().toUpperCase();
  const err = document.getElementById('portal-err');
  if(!id){err.textContent='Please enter your Application ID.';err.style.display='block';return;}
  err.style.display='none';
  window.location.href = URL+'?page=portal&id='+encodeURIComponent(id);
}
</script>
</body></html>`; }

// ── STUDENT APPLICATION ────────────────────────────────────
function studentApplyPage() {
  const url = ScriptApp.getService().getUrl();
  const MATHS = {
    "First Year": ["WTW 123 - Numerical Analysis","WTW 124 - Mathematics","WTW 143 - Calculus","WTW 144 - Mathematics","WTW 146 - Linear Algebra","WTW 147 - Foundational Mathematics","WTW 148 - Calculus","WTW 153 - Calculus","WTW 154 - Mathematics","WTW 162 - Dynamical Processes","WTW 164 - Mathematics","WTW 165 - Mathematics","JPO 126 - Additional Mathematics"],
    "Second Year": ["WTW 220 - Analysis","WTW 221 - Linear Algebra","WTW 224 - Technical Analysis","WTW 238 - Mathematics","WTW 248 - Vector Analysis","WTW 263 - Numerical Methods","WTW 264 - Differential Methods","WTW 285 - Discrete Structures"],
    "Third Year": ["WTW 320 - Complex Analysis","WTW 364 - Financial Engineering","WTW 383 - Numerical Analysis","WTW 387 - Continuum Mechanics","WTW 389 - Geometry"]
  };
  const STATS = {
    "First Year": ["BME 120 - Biometry","FBS 112 - Financial Management","STC 122 - Statistics","STC 147 - Foundational Statistics","STC 157 - Mathematical Concepts for Statistics","STK 120 - Statistics","WST 121 - Mathematical Statistics"],
    "Second Year": ["BES 220 - Engineering Statistics","GIS 220 - Geographic Data Analysis","GJI 220 - Statistics for Teachers","STK 220 - Statistics","WST 220 - Mathematical Statistics"],
    "Third Year": ["ESC 320 - Stochastic Communications System","GIS 320 - Spatial Analysis","STK 320 - Statistics","WST 321 - Time Series Analysis","WST 322 - Actuarial Statistics"]
  };

  const mathsData = JSON.stringify(MATHS);
  const statsData = JSON.stringify(STATS);

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Student Application — Delta Tutoring</title>${css()}</head><body>
${nav("apply")}
<div class="form-wrap">
  <h1 class="form-title">Student Application</h1>
  <p class="form-sub">This application serves the purpose of Delta's student intake process. Fields marked <span style="color:#DC2626;">*</span> are required.</p>

  <div id="err-box" class="err-box"></div>

  <div class="step-indicator" id="dots"></div>

  <!-- STEP 1: Personal Details -->
  <div class="form-step active" id="step-1">
    <div class="form-section">
      <p class="form-section-title">Personal Details</p>
      <div class="form-group"><label class="form-label">Name(s) <span>*</span></label><input type="text" id="s-name" class="form-input" placeholder="Your first name(s)"/></div>
      <div class="form-group"><label class="form-label">Surname <span>*</span></label><input type="text" id="s-surname" class="form-input" placeholder="Your surname"/></div>
      <div class="form-group"><label class="form-label">University <span>*</span></label><input type="text" id="s-university" class="form-input" placeholder="e.g. University of Pretoria"/></div>
      <div class="form-group"><label class="form-label">Faculty <span>*</span></label>
        <select id="s-faculty" class="form-select">
          <option value="">Select faculty</option>
          <option>EBIT</option><option>Education</option><option>EMS</option>
          <option>Humanities</option><option>NS</option><option>Veterinary</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">Degree / Programme <span>*</span></label><input type="text" id="s-degree" class="form-input" placeholder="e.g. BSc Computer Science"/></div>
      <div class="form-group"><label class="form-label">Contact Number <span>*</span></label><input type="tel" id="s-contact" class="form-input" placeholder="e.g. 0821234567"/></div>
      <div class="form-group"><label class="form-label">Email Address <span>*</span></label><input type="email" id="s-email" class="form-input" placeholder="your@email.com"/></div>
    </div>
    <div class="nav-btns"><button class="btn primary" onclick="sNext(1)">Next →</button></div>
  </div>

  <!-- STEP 2: Year of Study -->
  <div class="form-step" id="step-2">
    <div class="form-section">
      <p class="form-section-title">Year of Study</p>
      <p style="font-size:13px;color:var(--md);margin-bottom:1rem;">Select your current year — you will only see modules relevant to your year.</p>
      <div class="radio-group">
        <label class="radio-item"><input type="radio" name="s-year" value="First Year" onchange="sYearChange()"/><label>First Year</label></label>
        <label class="radio-item"><input type="radio" name="s-year" value="Second Year" onchange="sYearChange()"/><label>Second Year</label></label>
        <label class="radio-item"><input type="radio" name="s-year" value="Third Year" onchange="sYearChange()"/><label>Third Year</label></label>
      </div>
    </div>
    <div class="nav-btns">
      <button class="btn" onclick="sBack(2)">← Back</button>
      <button class="btn primary" onclick="sNext(2)">Next →</button>
    </div>
  </div>

  <!-- STEP 3: Modules -->
  <div class="form-step" id="step-3">
    <div class="form-section">
      <p class="form-section-title">Module Selection</p>
      <p style="font-size:13px;color:var(--md);margin-bottom:1.5rem;">Select the modules you need help with. You may select more than one.</p>
      <div class="form-group">
        <label class="form-label">Mathematics Modules</label>
        <div class="checkbox-grid" id="maths-modules"></div>
      </div>
      <div class="form-group" style="margin-top:1.5rem;">
        <label class="form-label">Statistics Modules</label>
        <div class="checkbox-grid" id="stats-modules"></div>
      </div>
    </div>
    <div class="nav-btns">
      <button class="btn" onclick="sBack(3)">← Back</button>
      <button class="btn primary" onclick="sNext(3)">Next →</button>
    </div>
  </div>

  <!-- STEP 4: Areas of Difficulty -->
  <div class="form-step" id="step-4">
    <div class="form-section">
      <p class="form-section-title">Areas of Difficulty</p>
      <div class="form-group">
        <label class="form-label">Which topics, chapters, concepts, assignments, or assessments are you struggling with? <span>*</span></label>
        <textarea id="s-difficulty" class="form-textarea" placeholder="Describe where you are struggling..."></textarea>
      </div>
      <div class="form-group">
        <label class="form-label">Would you like a free consultation to help determine where you struggle?</label>
        <div class="radio-group">
          <label class="radio-item"><input type="radio" name="s-consult" value="Yes"/><label>Yes please</label></label>
          <label class="radio-item"><input type="radio" name="s-consult" value="No"/><label>No, I know where I struggle</label></label>
        </div>
      </div>
    </div>
    <div class="nav-btns">
      <button class="btn" onclick="sBack(4)">← Back</button>
      <button class="btn primary" onclick="sSubmit()">Submit Application →</button>
    </div>
  </div>

  <!-- SUCCESS -->
  <div class="form-step" id="step-success">
    <div class="success-wrap">
      <div class="success-delta">Δ</div>
      <h2 class="success-title">Application Received!</h2>
      <p style="font-size:14px;color:var(--md);margin-bottom:16px;">Your application has been submitted successfully.</p>
      <div class="success-id" id="success-id"></div>
      <p style="font-size:13px;color:var(--md);max-width:400px;margin:0 auto 24px;">Keep this ID safe — you will need it to access your client portal. A confirmation email has been sent to you.</p>
      <p style="font-size:13px;color:var(--md);">Our team will review your application and contact you within <strong>24–48 hours</strong>.</p>
      <div style="margin-top:32px;display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
        <a href="${url}?page=portal" class="btn primary">Go to Client Portal</a>
        <a href="${url}" class="btn">Back to Home</a>
      </div>
    </div>
  </div>

</div>
${foot()}
<script>
const MATHS = ${mathsData};
const STATS = ${statsData};
const WEB_URL = '${url}';
let currentStep = 1;
const totalSteps = 4;

function updateDots(){
  const d = document.getElementById('dots');
  d.innerHTML = '';
  for(let i=1;i<=totalSteps;i++){
    const dot = document.createElement('div');
    dot.className = 'step-dot' + (i<currentStep?' done':i===currentStep?' current':'');
    d.appendChild(dot);
  }
}
updateDots();

function showStep(n){
  document.querySelectorAll('.form-step').forEach(s=>s.classList.remove('active'));
  document.getElementById('step-'+n).classList.add('active');
  window.scrollTo(0,0);
  currentStep = n;
  updateDots();
}

function sYearChange(){
  const yr = document.querySelector('input[name="s-year"]:checked')?.value;
  if(!yr) return;
  const mg = document.getElementById('maths-modules');
  const sg = document.getElementById('stats-modules');
  mg.innerHTML = ''; sg.innerHTML = '';
  (MATHS[yr]||[]).forEach(m=>{
    mg.innerHTML += '<label class="checkbox-item"><input type="checkbox" name="s-maths" value="'+m+'"/><label>'+m+'</label></label>';
  });
  (STATS[yr]||[]).forEach(m=>{
    sg.innerHTML += '<label class="checkbox-item"><input type="checkbox" name="s-stats" value="'+m+'"/><label>'+m+'</label></label>';
  });
}

function showErr(msg){
  const e = document.getElementById('err-box');
  e.textContent = msg; e.style.display='block';
  e.scrollIntoView({behavior:'smooth'});
}
function clearErr(){ document.getElementById('err-box').style.display='none'; }

function sNext(step){
  clearErr();
  if(step===1){
    if(!document.getElementById('s-name').value.trim()){showErr('Please enter your name.');return;}
    if(!document.getElementById('s-surname').value.trim()){showErr('Please enter your surname.');return;}
    if(!document.getElementById('s-university').value.trim()){showErr('Please enter your university.');return;}
    if(!document.getElementById('s-faculty').value){showErr('Please select your faculty.');return;}
    if(!document.getElementById('s-degree').value.trim()){showErr('Please enter your degree.');return;}
    if(!document.getElementById('s-contact').value.trim()){showErr('Please enter your contact number.');return;}
    if(!document.getElementById('s-email').value.trim()){showErr('Please enter your email address.');return;}
  }
  if(step===2){
    if(!document.querySelector('input[name="s-year"]:checked')){showErr('Please select your year of study.');return;}
    sYearChange();
  }
  if(step===3){
    const m = document.querySelectorAll('input[name="s-maths"]:checked');
    const s = document.querySelectorAll('input[name="s-stats"]:checked');
    if(m.length===0 && s.length===0){showErr('Please select at least one module.');return;}
  }
  showStep(step+1);
}

function sBack(step){ clearErr(); showStep(step-1); }

function sSubmit(){
  clearErr();
  const difficulty = document.getElementById('s-difficulty').value.trim();
  if(!difficulty){showErr('Please describe where you are struggling.');return;}

  const btn = event.target;
  btn.textContent = 'Submitting...'; btn.disabled = true;

  const data = {
    action: 'student-submit',
    name: document.getElementById('s-name').value.trim(),
    surname: document.getElementById('s-surname').value.trim(),
    university: document.getElementById('s-university').value.trim(),
    faculty: document.getElementById('s-faculty').value,
    degree: document.getElementById('s-degree').value.trim(),
    year: document.querySelector('input[name="s-year"]:checked')?.value||'',
    contact: document.getElementById('s-contact').value.trim(),
    email: document.getElementById('s-email').value.trim(),
    mathsModules: [...document.querySelectorAll('input[name="s-maths"]:checked')].map(i=>i.value),
    statsModules: [...document.querySelectorAll('input[name="s-stats"]:checked')].map(i=>i.value),
    difficulty: difficulty,
    consultation: document.querySelector('input[name="s-consult"]:checked')?.value||'No',
  };

  fetch(WEB_URL, {method:'POST', body:JSON.stringify(data)})
    .then(r=>r.json())
    .then(res=>{
      if(res.success){
        document.getElementById('success-id').textContent = res.appId;
        document.querySelectorAll('.form-step').forEach(s=>s.classList.remove('active'));
        document.getElementById('step-success').classList.add('active');
        document.getElementById('dots').style.display='none';
        window.scrollTo(0,0);
      } else {
        showErr('Something went wrong. Please try again.');
        btn.textContent='Submit Application →'; btn.disabled=false;
      }
    })
    .catch(()=>{showErr('Connection error. Please try again.');btn.textContent='Submit Application →';btn.disabled=false;});
}
</script>
</body></html>`; }

// ── TUTOR APPLICATION ──────────────────────────────────────
function tutorApplyPage() {
  const url = ScriptApp.getService().getUrl();
  const MATHS = {
    "First Year": ["WTW 123 - Numerical Analysis","WTW 124 - Mathematics","WTW 143 - Calculus","WTW 144 - Mathematics","WTW 146 - Linear Algebra","WTW 147 - Foundational Mathematics","WTW 148 - Calculus","WTW 153 - Calculus","WTW 154 - Mathematics","WTW 162 - Dynamical Processes","WTW 164 - Mathematics","WTW 165 - Mathematics","JPO 126 - Additional Mathematics"],
    "Second Year": ["WTW 220 - Analysis","WTW 221 - Linear Algebra","WTW 224 - Technical Analysis","WTW 238 - Mathematics","WTW 248 - Vector Analysis","WTW 263 - Numerical Methods","WTW 264 - Differential Methods","WTW 285 - Discrete Structures"],
    "Third Year / Final Year": ["WTW 320 - Complex Analysis","WTW 364 - Financial Engineering","WTW 383 - Numerical Analysis","WTW 387 - Continuum Mechanics","WTW 389 - Geometry"],
    "Honours / Masters": ["WTW 320 - Complex Analysis","WTW 364 - Financial Engineering","WTW 383 - Numerical Analysis","WTW 387 - Continuum Mechanics","WTW 389 - Geometry"]
  };
  const STATS = {
    "First Year": ["BME 120 - Biometry","FBS 112 - Financial Management","STC 122 - Statistics","STC 147 - Foundational Statistics","STC 157 - Mathematical Concepts for Statistics","STK 120 - Statistics","WST 121 - Mathematical Statistics"],
    "Second Year": ["BES 220 - Engineering Statistics","GIS 220 - Geographic Data Analysis","GJI 220 - Statistics for Teachers","STK 220 - Statistics","WST 220 - Mathematical Statistics"],
    "Third Year / Final Year": ["ESC 320 - Stochastic Communications System","GIS 320 - Spatial Analysis","STK 320 - Statistics","WST 321 - Time Series Analysis","WST 322 - Actuarial Statistics"],
    "Honours / Masters": ["ESC 320 - Stochastic Communications System","GIS 320 - Spatial Analysis","STK 320 - Statistics","WST 321 - Time Series Analysis","WST 322 - Actuarial Statistics"]
  };

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Tutor Application — Delta Tutoring</title>${css()}</head><body>
${nav("apply")}
<div class="form-wrap">
  <h1 class="form-title">Tutor Application</h1>
  <p class="form-sub">This application serves the purpose of Delta's tutor intake process. Fields marked <span style="color:#DC2626;">*</span> are required.</p>

  <div id="err-box" class="err-box"></div>
  <div class="step-indicator" id="dots"></div>

  <!-- STEP 1: Personal Details -->
  <div class="form-step active" id="step-1">
    <div class="form-section">
      <p class="form-section-title">Personal Details</p>
      <div class="form-group"><label class="form-label">Name(s) <span>*</span></label><input type="text" id="t-name" class="form-input" placeholder="Your first name(s)"/></div>
      <div class="form-group"><label class="form-label">Surname <span>*</span></label><input type="text" id="t-surname" class="form-input" placeholder="Your surname"/></div>
      <div class="form-group"><label class="form-label">University <span>*</span></label><input type="text" id="t-university" class="form-input" placeholder="e.g. University of Pretoria"/></div>
      <div class="form-group"><label class="form-label">Faculty <span>*</span></label>
        <select id="t-faculty" class="form-select">
          <option value="">Select faculty</option>
          <option>EBIT</option><option>Education</option><option>EMS</option>
          <option>Humanities</option><option>NS</option><option>Veterinary</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">Degree / Programme <span>*</span></label><input type="text" id="t-degree" class="form-input" placeholder="e.g. BSc Mathematics"/></div>
      <div class="form-group"><label class="form-label">Year of Study <span>*</span></label>
        <select id="t-year" class="form-select">
          <option value="">Select year</option>
          <option>First Year</option><option>Second Year</option>
          <option>Third Year / Final Year</option><option>Honours / Masters</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">Contact Number <span>*</span></label><input type="tel" id="t-contact" class="form-input" placeholder="e.g. 0821234567"/></div>
      <div class="form-group"><label class="form-label">Email Address <span>*</span></label><input type="email" id="t-email" class="form-input" placeholder="your@email.com"/></div>
    </div>
    <div class="nav-btns"><button class="btn primary" onclick="tNext(1)">Next →</button></div>
  </div>

  <!-- STEP 2: Subject Availability -->
  <div class="form-step" id="step-2">
    <div class="form-section">
      <p class="form-section-title">Subject Availability</p>
      <p style="font-size:13px;color:var(--md);margin-bottom:1rem;">Which year and subject would you like to tutor? Select one option — modules will appear based on your selection.</p>
      <div class="radio-group">
        <label class="radio-item"><input type="radio" name="t-subj" value="First Year (Maths)" onchange="tSubjChange()"/><label>First Year — Mathematics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="First Year (Stats)" onchange="tSubjChange()"/><label>First Year — Statistics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="First Year (Maths & Stats)" onchange="tSubjChange()"/><label>First Year — Mathematics & Statistics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="Second Year (Maths)" onchange="tSubjChange()"/><label>Second Year — Mathematics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="Second Year (Stats)" onchange="tSubjChange()"/><label>Second Year — Statistics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="Second Year (Maths & Stats)" onchange="tSubjChange()"/><label>Second Year — Mathematics & Statistics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="Third Year / Final Year (Maths)" onchange="tSubjChange()"/><label>Third / Final Year — Mathematics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="Third Year / Final Year (Stats)" onchange="tSubjChange()"/><label>Third / Final Year — Statistics</label></label>
        <label class="radio-item"><input type="radio" name="t-subj" value="Third Year / Final Year (Maths & Stats)" onchange="tSubjChange()"/><label>Third / Final Year — Mathematics & Statistics</label></label>
      </div>
    </div>
    <div class="nav-btns"><button class="btn" onclick="tBack(2)">← Back</button><button class="btn primary" onclick="tNext(2)">Next →</button></div>
  </div>

  <!-- STEP 3: Module Selection -->
  <div class="form-step" id="step-3">
    <div class="form-section">
      <p class="form-section-title">Module Details</p>
      <p style="font-size:13px;color:var(--md);margin-bottom:1.5rem;">Select all modules you are confident to tutor.</p>
      <div id="t-maths-wrap" class="form-group" style="display:none;">
        <label class="form-label">Mathematics Modules</label>
        <div class="checkbox-grid" id="t-maths-modules"></div>
      </div>
      <div id="t-stats-wrap" class="form-group" style="margin-top:1.5rem;display:none;">
        <label class="form-label">Statistics Modules</label>
        <div class="checkbox-grid" id="t-stats-modules"></div>
      </div>
    </div>
    <div class="nav-btns"><button class="btn" onclick="tBack(3)">← Back</button><button class="btn primary" onclick="tNext(3)">Next →</button></div>
  </div>

  <!-- STEP 4: Proficiency & Experience -->
  <div class="form-step" id="step-4">
    <div class="form-section">
      <p class="form-section-title">Proficiency & Experience</p>
      <div class="form-group">
        <label class="form-label">Which tools / languages are you proficient in?</label>
        <div class="checkbox-grid">
          ${["Excel","LaTeX","MATLAB","Python","R","SAS","None"].map(t=>`<label class="checkbox-item"><input type="checkbox" name="t-tools" value="${t}"/><label>${t}</label></label>`).join("")}
        </div>
      </div>
      <div class="form-group" style="margin-top:1.5rem;">
        <label class="form-label">Do you have any tutoring experience? <span>*</span></label>
        <div class="radio-group">
          <label class="radio-item"><input type="radio" name="t-exp" value="Yes"/><label>Yes</label></label>
          <label class="radio-item"><input type="radio" name="t-exp" value="No"/><label>No</label></label>
        </div>
      </div>
      <div class="form-group" style="margin-top:1rem;">
        <label class="form-label">Are you currently tutoring privately or for another company? <span>*</span></label>
        <div class="radio-group">
          <label class="radio-item"><input type="radio" name="t-outside" value="Yes"/><label>Yes</label></label>
          <label class="radio-item"><input type="radio" name="t-outside" value="No"/><label>No</label></label>
        </div>
      </div>
    </div>
    <div class="nav-btns"><button class="btn" onclick="tBack(4)">← Back</button><button class="btn primary" onclick="tNext(4)">Next →</button></div>
  </div>

  <!-- STEP 5: Interview & Concept Explaining -->
  <div class="form-step" id="step-5">
    <div class="form-section">
      <p class="form-section-title">Interview Availability</p>
      <div class="form-group">
        <label class="form-label">When are you available for an interview? <span>*</span></label>
        <input type="text" id="t-interview" class="form-input" placeholder="e.g. January 7, 2026"/>
      </div>
      <div style="background:var(--tl);border-left:3px solid var(--t);padding:14px 18px;border-radius:0 8px 8px 0;margin-top:1.5rem;">
        <p style="font-size:13px;font-weight:600;color:var(--dk);margin-bottom:6px;">Concept Explaining</p>
        <p style="font-size:13px;color:var(--md);line-height:1.6;">You are required to prepare on first year topics for the interview, as you will be required to explain them. This includes methods and solutions.</p>
      </div>
    </div>
    <div class="nav-btns"><button class="btn" onclick="tBack(5)">← Back</button><button class="btn primary" onclick="tSubmit()">Submit Application →</button></div>
  </div>

  <!-- SUCCESS -->
  <div class="form-step" id="step-success">
    <div class="success-wrap">
      <div class="success-delta">Δ</div>
      <h2 class="success-title">Application Received!</h2>
      <p style="font-size:14px;color:var(--md);margin-bottom:16px;">Your tutor application has been submitted successfully.</p>
      <div class="success-id" id="success-id"></div>
      <p style="font-size:13px;color:var(--md);max-width:400px;margin:0 auto 24px;">Keep this ID safe. A confirmation email has been sent to you. We will be in touch within 24–48 hours.</p>
      <div style="margin-top:32px;display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
        <a href="${url}?page=portal" class="btn primary">Go to Client Portal</a>
        <a href="${url}" class="btn">Back to Home</a>
      </div>
    </div>
  </div>

</div>
${foot()}
<script>
const MATHS = ${JSON.stringify(MATHS)};
const STATS = ${JSON.stringify(STATS)};
const WEB_URL = '${url}';
let currentStep = 1;
const totalSteps = 5;

function updateDots(){
  const d=document.getElementById('dots'); d.innerHTML='';
  for(let i=1;i<=totalSteps;i++){
    const dot=document.createElement('div');
    dot.className='step-dot'+(i<currentStep?' done':i===currentStep?' current':'');
    d.appendChild(dot);
  }
}
updateDots();

function showStep(n){
  document.querySelectorAll('.form-step').forEach(s=>s.classList.remove('active'));
  document.getElementById('step-'+n).classList.add('active');
  window.scrollTo(0,0); currentStep=n; updateDots();
}

function tSubjChange(){
  const val = document.querySelector('input[name="t-subj"]:checked')?.value||'';
  const yearKey = val.includes('First Year')?'First Year':val.includes('Second Year')?'Second Year':'Third Year / Final Year';
  const showMaths = !val.includes('Stats')||val.includes('Maths & Stats');
  const showStats = val.includes('Stats');
  const mw=document.getElementById('t-maths-wrap'), sw=document.getElementById('t-stats-wrap');
  const mg=document.getElementById('t-maths-modules'), sg=document.getElementById('t-stats-modules');
  mg.innerHTML=''; sg.innerHTML='';
  if(showMaths){ mw.style.display='block'; (MATHS[yearKey]||[]).forEach(m=>{ mg.innerHTML+='<label class="checkbox-item"><input type="checkbox" name="t-maths" value="'+m+'"/><label>'+m+'</label></label>'; }); } else mw.style.display='none';
  if(showStats){ sw.style.display='block'; (STATS[yearKey]||[]).forEach(m=>{ sg.innerHTML+='<label class="checkbox-item"><input type="checkbox" name="t-stats" value="'+m+'"/><label>'+m+'</label></label>'; }); } else sw.style.display='none';
}

function showErr(msg){ const e=document.getElementById('err-box'); e.textContent=msg; e.style.display='block'; e.scrollIntoView({behavior:'smooth'}); }
function clearErr(){ document.getElementById('err-box').style.display='none'; }

function tNext(step){
  clearErr();
  if(step===1){
    if(!document.getElementById('t-name').value.trim()){showErr('Please enter your name.');return;}
    if(!document.getElementById('t-surname').value.trim()){showErr('Please enter your surname.');return;}
    if(!document.getElementById('t-university').value.trim()){showErr('Please enter your university.');return;}
    if(!document.getElementById('t-faculty').value){showErr('Please select your faculty.');return;}
    if(!document.getElementById('t-degree').value.trim()){showErr('Please enter your degree.');return;}
    if(!document.getElementById('t-year').value){showErr('Please select your year of study.');return;}
    if(!document.getElementById('t-contact').value.trim()){showErr('Please enter your contact number.');return;}
    if(!document.getElementById('t-email').value.trim()){showErr('Please enter your email address.');return;}
  }
  if(step===2){ if(!document.querySelector('input[name="t-subj"]:checked')){showErr('Please select a subject and year level.');return;} tSubjChange(); }
  if(step===3){ const m=document.querySelectorAll('input[name="t-maths"]:checked'); const s=document.querySelectorAll('input[name="t-stats"]:checked'); if(m.length===0&&s.length===0){showErr('Please select at least one module.');return;} }
  if(step===4){ if(!document.querySelector('input[name="t-exp"]:checked')){showErr('Please indicate your tutoring experience.');return;} if(!document.querySelector('input[name="t-outside"]:checked')){showErr('Please indicate outside tutoring status.');return;} }
  showStep(step+1);
}

function tBack(step){ clearErr(); showStep(step-1); }

function tSubmit(){
  clearErr();
  if(!document.getElementById('t-interview').value.trim()){showErr('Please provide your interview availability.');return;}
  const btn=event.target; btn.textContent='Submitting...'; btn.disabled=true;
  const data={
    action:'tutor-submit',
    name:document.getElementById('t-name').value.trim(),
    surname:document.getElementById('t-surname').value.trim(),
    university:document.getElementById('t-university').value.trim(),
    faculty:document.getElementById('t-faculty').value,
    degree:document.getElementById('t-degree').value.trim(),
    year:document.getElementById('t-year').value,
    contact:document.getElementById('t-contact').value.trim(),
    email:document.getElementById('t-email').value.trim(),
    mathsModules:[...document.querySelectorAll('input[name="t-maths"]:checked')].map(i=>i.value),
    statsModules:[...document.querySelectorAll('input[name="t-stats"]:checked')].map(i=>i.value),
    tools:[...document.querySelectorAll('input[name="t-tools"]:checked')].map(i=>i.value),
    experience:document.querySelector('input[name="t-exp"]:checked')?.value||'No',
    outsideTutoring:document.querySelector('input[name="t-outside"]:checked')?.value||'No',
    interviewDate:document.getElementById('t-interview').value.trim(),
  };
  fetch(WEB_URL,{method:'POST',body:JSON.stringify(data)})
    .then(r=>r.json())
    .then(res=>{
      if(res.success){ document.getElementById('success-id').textContent=res.appId; document.querySelectorAll('.form-step').forEach(s=>s.classList.remove('active')); document.getElementById('step-success').classList.add('active'); document.getElementById('dots').style.display='none'; window.scrollTo(0,0); }
      else{ showErr('Something went wrong. Please try again.'); btn.textContent='Submit Application →'; btn.disabled=false; }
    })
    .catch(()=>{ showErr('Connection error. Please try again.'); btn.textContent='Submit Application →'; btn.disabled=false; });
}
</script>
</body></html>`; }

// ── PORTAL LOGIN ──────────────────────────────────────────
function portalLogin(errorMsg) {
  const url = ScriptApp.getService().getUrl();
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Client Portal — Delta Tutoring</title>${css()}</head><body>
${nav("portal")}
<div style="max-width:440px;margin:80px auto;padding:0 24px;">
  <div style="background:#fff;border-radius:14px;border:0.5px solid var(--bd);overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.07);">
    <div style="background:var(--t);padding:28px 32px;text-align:center;">
      <div style="font-family:'DM Serif Display',serif;font-size:44px;color:#fff;line-height:1;margin-bottom:8px;">Δ</div>
      <div style="font-size:16px;font-weight:600;color:#fff;margin-bottom:4px;">Client Portal</div>
      <div style="font-size:12px;color:var(--tl);">Enter your Application ID to access your package</div>
    </div>
    <div style="padding:28px 32px;">
      ${errorMsg?`<div style="background:#FEE2E2;border:1px solid #FECACA;border-radius:8px;padding:10px 14px;font-size:13px;color:#DC2626;margin-bottom:16px;">${errorMsg}</div>`:''}
      <label style="font-size:12px;font-weight:500;color:var(--md);margin-bottom:6px;display:block;text-transform:uppercase;letter-spacing:.06em;">Application ID</label>
      <input type="text" id="pid" class="form-input" placeholder="e.g. DT-STU-2026-0001" onkeydown="if(event.key==='Enter')pLogin()" style="margin-bottom:12px;"/>
      <a href="#" onclick="pLogin();return false;" class="btn primary full">Access My Package →</a>
      <p style="font-size:11px;color:#aaa;text-align:center;margin-top:16px;">Your ID was emailed when your application was accepted.<br/>Contact hello@deltatutoring.co.za for help.</p>
    </div>
  </div>
  <p style="text-align:center;margin-top:16px;"><a href="${url}" style="font-size:13px;color:var(--md);text-decoration:none;">← Back to home</a></p>
</div>
${foot()}
<script>
const WEB_URL='${url}';
function pLogin(){ const id=document.getElementById('pid').value.trim().toUpperCase(); if(!id){alert('Please enter your Application ID.');return;} window.location.href=WEB_URL+'?page=portal&id='+encodeURIComponent(id); }
</script>
</body></html>`; }

// ── PORTAL DASHBOARD ──────────────────────────────────────
function portalDashboard(c) {
  const url = ScriptApp.getService().getUrl();
  const isStudent = c.type === "student";
  const docs = isStudent ? getStudentDocs(c) : getTutorDocs(c);

  const chips = isStudent
    ? `<span class="chip">📦 ${c.package||''}</span>${(c.modules||[]).map(m=>`<span class="chip">📚 ${m}</span>`).join('')}<span class="chip">👤 ${c.tutor||'TBA'}</span>`
    : `${(c.modules||[]).map(m=>`<span class="chip">📚 ${m}</span>`).join('')}<span class="chip">👥 ${c.studentsAssigned||'None assigned'}</span><span class="chip">⏱ Max ${c.maxHours||20}hrs/month</span>`;

  const docCards = docs.map((d,i)=>`
    <div class="doc-card" onclick="openDoc(${i})">
      <div class="doc-icon">${d.icon}</div>
      <div class="doc-title">${d.title}</div>
      <div class="doc-desc">${d.desc}</div>
    </div>`).join('');

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>My Portal — Delta Tutoring</title>${css()}</head><body>
${nav("portal")}
<div class="portal-wrap">
  <div class="welcome-card">
    <div><div class="welcome-name">Welcome, ${c.name} 👋</div><div class="welcome-sub">${isStudent?'STUDENT PORTAL':'TUTOR PORTAL'} · ${c.month||''}</div></div>
    <div class="welcome-id">${c.id}</div>
  </div>
  <div class="chips">${chips}</div>
  <p style="font-size:11px;font-weight:600;color:var(--md);text-transform:uppercase;letter-spacing:.1em;margin-bottom:12px;">Your Package Documents</p>
  <div class="docs-grid">${docCards}</div>
  <p style="text-align:center;font-size:11px;color:#aaa;">Δ Delta Tutoring · hello@deltatutoring.co.za · Pretoria</p>
</div>

<div class="modal-overlay" id="modal" onclick="if(event.target.id==='modal')closeDoc()">
  <div class="modal-box">
    <div class="modal-head"><span class="modal-title" id="modal-title"></span><button class="modal-close" onclick="closeDoc()">Close ✕</button></div>
    <div class="modal-body" id="modal-body"></div>
  </div>
</div>

${foot()}
<script>
const DOCS=${JSON.stringify(docs)};
function openDoc(i){ const d=DOCS[i]; document.getElementById('modal-title').textContent=d.icon+' '+d.title; document.getElementById('modal-body').textContent=d.content; document.getElementById('modal').style.display='flex'; }
function closeDoc(){ document.getElementById('modal').style.display='none'; }
</script>
</body></html>`; }

// ── DOCUMENT CONTENT ──────────────────────────────────────
function getStudentDocs(c) { return [
  {icon:"📄",title:"Student Service Agreement",desc:"Your contract with Delta Tutoring",
   content:`DELTA TUTORING — STUDENT SERVICE AGREEMENT\n\nStudent: ${c.name}\nPackage: ${c.package||''}\nModules: ${(c.modules||[]).join(', ')}\nTutor Assigned: ${c.tutor||'TBA'}\n\n1. COMMENCEMENT\nBinding upon signature, consultation and first payment.\n\n2. PACKAGE\n${c.package||''} — valid current month only. Unused hours do not carry over.\n\n3. PAYMENT\nFull or bi-weekly. Instalment 1 before session 1. Instalment 2 by 15th.\n\n4. COMMITMENT\n• Attend punctually and prepared\n• Complete practice between sessions\n• Engage actively — tutor guides, not carries\n• No recording without written consent\n\n5. CANCELLATION\n24 hours notice required. Late cancellations forfeited.\n\n6. PROGRESS\nMonthly Momentum Report with SSS score and goal probability.\n\n7. GOVERNING LAW\nRepublic of South Africa. Mediation before litigation.`},
  {icon:"🧰",title:"Student Toolkit",desc:"Tools to support your learning",
   content:`DELTA TUTORING — STUDENT TOOLKIT\n\n📱 Study Circle — log daily study hours and streaks\n🧮 Wolfram Alpha — wolframalpha.com\n📈 Desmos — desmos.com\n🃏 Anki — apps.ankiweb.net\n💻 RStudio — rstudio.com\n📊 Overleaf — overleaf.com\n\nSESSION CHECKLIST\n☐ Review last session notes\n☐ Attempt practice questions beforehand\n☐ Write down questions to ask\n☐ Log hours on Study Circle\n☐ Bring textbook and past papers\n\nhello@deltatutoring.co.za`},
  {icon:"📊",title:"Momentum Report",desc:`Progress report — ${c.month||'current month'}`,
   content:`DELTA TUTORING — MOMENTUM REPORT\n\nStudent: ${c.name}\nModules: ${(c.modules||[]).join(', ')}\nMonth: ${c.month||''}\n\nYour Momentum Report will be sent at end of each month.\n\nIncludes:\n• SSS score /100\n• Month-on-month delta\n• Probability of reaching goal\n• Practice, Attendance, Confidence, Consistency, Topic Mastery breakdown\n• Top 3 action items\n• Tutor notes\n\nΔ Keep moving. Every session, every question, every study hour counts.`},
  {icon:"📘",title:"Student Handbook",desc:"How Delta Tutoring works",
   content:`DELTA TUTORING — STUDENT HANDBOOK\n\nPHILOSOPHY\nWe measure the delta — the change from where you started.\n\nRESPONSIBILITIES\n• Arrive on time and prepared\n• Complete practice between sessions\n• Log study hours on Study Circle\n• Communicate upcoming tests early\n• Give honest confidence ratings\n\nCANCELLATION\n24 hours notice required. Emergencies reviewed case by case.\n\nCOMMUNICATION\n• All queries via WhatsApp or email\n• Respond within 24 hours\n• Raise concerns with Delta Tutoring — not your tutor directly\n\nhello@deltatutoring.co.za`},
];}

function getTutorDocs(c) { return [
  {icon:"📄",title:"Tutor Service Agreement",desc:"Your contract with Delta Tutoring",
   content:`DELTA TUTORING — TUTOR SERVICE AGREEMENT\n\nTutor: ${c.name}\nModules: ${(c.modules||[]).join(', ')}\nStudents Assigned: ${c.studentsAssigned||'TBA'}\nMax Hours: ${c.maxHours||20}/month\n\n1. RELATIONSHIP\nIndependent service provider — not an employee.\n\n2. HOUR LIMITS\n• Max 20hrs/month\n• 5hr extension by written application only\n• No extension if outside tutoring declared\n\n3. REMUNERATION\nFirst Year: R120/hr | Second Year: R150/hr | Third Year+: R180/hr\nPayment within 5 business days of month end.\n\n4. OUTSIDE TUTORING\nMust be declared. Non-disclosure = immediate termination.\n\n5. CONDUCT\n• Arrive punctually and prepared\n• 24hrs notice for cancellations\n• No private arrangements with Delta students\n• Maintain full student confidentiality\n\n6. GOVERNING LAW\nRepublic of South Africa. Mediation before litigation.`},
  {icon:"🧰",title:"Tutor Toolkit",desc:"Resources for great sessions",
   content:`DELTA TUTORING — TUTOR TOOLKIT\n\nBEFORE SESSION\n☐ Review student's last Momentum Report\n☐ Note lowest Topic Mastery scores\n☐ Prepare targeted problems for weak areas\n\nDURING SESSION\n☐ 5-min check-in\n☐ Review last session homework\n☐ Introduce concept clearly\n☐ Student attempts independently first\n☐ End with 3-5 question mini-quiz\n☐ Assign next session practice\n\nAFTER SESSION\n☐ Log session details\n☐ Rate student engagement\n☐ Submit log by month end\n\nhello@deltatutoring.co.za`},
  {icon:"📘",title:"Tutor Handbook",desc:"Standards at Delta Tutoring",
   content:`DELTA TUTORING — TUTOR HANDBOOK\n\nYou were selected because you can explain — not just solve.\n\nPHILOSOPHY\nWe protect tutors. Hours capped. Students vetted. Your academics first — always.\n\nYOUR ROLE\nBuild student independence. Guide, don't carry.\n\nMONTHLY SUBMISSION\n• Session log\n• Engagement ratings\n• Topic quiz scores\n• Written student feedback\n\nCONDUCT\n• No private sessions with Delta students\n• Never share student information\n• Represent Delta Tutoring professionally\n\nPAYMENT\nLog by last working day → payment within 5 business days.\n\nhello@deltatutoring.co.za`},
  {icon:"📊",title:"Session Log Template",desc:"Submit at month end for payment",
   content:`DELTA TUTORING — SESSION LOG\n\nTutor: ${c.name} | Month: ${c.month||''}\n\nFormat:\nDate | Student ID | Module | Duration | Topics | Prepared(1-3) | Participated(1-3) | Questions(1-4) | Late Cancel(Y/N) | Notes\n\n─────────────────────────────────\n[Complete in your Excel workbook]\n─────────────────────────────────\n\nTOTAL HOURS: ___\nTOTAL SESSIONS: ___\n\nMONTHLY SUMMARY\nStudent progress:\nConcerns:\nTopics covered:\n\nSignature: ___________ Date: ___________\n\nΔ Delta Tutoring — hello@deltatutoring.co.za`},
];}
