// ═══════════════════════════════════════════════════════════════════
//  SevakCall — Code.gs  (Production-Ready Upgrade)
//  Sheets: Contacts | Volunteers | Activity | Events | Areas
// ═══════════════════════════════════════════════════════════════════

const CONTACTS_SHEET   = 'Contacts';
const VOLUNTEERS_SHEET = 'Volunteers';
const ACTIVITY_SHEET   = 'Activity';
const EVENTS_SHEET     = 'Events';
const AREAS_SHEET      = 'Areas';

// ─────────────────────────────────────────────────────────────────
//  RBAC
// ─────────────────────────────────────────────────────────────────
function getVolByEmail(email) {
  const rows = getSheet(VOLUNTEERS_SHEET).getDataRange().getValues();
  // Name | Email | Password | Role | Status
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]||'').trim().toLowerCase() === String(email||'').trim().toLowerCase()) {
      return {
        name:   String(rows[i][0]||'').trim(),
        email:  String(rows[i][1]||'').trim(),
        role:   String(rows[i][3]||'volunteer').trim().toLowerCase(),
        status: String(rows[i][4]||'active').trim().toLowerCase(),
        row:    i + 1
      };
    }
  }
  return null;
}

function isAdmin(r)     { return r === 'admin'; }
function isMod(r)       { return r === 'admin' || r === 'moderator'; }

// ─────────────────────────────────────────────────────────────────
//  ROUTER — GET
// ─────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const a = e.parameter.action || '';
    const em = e.parameter.email || '';
    let r;
    switch (a) {
      case 'login':          r = login(em, e.parameter.password);               break;
      case 'getMyContacts':  r = getMyContacts(em);                             break;
      case 'getAdminStats':  r = getAdminStats(em);                             break;
      case 'getEvents':      r = getEvents();                                   break;
      case 'getAllContacts':  r = getAllContacts();                               break;
      case 'getAreas':       r = getAreas();                                    break;
      case 'getActivityLog': r = getActivityLog(em, e.parameter.page,
                                 e.parameter.startDate, e.parameter.endDate);  break;
      case 'getUnassigned':  r = getUnassigned(em, e.parameter.q);             break;
      default:               r = { error: 'Unknown action: ' + a };
    }
    return jsonOut(r);
  } catch(err) { logAct('SYSTEM','ERROR',err.message); return jsonOut({ error: err.message }); }
}

// ─────────────────────────────────────────────────────────────────
//  ROUTER — POST
// ─────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const d    = JSON.parse(e.postData.contents);
    const caller = d.callerEmail ? getVolByEmail(d.callerEmail) : null;
    const role   = caller ? caller.role : 'volunteer';
    let r;
    switch (d.action) {
      // Volunteer
      case 'updateContact':     r = updateContact(d.ID, d.Status, d.Reference, d.callerEmail);       break;
      case 'markAttendance':    r = markAttendance(d.contactId, d.columnName, d.callerEmail);         break;
      case 'addContactAndMark': r = addContactAndMark(d.contactData, d.columnName, d.callerEmail);   break;
      // Moderator+
      case 'addContact':
        if (!isMod(role)) { r={success:false,error:'Access denied.'}; break; }
        r = addContact(d.contactData, d.callerEmail); break;
      case 'editContact':
        if (!isMod(role)) { r={success:false,error:'Access denied.'}; break; }
        r = editContact(d.contactData, d.callerEmail); break;
      case 'reassignContacts':
        if (!isMod(role)) { r={success:false,error:'Access denied.'}; break; }
        r = reassignContacts(d.from, d.to, d.contactIds, d.callerEmail); break;
      case 'assignPool':
        if (!isMod(role)) { r={success:false,error:'Access denied.'}; break; }
        r = assignPool(d.contactIds, d.toVolunteer, d.callerEmail); break;
      // Admin only
      case 'manageVolunteer':
        if (!isAdmin(role)) { r={success:false,error:'Access denied.'}; break; }
        r = manageVolunteer(d); break;
      case 'createEvent':
        if (!isAdmin(role)) { r={success:false,error:'Access denied.'}; break; }
        r = createEvent(d.title, d.date, d.time, d.duration, d.speaker, d.callerEmail); break;
      case 'manageArea':
        if (!isAdmin(role)) { r={success:false,error:'Access denied.'}; break; }
        r = manageArea(d); break;
      case 'syncCalendar':
        if (!isAdmin(role)) { r={success:false,error:'Access denied.'}; break; }
        r = syncToCalendar(d.contactId, d.callerEmail); break;
      case 'sendDailyGreetings': r = sendDailyGreetings(); break;
      default: r = { error: 'Unknown action: ' + d.action };
    }
    return jsonOut(r);
  } catch(err) { logAct('SYSTEM','ERROR',err.message); return jsonOut({ error: err.message }); }
}

// ─────────────────────────────────────────────────────────────────
//  LOGIN
// ─────────────────────────────────────────────────────────────────
function login(email, password) {
  try {
    const rows = getSheet(VOLUNTEERS_SHEET).getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const [name, rEmail, rPass, role, status] = rows[i];
      if (!name) continue;
      if (String(status||'').toLowerCase() === 'deactivated') continue;
      if (String(rEmail).trim().toLowerCase() === String(email).trim().toLowerCase() &&
          String(rPass).trim() === String(password).trim()) {
        const r = String(role||'volunteer').trim().toLowerCase();
        logAct(String(name).trim(), 'LOGIN', 'Signed in');
        return { success: true, user: { name: String(name).trim(), email: String(rEmail).trim(), role: r } };
      }
    }
    return { success: false, error: 'Invalid credentials.' };
  } catch(err) { return { success: false, error: err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  UNIQUE ID GENERATION
//  Format: [Serial]-[AreaCode]-[First3Name]-[Last4Phone]
//  e.g.  101-AM-MAN-5884
// ─────────────────────────────────────────────────────────────────
function generateContactID(areaCode, name, phone) {
  try {
    const serial = Math.max(1, getSheet(CONTACTS_SHEET).getLastRow() - 1 + 1);
    const code   = String(areaCode||'XX').toUpperCase().replace(/[^A-Z0-9]/g,'').substring(0,4);
    const nm3    = String(name||'UNK').toUpperCase().replace(/[^A-Z]/g,'').substring(0,3).padEnd(3,'X');
    const ph4    = String(phone||'0000').replace(/\D/g,'').slice(-4).padStart(4,'0');
    return `${serial}-${code}-${nm3}-${ph4}`;
  } catch(_) { return 'C' + Date.now(); }
}

// ─────────────────────────────────────────────────────────────────
//  DUPLICATE PHONE CHECK
// ─────────────────────────────────────────────────────────────────
function checkDupPhone(phone, excludeId) {
  const clean = String(phone||'').replace(/\D/g,'');
  if (clean.length !== 10) return { isDuplicate: false };
  const rows    = getSheet(CONTACTS_SHEET).getDataRange().getValues();
  const headers = rows[0];
  const pi = headers.indexOf('Phone'), ii = headers.indexOf('ID'), ni = headers.indexOf('Name');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][pi]||'').replace(/\D/g,'') === clean &&
        String(rows[i][ii]||'').trim() !== String(excludeId||'')) {
      return { isDuplicate: true, existingName: String(rows[i][ni]||''), existingId: String(rows[i][ii]||'') };
    }
  }
  return { isDuplicate: false };
}

// ─────────────────────────────────────────────────────────────────
//  AREAS
// ─────────────────────────────────────────────────────────────────
function getAreas() {
  try {
    const rows = getSheet(AREAS_SHEET).getDataRange().getValues();
    const areas = rows.slice(1).filter(r=>String(r[0]||'').trim())
      .map(r=>({ name: String(r[0]).trim(), code: String(r[1]||'').trim() }));
    return { areas };
  } catch(err) { return { areas:[], error:err.message }; }
}

function manageArea(d) {
  try {
    const sheet = getSheet(AREAS_SHEET);
    const rows  = sheet.getDataRange().getValues();
    if (d.mode === 'add') {
      if (!d.name) return { success:false, error:'Name required.' };
      for (let i=1;i<rows.length;i++) {
        if (String(rows[i][0]||'').trim().toLowerCase()===d.name.trim().toLowerCase())
          return { success:false, error:'Area already exists.' };
      }
      sheet.appendRow([d.name.trim(), (d.code||'').trim().toUpperCase()]);
      logAct(d.callerEmail||'admin','ADD_AREA',d.name);
      return { success:true };
    }
    if (d.mode === 'delete') {
      for (let i=1;i<rows.length;i++) {
        if (String(rows[i][0]||'').trim().toLowerCase()===d.name.trim().toLowerCase()) {
          sheet.deleteRow(i+1);
          logAct(d.callerEmail||'admin','DELETE_AREA',d.name);
          return { success:true };
        }
      }
      return { success:false, error:'Not found.' };
    }
    return { success:false, error:'Unknown mode.' };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  GET MY CONTACTS
// ─────────────────────────────────────────────────────────────────
function getMyContacts(email) {
  try {
    const vol = getVolByEmail(email);
    if (!vol) return { contacts:[], template:null, error:'Volunteer not found.' };
    const cSheet  = getSheet(CONTACTS_SHEET);
    const cRows   = cSheet.getDataRange().getValues();
    const headers = cRows[0];
    const ai = headers.indexOf('Assigned_To');
    const contacts = cRows.slice(1)
      .filter(r=>String(r[ai]||'').trim().toLowerCase()===vol.name.toLowerCase())
      .map(r=>{ const o={}; headers.forEach((h,i)=>o[h]=r[i]??''); return o; });
    let template = null;
    try {
      const ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Message');
      if (ms) template = { text: String(ms.getRange('A2').getValue()||''), image: String(ms.getRange('B2').getValue()||'') };
    } catch(_) {}
    return { contacts, template, role: vol.role };
  } catch(err) { return { contacts:[], template:null, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  GET ALL CONTACTS
// ─────────────────────────────────────────────────────────────────
function getAllContacts() {
  try {
    const rows = getSheet(CONTACTS_SHEET).getDataRange().getValues();
    const headers = rows[0];
    return { contacts: rows.slice(1).map(r=>{ const o={}; headers.forEach((h,i)=>o[h]=r[i]??''); return o; }), headers };
  } catch(err) { return { contacts:[], headers:[], error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  GET UNASSIGNED (for contact pool tool)
// ─────────────────────────────────────────────────────────────────
function getUnassigned(email, q) {
  try {
    const vol = getVolByEmail(email);
    if (!vol || !isMod(vol.role)) return { success:false, error:'Access denied.' };
    const rows = getSheet(CONTACTS_SHEET).getDataRange().getValues();
    const h    = rows[0];
    const ai=h.indexOf('Assigned_To'), ni=h.indexOf('Name'), pi=h.indexOf('Phone'),
          xi=h.indexOf('Area'), ii=h.indexOf('ID');
    const qL = (q||'').toLowerCase().trim();
    const contacts = rows.slice(1)
      .filter(r=>!String(r[ai]||'').trim())
      .filter(r=>!qL || String(r[ni]||'').toLowerCase().includes(qL) ||
                        String(r[pi]||'').includes(qL) ||
                        String(r[xi]||'').toLowerCase().includes(qL))
      .slice(0,50)
      .map(r=>({ ID:String(r[ii]||''), Name:String(r[ni]||''), Phone:String(r[pi]||''), Area:String(r[xi]||'') }));
    return { contacts };
  } catch(err) { return { contacts:[], error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  ADD CONTACT  (Moderator+)
// ─────────────────────────────────────────────────────────────────
function addContact(cd, callerEmail) {
  try {
    if (!cd) return { success:false, error:'No data provided.' };
    const phone = String(cd.Phone||'').replace(/\D/g,'');
    if (phone.length !== 10) return { success:false, error:'Phone must be exactly 10 digits.' };
    const dup = checkDupPhone(phone, null);
    if (dup.isDuplicate) return { success:false, error:`Phone already exists for: ${dup.existingName} (ID: ${dup.existingId})` };

    let areaCode = 'XX';
    try {
      const ar = getSheet(AREAS_SHEET).getDataRange().getValues();
      const match = ar.slice(1).find(r=>String(r[0]||'').toLowerCase()===String(cd.Area||'').toLowerCase());
      if (match) areaCode = String(match[1]||'XX');
    } catch(_) {}

    const newId = generateContactID(areaCode, cd.Name, phone);
    const sheet = getSheet(CONTACTS_SHEET);
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

    const newRow = headers.map(h => {
      switch(h) {
        case 'ID':               return newId;
        case 'Name':             return String(cd.Name||'').trim();
        case 'Phone':            return phone;
        case 'DOB':              return cd.DOB||'';
        case 'DOA':              return cd.DOA||'';
        case 'Profession':       return String(cd.Profession||'').trim();
        case 'Skill':            return String(cd.Skill||'').trim();
        case 'Mandal':           return String(cd.Mandal||'').trim();
        case 'Note':             return String(cd.Note||'').trim();
        case 'Complete_Address': return String(cd.Complete_Address||'').trim();
        case 'Area':             return String(cd.Area||'').trim();
        case 'Assigned_To':      return String(cd.Assigned_To||'').trim();
        case 'Status':           return '';
        case 'Reference':        return String(cd.Reference||'').trim();
        default:                 return '';
      }
    });

    sheet.appendRow(newRow);
    logAct(callerEmail||'system', 'ADD_CONTACT', `${cd.Name} (${newId})`);
    if (cd.DOB || cd.DOA) { try { syncToCalendar(newId, callerEmail); } catch(_) {} }
    return { success:true, newId, name:cd.Name };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  EDIT CONTACT  (Moderator+)
// ─────────────────────────────────────────────────────────────────
function editContact(cd, callerEmail) {
  try {
    if (!cd||!cd.ID) return { success:false, error:'Contact ID required.' };
    if (cd.Phone) {
      const ph = String(cd.Phone).replace(/\D/g,'');
      if (ph.length!==10) return { success:false, error:'Phone must be exactly 10 digits.' };
      const dup = checkDupPhone(ph, cd.ID);
      if (dup.isDuplicate) return { success:false, error:`Phone used by: ${dup.existingName}` };
    }
    const sheet = getSheet(CONTACTS_SHEET);
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const ii    = h.indexOf('ID');
    const fields = ['Name','Phone','DOB','DOA','Profession','Skill','Mandal','Note','Complete_Address','Area','Assigned_To','Status','Reference'];
    for (let i=1;i<rows.length;i++) {
      if (String(rows[i][ii]||'').trim()===String(cd.ID).trim()) {
        fields.forEach(f=>{
          if (cd[f]!==undefined) {
            const ci=h.indexOf(f);
            if (ci>=0) sheet.getRange(i+1,ci+1).setValue(f==='Phone'?String(cd[f]).replace(/\D/g,''):cd[f]||'');
          }
        });
        logAct(callerEmail||'system','EDIT_CONTACT',String(cd.Name||cd.ID));
        return { success:true };
      }
    }
    return { success:false, error:'Contact not found.' };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  UPDATE CONTACT STATUS  (Volunteer)
// ─────────────────────────────────────────────────────────────────
function updateContact(id, status, reference, callerEmail) {
  try {
    const sheet = getSheet(CONTACTS_SHEET);
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const ii=h.indexOf('ID')+1, si=h.indexOf('Status')+1, ri=h.indexOf('Reference')+1;
    for (let i=1;i<rows.length;i++) {
      if (String(rows[i][ii-1]).trim()===String(id).trim()) {
        sheet.getRange(i+1,si).setValue(status||'');
        sheet.getRange(i+1,ri).setValue(reference||'');
        logAct(callerEmail||String(rows[i][h.indexOf('Assigned_To')]||''), 'UPDATE_STATUS',
          `${String(rows[i][h.indexOf('Name')]||'')}: ${status}`);
        return { success:true };
      }
    }
    return { success:false, error:'Contact not found.' };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  ASSIGN CONTACT POOL  (Moderator+)
// ─────────────────────────────────────────────────────────────────
function assignPool(contactIds, toVolunteer, callerEmail) {
  try {
    if (!contactIds||!contactIds.length) return { success:false, error:'No contacts selected.' };
    if (!toVolunteer) return { success:false, error:'Target volunteer required.' };
    const sheet = getSheet(CONTACTS_SHEET);
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const ii=h.indexOf('ID'), ai=h.indexOf('Assigned_To')+1;
    const ids = new Set(contactIds.map(String));
    let count=0;
    for (let i=1;i<rows.length;i++) {
      if (ids.has(String(rows[i][ii]||'').trim())) { sheet.getRange(i+1,ai).setValue(toVolunteer); count++; }
    }
    logAct(callerEmail||'system','ASSIGN_POOL',`${count} contacts → ${toVolunteer}`);
    return { success:true, assigned:count };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  REASSIGN CONTACTS  (Moderator+)
// ─────────────────────────────────────────────────────────────────
function reassignContacts(fromName, toName, contactIds, callerEmail) {
  try {
    const sheet = getSheet(CONTACTS_SHEET);
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const ai=h.indexOf('Assigned_To')+1, si=h.indexOf('Status'), ii=h.indexOf('ID');
    const ids = contactIds&&contactIds.length ? new Set(contactIds.map(String)) : null;
    let count=0;
    for (let i=1;i<rows.length;i++) {
      const asgn=String(rows[i][ai-1]||'').trim().toLowerCase();
      const stat=String(rows[i][si]||'').trim();
      const rid =String(rows[i][ii]||'').trim();
      if (asgn===fromName.toLowerCase() && (ids ? ids.has(rid) : !stat)) {
        sheet.getRange(i+1,ai).setValue(toName);
        count++;
      }
    }
    logAct(callerEmail||'system','REASSIGN',`${count}: ${fromName}→${toName}`);
    return { success:true, reassigned:count };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  MANAGE VOLUNTEER  (Admin)
// ─────────────────────────────────────────────────────────────────
function manageVolunteer(d) {
  try {
    const sheet = getSheet(VOLUNTEERS_SHEET);
    const rows  = sheet.getDataRange().getValues();
    // Name | Email | Password | Role | Status
    if (d.mode==='add') {
      if (!d.name||!d.email||!d.password) return { success:false, error:'Name, email, password required.' };
      for (let i=1;i<rows.length;i++) {
        if (String(rows[i][1]||'').toLowerCase()===d.email.toLowerCase()) return { success:false, error:'Email exists.' };
      }
      sheet.appendRow([d.name, d.email, d.password, d.role||'volunteer', 'active']);
      if (d.area) {
        const cs=getSheet(CONTACTS_SHEET), cr=cs.getDataRange().getValues(), hh=cr[0];
        const ai=hh.indexOf('Assigned_To')+1, xi=hh.indexOf('Area')+1, si=hh.indexOf('Status');
        for (let i=1;i<cr.length;i++) {
          if (String(cr[i][xi-1]||'').toLowerCase()===d.area.toLowerCase()&&!String(cr[i][ai-1]||'').trim()&&!String(cr[i][si]||'').trim())
            cs.getRange(i+1,ai).setValue(d.name);
        }
      }
      logAct(d.callerEmail||'admin','ADD_VOL',d.name);
      return { success:true };
    }
    if (d.mode==='edit') {
      for (let i=1;i<rows.length;i++) {
        if (String(rows[i][0]).toLowerCase()===d.originalName.toLowerCase()) {
          if (d.name)     sheet.getRange(i+1,1).setValue(d.name);
          if (d.email)    sheet.getRange(i+1,2).setValue(d.email);
          if (d.password) sheet.getRange(i+1,3).setValue(d.password);
          if (d.role)     sheet.getRange(i+1,4).setValue(d.role);
          if (d.name && d.name!==d.originalName) {
            const cs=getSheet(CONTACTS_SHEET), cr=cs.getDataRange().getValues();
            const ai=cr[0].indexOf('Assigned_To')+1;
            for (let j=1;j<cr.length;j++) {
              if (String(cr[j][ai-1]||'').toLowerCase()===d.originalName.toLowerCase()) cs.getRange(j+1,ai).setValue(d.name);
            }
          }
          logAct(d.callerEmail||'admin','EDIT_VOL',d.name||d.originalName);
          return { success:true };
        }
      }
      return { success:false, error:'Volunteer not found.' };
    }
    if (d.mode==='deactivate') {
      for (let i=1;i<rows.length;i++) {
        if (String(rows[i][0]).toLowerCase()===d.originalName.toLowerCase()) {
          sheet.getRange(i+1,5).setValue('deactivated');
          logAct(d.callerEmail||'admin','DEACTIVATE_VOL',d.originalName);
          return { success:true };
        }
      }
      return { success:false, error:'Not found.' };
    }
    return { success:false, error:'Unknown mode.' };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  ADMIN STATS
// ─────────────────────────────────────────────────────────────────
function getAdminStats(callerEmail) {
  try {
    const caller = callerEmail ? getVolByEmail(callerEmail) : null;
    if (caller && !isMod(caller.role)) return { error:'Access denied.' };

    const cSheet=getSheet(CONTACTS_SHEET), cRows=cSheet.getDataRange().getValues();
    const h=cRows[0], data=cRows.slice(1);
    const si=h.indexOf('Status'), ai=h.indexOf('Assigned_To'), xi=h.indexOf('Area');

    const totalContacts=data.length;
    const totalCalled=data.filter(r=>String(r[si]||'').trim()).length;
    const statusBreakdown={};
    data.forEach(r=>{ const s=String(r[si]||'').trim(); if(s) statusBreakdown[s]=(statusBreakdown[s]||0)+1; });
    const allContacts=data.map(r=>{ const o={}; h.forEach((hh,i)=>o[hh]=r[i]??''); return o; });

    const vSheet=getSheet(VOLUNTEERS_SHEET), vRows=vSheet.getDataRange().getValues();
    const volunteers=vRows.slice(1)
      .filter(r=>String(r[0]||'').trim()&&String(r[4]||'').toLowerCase()!=='deactivated')
      .map(r=>{
        const name=String(r[0]).trim(), email=String(r[1]).trim(), role=String(r[3]||'volunteer').toLowerCase();
        const mine=data.filter(c=>String(c[ai]||'').toLowerCase()===name.toLowerCase());
        const total=mine.length, called=mine.filter(c=>String(c[si]||'').trim()).length;
        const interested=mine.filter(c=>['Interested','Already Volunteer'].includes(String(c[si]||'').trim())).length;
        const sb={}; mine.forEach(c=>{const s=String(c[si]||'').trim();if(s)sb[s]=(sb[s]||0)+1;});
        const areas=[...new Set(mine.map(c=>String(c[xi]||'').trim()).filter(Boolean))];
        return {name,email,role,total,called,interested,statusBreakdown:sb,areas};
      });

    const areaStats={};
    data.forEach(r=>{
      const area=String(r[xi]||'').trim()||'Unknown', status=String(r[si]||'').trim();
      if(!areaStats[area]) areaStats[area]={total:0,called:0,interested:0};
      areaStats[area].total++;
      if(status) areaStats[area].called++;
      if(['Interested','Already Volunteer'].includes(status)) areaStats[area].interested++;
    });

    let activityLog=[];
    try {
      const ar=getSheet(ACTIVITY_SHEET).getDataRange().getValues();
      activityLog=ar.slice(1).reverse().slice(0,40).map(r=>({
        time:    r[0]?Utilities.formatDate(new Date(r[0]),Session.getScriptTimeZone(),'dd MMM HH:mm'):'',
        user:    String(r[1]||''), action:String(r[2]||''), details:String(r[3]||''),
      }));
    } catch(_) {}

    return {totalContacts,totalCalled,statusBreakdown,allContacts,volunteers,areaStats,activityLog};
  } catch(err) { return { error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  PAGINATED ACTIVITY LOG
// ─────────────────────────────────────────────────────────────────
function getActivityLog(callerEmail, page, startDate, endDate) {
  try {
    const caller=callerEmail?getVolByEmail(callerEmail):null;
    if (!caller||!isMod(caller.role)) return { error:'Access denied.' };
    const PAGE=100, pg=Math.max(1,parseInt(page)||1);
    let rows=getSheet(ACTIVITY_SHEET).getDataRange().getValues().slice(1).reverse();
    if (startDate) { const sd=new Date(startDate); sd.setHours(0,0,0,0); rows=rows.filter(r=>r[0]&&new Date(r[0])>=sd); }
    if (endDate)   { const ed=new Date(endDate);   ed.setHours(23,59,59,999); rows=rows.filter(r=>r[0]&&new Date(r[0])<=ed); }
    const total=rows.length, totalPages=Math.ceil(total/PAGE)||1;
    const log=rows.slice((pg-1)*PAGE,pg*PAGE).map(r=>({
      time:    r[0]?Utilities.formatDate(new Date(r[0]),Session.getScriptTimeZone(),'dd MMM yyyy HH:mm'):'',
      user:    String(r[1]||''), action:String(r[2]||''), details:String(r[3]||''),
    }));
    return { log, page:pg, totalPages, totalRows:total };
  } catch(err) { return { log:[],page:1,totalPages:1,totalRows:0,error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  GOOGLE CALENDAR SYNC  (Admin)
// ─────────────────────────────────────────────────────────────────
function syncToCalendar(contactId, callerEmail) {
  try {
    const rows=getSheet(CONTACTS_SHEET).getDataRange().getValues(), h=rows[0];
    const ii=h.indexOf('ID'), ni=h.indexOf('Name'), di=h.indexOf('DOB'), ai=h.indexOf('DOA');
    let row=null;
    for (let i=1;i<rows.length;i++) { if(String(rows[i][ii]||'').trim()===String(contactId).trim()){row=rows[i];break;} }
    if (!row) return { success:false, error:'Contact not found.' };
    const name=String(row[ni]||'Unknown'), cal=CalendarApp.getDefaultCalendar(), created=[];
    const pDate=v=>{ if(!v)return null; const d=v instanceof Date?v:new Date(v); return isNaN(d.getTime())?null:d; };
    const yr=new Date().getFullYear();
    [['DOB','🎂 Birthday',di],['DOA','💍 Anniversary',ai]].forEach(([key,emoji,idx])=>{
      const d=pDate(row[idx]); if(!d)return;
      const evDate=new Date(yr,d.getMonth(),d.getDate());
      const title=`${emoji}: ${name}`;
      if(!cal.getEventsForDay(evDate).find(e=>e.getTitle()===title)){
        cal.createAllDayEvent(title,evDate,{recurrence:CalendarApp.newRecurrence().addYearlyRule(),description:`${key} of ${name} — SevakCall ID: ${contactId}`});
        created.push(key);
      }
    });
    logAct(callerEmail||'system','SYNC_CAL',`${name}: ${created.join(',')||'none'}`);
    return { success:true, created };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  DAILY GREETING EMAIL  (Time-driven trigger)
// ─────────────────────────────────────────────────────────────────
function sendDailyGreetings() {
  try {
    const today=new Date(), m=today.getMonth()+1, d=today.getDate();
    const rows=getSheet(CONTACTS_SHEET).getDataRange().getValues(), h=rows[0];
    const ni=h.indexOf('Name'),pi=h.indexOf('Phone'),di=h.indexOf('DOB'),ai=h.indexOf('DOA'),xi=h.indexOf('Area');
    const bdays=[], annivs=[];
    const match=v=>{ if(!v)return false; const dd=v instanceof Date?v:new Date(v); return !isNaN(dd.getTime())&&dd.getMonth()+1===m&&dd.getDate()===d; };
    rows.slice(1).forEach(r=>{
      const nm=String(r[ni]||'').trim(); if(!nm)return;
      const obj={name:nm,phone:String(r[pi]||''),area:String(r[xi]||'')};
      if(match(r[di]))bdays.push(obj);
      if(match(r[ai]))annivs.push(obj);
    });
    if(!bdays.length&&!annivs.length) return { success:true, sent:false };
    let adminEmail=Session.getActiveUser().getEmail();
    try { const vr=getSheet(VOLUNTEERS_SHEET).getDataRange().getValues(); const ar=vr.slice(1).find(r=>String(r[3]||'').toLowerCase()==='admin'); if(ar)adminEmail=String(ar[1]); } catch(_) {}
    const ds=Utilities.formatDate(today,Session.getScriptTimeZone(),'EEEE, dd MMMM yyyy');
    const tbl=(arr,bg)=>`<table style="width:100%;border-collapse:collapse;margin-bottom:20px"><tr style="background:${bg}"><th style="padding:8px;text-align:left">Name</th><th>Phone</th><th>Area</th></tr>${arr.map(x=>`<tr style="border-bottom:1px solid #eee"><td style="padding:8px">${x.name}</td><td style="padding:8px">${x.phone}</td><td style="padding:8px">${x.area}</td></tr>`).join('')}</table>`;
    const html=`<div style="font-family:Arial;max-width:600px;margin:auto;border:1px solid #eee;border-radius:12px;overflow:hidden"><div style="background:#FF6B1A;padding:20px;text-align:center"><h1 style="color:#fff;margin:0">🙏 SevakCall Daily Report</h1><p style="color:rgba(255,255,255,0.85);margin:4px 0 0">${ds}</p></div><div style="padding:24px">${bdays.length?`<h2 style="color:#1F9E5B">🎂 Birthdays (${bdays.length})</h2>${tbl(bdays,'#F0FAF4')}`:''} ${annivs.length?`<h2 style="color:#7C3AED">💍 Anniversaries (${annivs.length})</h2>${tbl(annivs,'#EDE9FE')}`:''}</div></div>`;
    GmailApp.sendEmail(adminEmail,`🙏 SevakCall — ${bdays.length} Birthdays, ${annivs.length} Anniversaries Today`,'Enable HTML.',{htmlBody:html});
    logAct('SYSTEM','DAILY_EMAIL',`${bdays.length} DOB, ${annivs.length} DOA`);
    return { success:true, sent:true, birthdays:bdays.length, anniversaries:annivs.length };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  EVENTS
// ─────────────────────────────────────────────────────────────────
function getEvents() {
  try {
    const rows=getSheet(EVENTS_SHEET).getDataRange().getValues();
    if(rows.length<2) return { events:[] };
    const events=rows.slice(1).map(r=>{
      let t=String(r[2]||'').trim(); if(r[2] instanceof Date) t=Utilities.formatDate(r[2],Session.getScriptTimeZone(),'HH:mm');
      let dt=''; if(r[1]){const dd=new Date(r[1]); if(!isNaN(dd.getTime()))dt=Utilities.formatDate(dd,Session.getScriptTimeZone(),'yyyy-MM-dd'); else dt=String(r[1]).trim();}
      return {title:String(r[0]||'').trim(),date:dt,time:t,duration:parseInt(r[3])||120,speaker:String(r[4]||'').trim(),columnName:String(r[5]||'').trim()};
    }).filter(e=>e.columnName);
    return { events };
  } catch(err) { return { events:[], error:err.message }; }
}

function createEvent(title, date, time, duration, speaker, callerEmail) {
  try {
    if(!title||!date||!time) return { success:false, error:'Title, date, and time required.' };
    const safeName='Sabha_'+String(date).replace(/[^0-9\-]/g,'').substring(0,10)+'_'+String(title).replace(/[^a-zA-Z0-9]/g,'').substring(0,5);
    const cSheet=getSheet(CONTACTS_SHEET), hds=cSheet.getRange(1,1,1,cSheet.getLastColumn()).getValues()[0];
    if(hds.includes(safeName)) return { success:false, error:`Column "${safeName}" already exists.` };
    getSheet(EVENTS_SHEET).appendRow([title,date,time,duration||120,speaker||'TBA',safeName]);
    const nc=cSheet.getLastColumn()+1;
    cSheet.getRange(1,nc).setValue(safeName).setBackground('#7C3AED').setFontColor('#FFFFFF').setFontWeight('bold');
    logAct(callerEmail||'admin','CREATE_SABHA',`${title} (${date})`);
    return { success:true, columnName:safeName, title, date };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  MARK ATTENDANCE
// ─────────────────────────────────────────────────────────────────
function markAttendance(contactId, columnName, callerEmail) {
  try {
    if(!contactId||!columnName) return { success:false, error:'contactId and columnName required.' };
    const sheet=getSheet(CONTACTS_SHEET), rows=sheet.getDataRange().getValues(), h=rows[0];
    const ii=h.indexOf('ID'), ei=h.indexOf(columnName), ni=h.indexOf('Name');
    if(ii<0) return { success:false, error:'ID column not found.' };
    if(ei<0) return { success:false, error:`Column "${columnName}" not found.` };
    for(let i=1;i<rows.length;i++){
      if(String(rows[i][ii]).trim()===String(contactId).trim()){
        sheet.getRange(i+1,ei+1).setValue('Present');
        logAct(callerEmail||'','MARK_ATTN',`${String(rows[i][ni]||'')} → ${columnName}`);
        return { success:true };
      }
    }
    return { success:false, error:'Contact not found.' };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  ADD CONTACT & MARK PRESENT  (Sabha walk-in)
// ─────────────────────────────────────────────────────────────────
function addContactAndMark(cd, columnName, callerEmail) {
  try {
    if(!cd||!columnName) return { success:false, error:'contactData and columnName required.' };
    const phone=String(cd.Phone||'').replace(/\D/g,'');
    if(phone.length===10){
      const dup=checkDupPhone(phone,null);
      if(dup.isDuplicate){
        const mr=markAttendance(dup.existingId,columnName,callerEmail);
        return {...mr,wasExisting:true,existingId:dup.existingId,name:dup.existingName};
      }
    }
    const sheet=getSheet(CONTACTS_SHEET), hds=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    let ac='XX'; try{ const ar=getSheet(AREAS_SHEET).getDataRange().getValues(); const m=ar.slice(1).find(r=>String(r[0]||'').toLowerCase()===String(cd.Area||'').toLowerCase()); if(m)ac=String(m[1]||'XX'); }catch(_){}
    const newId=generateContactID(ac,cd.Name,phone);
    const newRow=hds.map(h=>{
      if(h==='ID')return newId; if(h==='Name')return String(cd.Name||'').trim();
      if(h==='Phone')return phone||String(cd.Phone||'').trim(); if(h==='Mandal')return String(cd.Mandal||'').trim();
      if(h==='Area')return String(cd.Area||'').trim(); if(h==='Note')return String(cd.Note||'Sabha Walk-in').trim();
      if(h===columnName)return 'Present'; return '';
    });
    sheet.appendRow(newRow);
    logAct(callerEmail||'','ADD_WALKIN',`${cd.Name} → ${columnName}`);
    return { success:true, newId, name:cd.Name };
  } catch(err) { return { success:false, error:err.message }; }
}

// ─────────────────────────────────────────────────────────────────
//  INSTALL DAILY TRIGGER  (run once manually)
// ─────────────────────────────────────────────────────────────────
function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t=>t.getHandlerFunction()==='sendDailyGreetings')
    .forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyGreetings').timeBased().everyDays(1).atHour(8).create();
}
// ─────────────────────────────────────────────────────────────────
//  ONE-TIME SCRIPT: BULK GENERATE IDs FOR EXISTING CONTACTS
// ─────────────────────────────────────────────────────────────────
function bulkGenerateMissingIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = ss.getSheetByName('Contacts');
  const aSheet = ss.getSheetByName('Areas');
  
  // 1. Get Area Codes mapped out (e.g., "Amer" -> "AM")
  const areaData = aSheet.getDataRange().getValues().slice(1);
  const areaMap = {};
  areaData.forEach(r => {
    if(r[0]) areaMap[String(r[0]).trim().toLowerCase()] = String(r[1]||'XX').trim().toUpperCase();
  });

  // 2. Read Contacts Sheet
  const data = cSheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('ID');
  const nameCol = headers.indexOf('Name');
  const phoneCol = headers.indexOf('Phone');
  const areaCol = headers.indexOf('Area');

  if(idCol === -1 || nameCol === -1 || phoneCol === -1) {
    throw new Error("Could not find ID, Name, or Phone columns. Check headers.");
  }

  const outValues = [];
  let generatedCount = 0;

  // 3. Loop through all rows
  for(let i = 1; i < data.length; i++) {
    let existingId = String(data[i][idCol]).trim();
    
    // If it already has an ID, leave it alone
    if (existingId) {
      outValues.push([existingId]);
      continue;
    }
    
    // Generate new ID
    let name = data[i][nameCol];
    let phone = data[i][phoneCol];
    let areaName = data[i][areaCol];
    let areaCode = areaMap[String(areaName).trim().toLowerCase()] || 'XX';
    
    const serial = i; // Row number acts as the serial
    const code   = String(areaCode).replace(/[^A-Z0-9]/g,'').substring(0,4);
    const nm3    = String(name||'UNK').toUpperCase().replace(/[^A-Z]/g,'').substring(0,3).padEnd(3,'X');
    const ph4    = String(phone||'0000').replace(/\D/g,'').slice(-4).padStart(4,'0');
    
    outValues.push([`${serial}-${code}-${nm3}-${ph4}`]);
    generatedCount++;
  }
  
  // 4. Write all IDs back to the sheet in one fast operation
  cSheet.getRange(2, idCol + 1, outValues.length, 1).setValues(outValues);
  
  // 5. Clear the cache so the app updates immediately
  try { CacheService.getScriptCache().remove('sheet_Contacts'); } catch(e){}
  
  SpreadsheetApp.getUi().alert(`✅ Success! Generated ${generatedCount} missing IDs.`);
}

// ─────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────
function getSheet(name) {
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if(!s) throw new Error(`Sheet "${name}" not found.`);
  return s;
}
function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function logAct(user, action, details) {
  try { getSheet(ACTIVITY_SHEET).appendRow([new Date(), String(user||''), String(action||''), String(details||'')]); } catch(_) {}
}