/**
 * LTC Home Care Management System - Google Apps Script backend
 * Bound script: paste this file into Extensions > Apps Script of the target Google Sheet.
 * No Spreadsheet ID is required; the script uses SpreadsheetApp.getActiveSpreadsheet().
 * Deploy as Web app: Execute as Me, Who has access: Anyone.
 */

const APP = {
  VERSION: '1.12.0',
  TZ: 'Asia/Bangkok',
  SESSION_HOURS: 8,
  DEFAULT_ADMIN: 'admin',
  DEFAULT_PASSWORD: '1234',
  SHEETS: {
    SETTINGS: ['Key','Value','UpdatedAt','UpdatedBy'],
    USERS: ['UserID','Username','PasswordHash','Salt','DisplayName','Role','Status','LastLogin','CreatedAt','UpdatedAt','MustChangePassword'],
    SESSIONS: ['Token','UserID','Username','Role','CreatedAt','ExpiresAt','LastSeen','Status'],
    PATIENTS: ['PatientID','CitizenID','Prefix','FirstName','LastName','BirthDate','Age','Sex','Phone','Address','Village','Subdistrict','District','Province','Latitude','Longitude','CareRight','ADLScore','PatientGroup','DependencyLevel','RiskLevel','Status','CM','CG','PrimaryCaregiver','CaregiverPhone','Diseases','DrugAllergy','Medications','StartDate','LastVisit','NextVisit','Consent','Notes','CreatedAt','CreatedBy','UpdatedAt','UpdatedBy'],
    ASSESSMENTS: ['AssessmentID','PatientID','AssessmentDate','AssessmentType','ADLScore','DependencyLevel','FallRisk','PressureSoreRisk','NutritionRisk','Cognition','Depression','PainScore','CaregiverBurden','EnvironmentRisk','Summary','Assessor','CreatedAt','CreatedBy','FallRiskScore','NutritionScore','SuicideRisk','Advice','AssessmentData','UpdatedAt','UpdatedBy'],
    CARE_PLANS: ['CarePlanID','PatientID','Version','PlanDate','StartDate','EndDate','Problems','Goals','Activities','Frequency','Responsible','Indicators','BudgetPlan','Status','ApprovedBy','ApprovedAt','ReviewDate','Notes','CreatedAt','CreatedBy','UpdatedAt','UpdatedBy'],
    VISITS: ['VisitID','PatientID','ScheduleDate','VisitDate','StartTime','EndTime','Visitor','VisitType','Status','Temperature','Pulse','Respiration','Systolic','Diastolic','SpO2','DTX','Weight','PainScore','Activities','Problems','Actions','Advice','Referral','NextVisit','Latitude','Longitude','PhotoURL','CMReview','Reviewer','ReviewedAt','Notes','CreatedAt','CreatedBy','UpdatedAt','UpdatedBy'],
    ALERTS: ['AlertID','PatientID','AlertDate','Level','Category','Message','Status','AssignedTo','AcknowledgedAt','Response','ClosedAt','CreatedBy','CreatedAt'],
    STAFF: ['StaffID','UserID','Prefix','FirstName','LastName','Role','BirthDate','Age','Phone','Address','HouseNo','Moo','Village','Subdistrict','District','Province','BankName','BankAccountName','TrainingDate','CertificateExpiry','Status','PatientCount','Notes','CreatedAt','CreatedBy','UpdatedAt','UpdatedBy'],
    EQUIPMENT: ['EquipmentID','EquipmentType','AssetNo','SerialNo','Status','PatientID','LoanDate','DueDate','ReturnDate','Condition','MaintenanceDue','Notes','CreatedAt','UpdatedAt'],
    BUDGETS: ['BudgetID','FiscalYear','Source','PatientID','CarePlanID','Category','Description','PlannedAmount','ActualAmount','DocumentNo','ExpenseDate','Status','Requester','Approver','ApprovedAt','Notes','CreatedAt','UpdatedAt'],
    CONTRACTS: ['ContractID','ContractNo','StaffID','ContractDate','StartDate','EndDate','DurationMonths','MonthlyAmount','TotalAmount','PaymentEveryMonths','DailyDeduction','Status','ScopeOverride','Notes','CreatedAt','CreatedBy','UpdatedAt','UpdatedBy'],
    AUDIT_LOG: ['LogID','Timestamp','Username','Role','Action','Entity','EntityID','Detail','IPAddress'],
    CODEBOOK: ['Category','Code','ThaiLabel','Description','Active']
  }
};

// ลดการอ่าน/เขียน Google Sheet ซ้ำในทุก API call เพื่อให้หน้าจอเปิดเร็วขึ้น
const PERFORMANCE = {
  READY_TTL: 21600,
  AUTH_TTL: 300,
  DATA_TTL: 300,
  DASHBOARD_TTL: 120,
  SESSION_TOUCH_TTL: 300,
  CACHE_CHUNK_BYTES: 80000,
  CACHEABLE_SHEETS: ['SETTINGS','PATIENTS','ASSESSMENTS','CARE_PLANS','VISITS','ALERTS','STAFF','EQUIPMENT','BUDGETS','CONTRACTS','CODEBOOK']
};

const CONTRACT_MASTER_DEFAULTS = {
  CONTRACT_PLACE: 'โรงพยาบาลส่งเสริมสุขภาพตำบลบ้านชมภูพานเหนือ',
  CONTRACT_ORGANIZATION: 'โรงพยาบาลส่งเสริมสุขภาพตำบลบ้านชมภูพานเหนือ',
  CONTRACT_ORG_SUBDISTRICT: 'สร้างค้อ',
  CONTRACT_ORG_DISTRICT: 'ภูพาน',
  CONTRACT_ORG_PROVINCE: 'สกลนคร',
  CONTRACT_EMPLOYER_NAME: '',
  CONTRACT_EMPLOYER_POSITION: 'ผู้อำนวยการโรงพยาบาลส่งเสริมสุขภาพตำบลบ้านชมภูพานเหนือ',
  CONTRACT_PROJECT_TITLE: 'โครงการจ้างเหมาบริการรายบุคคลเพื่อปฏิบัติงานดูแลผู้สูงอายุที่มีภาวะพึ่งพิงและบุคคลอื่นที่มีภาวะพึ่งพิง (Care giver)',
  CONTRACT_DEFAULT_START_DATE: '',
  CONTRACT_DEFAULT_END_DATE: '',
  CONTRACT_DEFAULT_DURATION_MONTHS: '',
  CONTRACT_DEFAULT_MONTHLY_AMOUNT: '1500',
  CONTRACT_DEFAULT_TOTAL_AMOUNT: '',
  CONTRACT_DEFAULT_PAYMENT_EVERY_MONTHS: '3',
  CONTRACT_DEFAULT_DAILY_DEDUCTION: '100',
  CONTRACT_DEFAULT_BANK_NAME: 'ธนาคารเพื่อการเกษตรและสหกรณ์การเกษตร',
  CONTRACT_DEFAULT_BANK_BRANCH: 'ภูพาน',
  CONTRACT_WITNESS1_NAME: 'นางศันสนีย์ มีพรม',
  CONTRACT_WITNESS1_POSITION: 'พยาบาลวิชาชีพชำนาญการ',
  CONTRACT_WITNESS2_NAME: 'นางคณัสชญาณ์ ประทุมภา',
  CONTRACT_WITNESS2_POSITION: 'เจ้าพนักงานสาธารณสุขปฏิบัติงาน',
  CONTRACT_DEFAULT_WORK_AREA: 'ตำบลสร้างค้อ',
  CONTRACT_DEFAULT_SCOPE: 'ปฏิบัติงานดูแลผู้สูงอายุที่มีภาวะพึ่งพิงและบุคคลอื่นที่มีภาวะพึ่งพิงในพื้นที่\nประเมินและคัดกรองผู้สูงอายุตามแบบประเมิน ADL\nรวบรวมข้อมูล เอกสาร และประสานงานกับชุมชนและ Care Manager\nบันทึกและส่งรายงานการเยี่ยมประจำเดือนตามรอบที่กำหนด',
  CONTRACT_GUARANTEE_NOTE: 'การกำหนดหลักประกันการปฏิบัติตามข้อตกลงให้เป็นดุลยพินิจของผู้สั่งจ้างตามความเหมาะสม'
};

const CONTRACT_MASTER_KEYS = Object.keys(CONTRACT_MASTER_DEFAULTS);

const DEFAULT_CG_USERS = ['งามฉัตร','สุลาวรรณ','บุญจัน','สันดอน','อัมพร','แสงจันทร์','จงกล','ศิราณี'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบ LTC')
    .addItem('ติดตั้ง/ซ่อมแซมโครงสร้างระบบ', 'setupSystem')
    .addItem('รีเซ็ตรหัส admin เป็น 1234', 'resetAdminPassword')
    .addItem('ล้าง Session ที่หมดอายุ', 'cleanupSessions')
    .addToUi();
}

function setupSystem() {
  clearAllDataCaches_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersBeforeSetup = ss.getSheetByName('USERS');
  const hadMustChangePassword = usersBeforeSetup && usersBeforeSetup.getLastColumn() && usersBeforeSetup.getRange(1,1,1,usersBeforeSetup.getLastColumn()).getValues()[0].includes('MustChangePassword');
  Object.keys(APP.SHEETS).forEach(name => ensureSheet_(ss, name, APP.SHEETS[name]));
  clearSensitiveStaffData_();
  seedSettings_();
  seedAdmin_();
  seedCgUsers_(!hadMustChangePassword);
  seedCodebook_();
  SpreadsheetApp.flush();
  clearAllDataCaches_();
  CacheService.getScriptCache().put(readyCacheKey_(), '1', PERFORMANCE.READY_TTL);
  return { ok: true, message: 'ติดตั้งโครงสร้างระบบ LTC เรียบร้อยแล้ว' };
}

// Full setup formats every sheet, so normal API calls use a cheap check.
function ensureSystemReady_() {
  const cache = CacheService.getScriptCache();
  if (cache.get(readyCacheKey_()) === '1') return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const patients = ss.getSheetByName('PATIENTS');
  const assessments = ss.getSheetByName('ASSESSMENTS');
  const staff = ss.getSheetByName('STAFF');
  const users = ss.getSheetByName('USERS');
  const missingCore = !patients || !ss.getSheetByName('USERS') || !ss.getSheetByName('SESSIONS');
  const missingPatientGroup = patients && !patients.getRange(1,1,1,Math.max(patients.getLastColumn(),1)).getValues()[0].includes('PatientGroup');
  const missingAssessmentData = assessments && !assessments.getRange(1,1,1,Math.max(assessments.getLastColumn(),1)).getValues()[0].includes('AssessmentData');
  const missingStaffUserId = staff && !staff.getRange(1,1,1,Math.max(staff.getLastColumn(),1)).getValues()[0].includes('UserID');
  const missingMustChangePassword = users && !users.getRange(1,1,1,Math.max(users.getLastColumn(),1)).getValues()[0].includes('MustChangePassword');
  if (missingCore || !assessments || !staff || !ss.getSheetByName('CONTRACTS') || missingPatientGroup || missingAssessmentData || missingStaffUserId || missingMustChangePassword) setupSystem();
  else clearSensitiveStaffData_();
  cache.put(readyCacheKey_(), '1', PERFORMANCE.READY_TTL);
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'ping';
    return jsonResponse_(route_(action, e.parameter || {}, e));
  } catch (err) {
    return jsonResponse_({ ok: false, error: safeError_(err) });
  }
}

function doPost(e) {
  try {
    const payload = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    return jsonResponse_(route_(payload.action || 'ping', payload, e));
  } catch (err) {
    return jsonResponse_({ ok: false, error: safeError_(err) });
  }
}

function route_(action, payload, event) {
  if (action === 'ping') return { ok: true, version: APP.VERSION, serverTime: now_(), appName: getSetting_('APP_NAME') || 'ระบบเยี่ยมบ้านผู้ป่วย LTC' };
  ensureSystemReady_();
  if (action === 'login') return login_(payload.username, payload.password, event);
  if (action === 'publicDashboard') return dashboard_(null, true);

  const auth = requireAuth_(payload.token);
  touchSession_(payload.token);
  const ctx = { username: auth.Username, displayName: auth.DisplayName || auth.Username, role: auth.Role, userId: auth.UserID, token: payload.token, mustChangePassword: isTruthy_(auth.MustChangePassword), ip: clientIp_(event) };

  if (ctx.mustChangePassword && !['bootstrap','me','changePassword','logout'].includes(action)) {
    throw new Error('กรุณาตั้งรหัสผ่านใหม่ก่อนใช้งานระบบ');
  }

  switch (action) {
    case 'bootstrap': return { ok:true, user:userSummary_(auth), data:dashboard_(ctx,false).data };
    case 'me': return { ok: true, user:userSummary_(auth) };
    case 'logout': return logout_(payload.token, ctx);
    case 'dashboard': return dashboard_(ctx, false);
    case 'listPatients': return listRecords_('PATIENTS', payload, ctx);
    case 'getPatient': return getRecord_('PATIENTS', payload.id, ctx);
    case 'savePatient': requireRole_(ctx, ['admin','manager','cm']); return saveRecord_('PATIENTS', 'PatientID', 'PT', payload.data, ctx);
    case 'deletePatient': return archiveRecord_('PATIENTS', 'PatientID', payload.id, ctx);
    case 'listVisits': return listRecords_('VISITS', payload, ctx);
    case 'saveVisit': return saveVisit_(payload.data, ctx);
    case 'uploadVisitPhotos': return uploadVisitPhotos_(payload.photos, payload.patientId, ctx);
    case 'getVisitPhotos': return getVisitPhotos_(payload.visitId, ctx);
    case 'listAssessments': return listRecords_('ASSESSMENTS', payload, ctx);
    case 'saveAssessment': return saveAssessment_(payload.data, ctx);
    case 'listCarePlans': return listRecords_('CARE_PLANS', payload, ctx);
    case 'saveCarePlan': requireRole_(ctx, ['admin','manager','cm']); return saveRecord_('CARE_PLANS', 'CarePlanID', 'CP', payload.data, ctx);
    case 'listAlerts': return listRecords_('ALERTS', payload, ctx);
    case 'saveAlert': return savePatientLinkedRecord_('ALERTS', 'AlertID', 'AL', payload.data, ctx);
    case 'listStaff': return listStaff_(payload, ctx);
    case 'saveStaff': requireRole_(ctx, ['admin','manager','cm']); return saveRecord_('STAFF', 'StaffID', 'ST', payload.data, ctx);
    case 'getMyStaffProfile': requireRole_(ctx, ['cg']); return getMyStaffProfile_(ctx);
    case 'saveMyStaffProfile': requireRole_(ctx, ['cg']); return saveMyStaffProfile_(payload.data, ctx);
    case 'getContractMaster': requireRole_(ctx, ['admin','manager','cm']); return getContractMaster_(ctx);
    case 'saveContractMaster': requireRole_(ctx, ['admin','manager','cm']); return saveContractMaster_(payload.data, ctx);
    case 'listContracts': requireRole_(ctx, ['admin','manager','cm']); return listRecords_('CONTRACTS', payload, ctx);
    case 'saveContract': requireRole_(ctx, ['admin','manager','cm']); return saveContract_(payload.data, ctx);
    case 'getContractPrintData': requireRole_(ctx, ['admin','manager','cm']); return getContractPrintData_(payload.id, ctx);
    case 'listEquipment': return listRecords_('EQUIPMENT', payload, ctx);
    case 'saveEquipment': return saveRecord_('EQUIPMENT', 'EquipmentID', 'EQ', payload.data, ctx);
    case 'listBudgets': requireRole_(ctx, ['admin','manager','cm']); return listRecords_('BUDGETS', payload, ctx);
    case 'saveBudget': requireRole_(ctx, ['admin','manager','cm']); return saveRecord_('BUDGETS', 'BudgetID', 'BG', payload.data, ctx);
    case 'listUsers': requireRole_(ctx, ['admin']); return listUsers_(ctx);
    case 'saveUser': requireRole_(ctx, ['admin']); return saveUser_(payload.data, ctx);
    case 'adminResetUserPassword': requireRole_(ctx, ['admin']); return adminResetUserPassword_(payload.data, ctx);
    case 'changePassword': return changePassword_(payload, ctx);
    case 'getSettings': requireRole_(ctx, ['admin']); return { ok: true, data: rowsAsObjects_('SETTINGS') };
    case 'saveSettings': requireRole_(ctx, ['admin']); return saveSettings_(payload.data, ctx);
    case 'auditLog': requireRole_(ctx, ['admin','manager']); return listRecords_('AUDIT_LOG', payload, ctx);
    default: throw new Error('ไม่พบคำสั่ง: ' + action);
  }
}

function login_(username, password, event) {
  username = String(username || '').trim().toLowerCase();
  const users = rowsAsObjects_('USERS');
  const user = users.find(u => String(u.Username).toLowerCase() === username && String(u.Status).toLowerCase() === 'active');
  if (!user || hashPassword_(String(password || ''), String(user.Salt)) !== String(user.PasswordHash)) {
    audit_({username: username || '-', role: '-', ip: clientIp_(event)}, 'LOGIN_FAILED', 'USERS', '', 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }
  cleanupSessions();
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const created = new Date();
  const expires = new Date(created.getTime() + APP.SESSION_HOURS * 3600000);
  appendObject_('SESSIONS', { Token: token, UserID: user.UserID, Username: user.Username, Role: user.Role, CreatedAt: created, ExpiresAt: expires, LastSeen: created, Status: 'active' });
  updateById_('USERS', 'UserID', user.UserID, { LastLogin: created, UpdatedAt: created });
  cacheAuth_(token, user, expires);
  CacheService.getScriptCache().put(touchCacheKey_(token),'1',PERFORMANCE.SESSION_TOUCH_TTL);
  audit_({username:user.Username, role:user.Role, ip:clientIp_(event)}, 'LOGIN_SUCCESS', 'USERS', user.UserID, 'เข้าสู่ระบบสำเร็จ');
  return { ok: true, token: token, expiresAt: formatDate_(expires), user:userSummary_(user) };
}

function requireAuth_(token) {
  if (!token) throw new Error('กรุณาเข้าสู่ระบบ');
  const cached = getCachedAuth_(token);
  if (cached) return cached;
  const session = rowsAsObjects_('SESSIONS').find(s => s.Token === token && s.Status === 'active');
  if (!session || new Date(session.ExpiresAt).getTime() <= Date.now()) throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const user = rowsAsObjects_('USERS').find(u => u.UserID === session.UserID && u.Status === 'active');
  if (!user) throw new Error('บัญชีผู้ใช้ถูกระงับ');
  cacheAuth_(token, user, session.ExpiresAt);
  CacheService.getScriptCache().put(touchCacheKey_(token),'1',PERFORMANCE.SESSION_TOUCH_TTL);
  return user;
}

function userSummary_(user) {
  return {
    username:user.Username,
    displayName:user.DisplayName || user.Username,
    role:user.Role,
    mustChangePassword:isTruthy_(user.MustChangePassword)
  };
}

function isTruthy_(value) {
  if (value === true || value === 1) return true;
  return ['true','1','yes','y','required','ต้องเปลี่ยน'].includes(String(value || '').trim().toLowerCase());
}

function isCg_(ctx) {
  return !!ctx && String(ctx.role || '').trim().toLowerCase() === 'cg';
}

function normalizeAssignee_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g,' ');
}

function isPatientAssignedTo_(patient, ctx) {
  if (!isCg_(ctx)) return true;
  const assigned = normalizeAssignee_(patient && patient.CG);
  if (!assigned) return false;
  const identities = [ctx.username,ctx.displayName].map(normalizeAssignee_).filter(Boolean);
  return identities.some(identity => assigned === identity || assigned.includes(identity));
}

function assignedPatientIds_(ctx) {
  return new Set(rowsAsObjects_('PATIENTS').filter(patient => isPatientAssignedTo_(patient,ctx)).map(patient => String(patient.PatientID)));
}

function assertPatientAccess_(patientId, ctx) {
  if (!isCg_(ctx)) return;
  if (!patientId || !assignedPatientIds_(ctx).has(String(patientId))) throw new Error('คุณไม่มีสิทธิ์เข้าถึงผู้ป่วยรายนี้');
}

function requireRole_(ctx, roles) {
  if (!roles.includes(String(ctx.role).toLowerCase())) throw new Error('คุณไม่มีสิทธิ์ใช้คำสั่งนี้');
}

function logout_(token, ctx) {
  updateById_('SESSIONS','Token',token,{Status:'closed',LastSeen:new Date()});
  CacheService.getScriptCache().remove(authCacheKey_(token));
  audit_(ctx,'LOGOUT','SESSIONS',token.substring(0,8),'ออกจากระบบ');
  return {ok:true};
}

function isCurrentCarePatient_(patient) {
  return !['archived','เสียชีวิต','deceased','ย้ายพื้นที่','ส่งต่อ','referred'].includes(String(patient && patient.Status || '').trim().toLowerCase());
}

function cgWorkload_(patients, ctx) {
  const currentPatients = (patients || []).filter(isCurrentCarePatient_);
  if (isCg_(ctx)) {
    return {
      items:[{name:visitorNameForCtx_(ctx) || ctx.displayName || ctx.username,count:currentPatients.length}],
      assigned:currentPatients.length,
      unassigned:0
    };
  }

  const users = rowsAsObjects_('USERS').filter(user => String(user.Role || '').trim().toLowerCase() === 'cg' && String(user.Status || 'active').trim().toLowerCase() === 'active');
  const usersById = users.reduce((map,user) => { map[String(user.UserID || '')] = user; return map; },{});
  const entries = [];
  const addEntry = (name, aliases) => {
    const displayName = String(name || '').trim();
    const normalized = [...new Set((aliases || []).concat([displayName]).map(normalizeAssignee_).filter(Boolean))];
    if (!displayName || !normalized.length) return null;
    const existing = entries.find(entry => entry.aliases.some(alias => normalized.includes(alias)));
    if (existing) {
      normalized.forEach(alias => { if (!existing.aliases.includes(alias)) existing.aliases.push(alias); });
      return existing;
    }
    const entry = {name:displayName,count:0,aliases:normalized};
    entries.push(entry);
    return entry;
  };

  rowsAsObjects_('STAFF').filter(staff => String(staff.Role || '').trim().toLowerCase() === 'cg' && !['inactive','archived','ระงับ'].includes(String(staff.Status || '').trim().toLowerCase())).forEach(staff => {
    const fullName = (String(staff.Prefix || '')+String(staff.FirstName || '')+' '+String(staff.LastName || '')).trim();
    const user = usersById[String(staff.UserID || '')];
    addEntry(fullName || (user && (user.DisplayName || user.Username)),[
      staff.FirstName,
      (String(staff.FirstName || '')+' '+String(staff.LastName || '')).trim(),
      user && user.DisplayName,
      user && user.Username
    ]);
  });
  users.forEach(user => addEntry(user.DisplayName || user.Username,[user.DisplayName,user.Username]));

  let unassigned = 0;
  currentPatients.forEach(patient => {
    const assigned = normalizeAssignee_(patient.CG);
    if (!assigned) { unassigned++; return; }
    let entry = entries.find(item => item.aliases.some(alias => assigned === alias || assigned.includes(alias) || alias.includes(assigned)));
    if (!entry) entry = addEntry(String(patient.CG).trim(),[patient.CG]);
    if (entry) entry.count++;
  });
  const items = entries.map(entry => ({name:entry.name,count:entry.count})).sort((a,b) => b.count-a.count || String(a.name).localeCompare(String(b.name),'th'));
  return {items:items,assigned:items.reduce((sum,item) => sum+item.count,0),unassigned:unassigned};
}

function dashboard_(ctx, isPublic) {
  const cache = CacheService.getScriptCache();
  const cacheKey = dashboardCacheKey_(isPublic);
  const scopedForCg = isCg_(ctx);
  if (!scopedForCg) {
    try {
      const cached = cache.get(cacheKey);
      if (cached) return {ok:true,data:JSON.parse(cached),public:isPublic,cached:true};
    } catch (e) {}
  }
  let patients = rowsAsObjects_('PATIENTS').filter(p => String(p.Status).toLowerCase() !== 'archived');
  if (scopedForCg) patients = patients.filter(patient => isPatientAssignedTo_(patient,ctx));
  const patientIds = new Set(patients.map(patient => String(patient.PatientID)));
  const visits = rowsAsObjects_('VISITS').filter(row => !scopedForCg || patientIds.has(String(row.PatientID)));
  const alerts = rowsAsObjects_('ALERTS').filter(row => !scopedForCg || patientIds.has(String(row.PatientID)));
  const today = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd');
  const levels = countBy_(patients, 'DependencyLevel');
  const risks = countBy_(patients, 'RiskLevel');
  const result = {
    totalPatients: patients.length,
    activePatients: patients.filter(p => ['active','กำลังดูแล'].includes(String(p.Status).toLowerCase())).length,
    highRisk: patients.filter(p => ['สูง','แดง','high','red'].includes(String(p.RiskLevel).toLowerCase())).length,
    visitsToday: visits.filter(v => dateOnly_(v.ScheduleDate || v.VisitDate) === today).length,
    overdueVisits: visits.filter(v => dateOnly_(v.ScheduleDate) < today && !['เยี่ยมสำเร็จ','completed','ยกเลิก'].includes(String(v.Status).toLowerCase())).length,
    openAlerts: alerts.filter(a => !['closed','ปิดเหตุ'].includes(String(a.Status).toLowerCase())).length,
    dependency: levels,
    risk: risks,
    updatedAt: now_()
  };
  if (!isPublic && ctx) result.cgWorkload = cgWorkload_(patients,ctx);
  if (!isPublic && ctx && !scopedForCg) {
    const equipment = rowsAsObjects_('EQUIPMENT');
    const budgets = rowsAsObjects_('BUDGETS');
    const plannedBudget = budgets.reduce((sum, b) => sum + (Number(b.PlannedAmount) || 0), 0);
    const actualBudget = budgets.reduce((sum, b) => sum + (Number(b.ActualAmount) || 0), 0);
    result.equipment = {
      total: equipment.length,
      available: equipment.filter(x => ['พร้อมใช้','available'].includes(String(x.Status).toLowerCase())).length,
      loaned: equipment.filter(x => ['ยืม','loaned','in use'].includes(String(x.Status).toLowerCase())).length,
      overdue: equipment.filter(x => dateOnly_(x.DueDate) && dateOnly_(x.DueDate) < today && !x.ReturnDate).length,
      maintenanceDue: equipment.filter(x => dateOnly_(x.MaintenanceDue) && dateOnly_(x.MaintenanceDue) <= today).length
    };
    result.budget = {
      planned: plannedBudget,
      actual: actualBudget,
      remaining: plannedBudget - actualBudget,
      percentUsed: plannedBudget > 0 ? Math.round(actualBudget / plannedBudget * 1000) / 10 : 0
    };
  }
  if (!scopedForCg) {
    try { cache.put(cacheKey,JSON.stringify(result),PERFORMANCE.DASHBOARD_TTL); } catch (e) {}
  }
  return { ok:true, data:result, public:isPublic };
}

function listRecords_(sheetName, payload, ctx) {
  let rows = rowsAsObjects_(sheetName);
  if (isCg_(ctx)) {
    if (sheetName === 'PATIENTS') rows = rows.filter(row => isPatientAssignedTo_(row,ctx));
    else if (APP.SHEETS[sheetName] && APP.SHEETS[sheetName].includes('PatientID')) {
      const allowed = assignedPatientIds_(ctx);
      rows = rows.filter(row => allowed.has(String(row.PatientID)));
    }
  }
  const idField = APP.SHEETS[sheetName][0];
  if (payload.id) rows = rows.filter(r => String(r[idField]) === String(payload.id));
  const q = String(payload.q || '').trim().toLowerCase();
  if (q) rows = rows.filter(r => Object.values(r).some(v => String(v || '').toLowerCase().includes(q)));
  if (payload.patientId) rows = rows.filter(r => String(r.PatientID) === String(payload.patientId));
  if (sheetName === 'VISITS' && payload.month) {
    const month = String(payload.month).trim();
    if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('รูปแบบเดือนไม่ถูกต้อง');
    rows = rows.filter(r => dateOnly_(r.VisitDate || r.ScheduleDate).substring(0,7) === month);
  }
  if (payload.status) rows = rows.filter(r => String(r.Status).toLowerCase() === String(payload.status).toLowerCase());
  const total = rows.length;
  const limit = Math.min(Math.max(Number(payload.limit || 200), 1), 1000);
  const offset = Math.max(Number(payload.offset || 0), 0);
  rows = rows.slice(offset, offset + limit);
  if (Array.isArray(payload.fields) && payload.fields.length) {
    const allowed = new Set(APP.SHEETS[sheetName]);
    const fields = payload.fields.filter(field => allowed.has(field));
    rows = rows.map(row => fields.reduce((item,field) => { item[field]=row[field]; return item; },{}));
  }
  // Do not append an audit row for every list refresh; that extra Sheet write
  // noticeably delays read-only screens. Create/update/login remain audited.
  return {ok:true,data:rows,total:total,limit:limit,offset:offset};
}

function clearSensitiveStaffData_() {
  const props = PropertiesService.getScriptProperties();
  const migrationKey = 'LTC_STAFF_SENSITIVE_FIELDS_REMOVED_V1';
  if (props.getProperty(migrationKey) === '1') return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STAFF');
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
  const rows = Math.max(sheet.getLastRow()-1,0);
  if (rows) ['CitizenID','BankBranch','BankAccountNo'].forEach(field => {
    const col = headers.indexOf(field);
    if (col >= 0) sheet.getRange(2,col+1,rows,1).clearContent();
  });
  clearDataCache_('STAFF');
  props.setProperty(migrationKey,'1');
}

function safeStaffRecord_(row) {
  const safe = Object.assign({},row);
  delete safe.CitizenID;
  delete safe.BankBranch;
  delete safe.BankAccountNo;
  return safe;
}

// CG ต้องใช้รายชื่อบุคลากรสำหรับเลือกผู้เยี่ยม แต่ไม่ควรได้รับที่อยู่หรือข้อมูลธนาคารของบุคลากรอื่น
function listStaff_(payload, ctx) {
  const privileged = ['admin','manager','cm'].includes(String(ctx.role || '').toLowerCase());
  const allowed = privileged
    ? APP.SHEETS.STAFF.slice()
    : ['StaffID','Prefix','FirstName','LastName','Role','Phone','Village','TrainingDate','CertificateExpiry','Status','PatientCount'];
  let rows = rowsAsObjects_('STAFF').map(safeStaffRecord_).map(row => allowed.reduce((safe, key) => { safe[key] = row[key]; return safe; }, {}));
  if (payload.id) rows = rows.filter(row => String(row.StaffID) === String(payload.id));
  const q = String(payload.q || '').trim().toLowerCase();
  if (q) rows = rows.filter(row => Object.values(row).some(value => String(value || '').toLowerCase().includes(q)));
  if (payload.status) rows = rows.filter(row => String(row.Status).toLowerCase() === String(payload.status).toLowerCase());
  const total = rows.length;
  const limit = Math.min(Math.max(Number(payload.limit || 200), 1), 1000);
  const offset = Math.max(Number(payload.offset || 0), 0);
  rows = rows.slice(offset,offset+limit);
  if (Array.isArray(payload.fields) && payload.fields.length) {
    const visible = new Set(allowed);
    const fields = payload.fields.filter(field => visible.has(field));
    rows = rows.map(row => fields.reduce((item,field) => { item[field]=row[field]; return item; },{}));
  }
  return {ok:true,data:rows,total:total,limit:limit,offset:offset};
}

// ผูกบัญชี CG กับระเบียนบุคลากรของตนเอง เพื่อให้แก้ไขข้อมูลส่วนตัวได้โดยไม่เห็นข้อมูลลับของบุคลากรคนอื่น
function findMyStaffProfile_(ctx) {
  const rows = rowsAsObjects_('STAFF');
  const byUserId = rows.find(row => String(row.UserID || '') === String(ctx.userId));
  if (byUserId) return byUserId;
  const identities = [ctx.username,ctx.displayName].map(normalizeAssignee_).filter(Boolean);
  return rows.find(row => {
    if (row.UserID || String(row.Role || '').trim().toLowerCase() !== 'cg') return false;
    const firstName = normalizeAssignee_(row.FirstName);
    const fullName = normalizeAssignee_((String(row.FirstName || '')+' '+String(row.LastName || '')).trim());
    return identities.some(identity => identity === firstName || identity === fullName);
  }) || null;
}

function visitorNameForCtx_(ctx) {
  if (!isCg_(ctx)) return '';
  const profile = findMyStaffProfile_(ctx);
  if (profile) {
    const fullName = (String(profile.Prefix || '')+String(profile.FirstName || '')+' '+String(profile.LastName || '')).trim();
    if (fullName) return fullName;
  }
  return String(ctx.displayName || ctx.username || '').trim();
}

function ensureMyStaffProfile_(ctx) {
  let profile = findMyStaffProfile_(ctx);
  if (profile) {
    if (!profile.UserID) {
      updateById_('STAFF','StaffID',profile.StaffID,{UserID:ctx.userId,UpdatedAt:new Date(),UpdatedBy:ctx.username});
      profile = Object.assign({},profile,{UserID:ctx.userId});
    }
    return profile;
  }
  const now = new Date();
  const row = {
    StaffID:makeId_('ST'),UserID:ctx.userId,FirstName:ctx.displayName || ctx.username,LastName:'',Role:'CG',Status:'active',
    CreatedAt:now,CreatedBy:ctx.username,UpdatedAt:now,UpdatedBy:ctx.username
  };
  appendObject_('STAFF',row);
  audit_(ctx,'CREATE_SELF_PROFILE','STAFF',row.StaffID,'สร้างข้อมูลบุคลากรสำหรับบัญชี CG');
  return row;
}

function getMyStaffProfile_(ctx) {
  const profile = ensureMyStaffProfile_(ctx);
  audit_(ctx,'READ_SELF_PROFILE','STAFF',profile.StaffID,'เปิดดูข้อมูลบุคลากรของตนเอง');
  return {ok:true,data:safeStaffRecord_(profile)};
}

function saveMyStaffProfile_(data, ctx) {
  if (!data || typeof data !== 'object') throw new Error('ข้อมูลส่วนตัวไม่ถูกต้อง');
  const profile = ensureMyStaffProfile_(ctx);
  const allowed = ['Prefix','FirstName','LastName','BirthDate','Age','Phone','Address','HouseNo','Moo','Village','Subdistrict','District','Province','BankName','BankAccountName','TrainingDate','CertificateExpiry','Notes'];
  const clean = sanitizeObject_(data,allowed);
  const firstName = String(clean.FirstName !== undefined ? clean.FirstName : profile.FirstName || '').trim();
  const lastName = String(clean.LastName !== undefined ? clean.LastName : profile.LastName || '').trim();
  if (!firstName || !lastName) throw new Error('กรุณาระบุชื่อและนามสกุลให้ครบ');
  clean.UpdatedAt = new Date();
  clean.UpdatedBy = ctx.username;
  if (!updateById_('STAFF','StaffID',profile.StaffID,clean)) throw new Error('ไม่พบข้อมูลบุคลากรของคุณ');
  audit_(ctx,'UPDATE_SELF_PROFILE','STAFF',profile.StaffID,'แก้ไขข้อมูลบุคลากรของตนเอง');
  const updated = rowsAsObjects_('STAFF').find(row => String(row.StaffID) === String(profile.StaffID));
  return {ok:true,id:profile.StaffID,data:safeStaffRecord_(updated || Object.assign({},profile,clean))};
}

function getRecord_(sheetName, id, ctx) {
  const idField = APP.SHEETS[sheetName][0];
  const row = rowsAsObjects_(sheetName).find(r => String(r[idField]) === String(id));
  if (!row) throw new Error('ไม่พบข้อมูล');
  if (isCg_(ctx)) {
    if (sheetName === 'PATIENTS' && !isPatientAssignedTo_(row,ctx)) throw new Error('คุณไม่มีสิทธิ์ดูข้อมูลผู้ป่วยรายนี้');
    if (sheetName !== 'PATIENTS' && Object.prototype.hasOwnProperty.call(row,'PatientID')) assertPatientAccess_(row.PatientID,ctx);
  }
  audit_(ctx,'READ',sheetName,id,'เปิดดูข้อมูล');
  return {ok:true,data:row};
}

function savePatientLinkedRecord_(sheetName, idField, prefix, data, ctx) {
  if (!data || !data.PatientID) throw new Error('กรุณาเลือกผู้ป่วย');
  assertPatientAccess_(data.PatientID,ctx);
  return saveRecord_(sheetName,idField,prefix,data,ctx);
}

function saveRecord_(sheetName, idField, prefix, data, ctx) {
  if (!data || typeof data !== 'object') throw new Error('ข้อมูลไม่ถูกต้อง');
  const now = new Date();
  const clean = sanitizeObject_(data, APP.SHEETS[sheetName]);
  if (isCg_(ctx) && APP.SHEETS[sheetName].includes('PatientID')) {
    if (clean[idField]) {
      const existing = rowsAsObjects_(sheetName).find(row => String(row[idField]) === String(clean[idField]));
      if (!existing) throw new Error('ไม่พบข้อมูลสำหรับแก้ไข');
      assertPatientAccess_(existing.PatientID,ctx);
    }
    if (clean.PatientID) assertPatientAccess_(clean.PatientID,ctx);
  }
  if (!clean[idField]) {
    clean[idField] = makeId_(prefix);
    clean.CreatedAt = now;
    clean.CreatedBy = ctx.username;
    clean.UpdatedAt = now;
    clean.UpdatedBy = ctx.username;
    appendObject_(sheetName, clean);
    audit_(ctx,'CREATE',sheetName,clean[idField],'สร้างข้อมูล');
  } else {
    clean.UpdatedAt = now;
    clean.UpdatedBy = ctx.username;
    if (!updateById_(sheetName,idField,clean[idField],clean)) throw new Error('ไม่พบข้อมูลสำหรับแก้ไข');
    audit_(ctx,'UPDATE',sheetName,clean[idField],'แก้ไขข้อมูล');
  }
  return {ok:true,id:clean[idField],data:clean};
}

function saveVisit_(data, ctx) {
  if (!data || !data.PatientID) throw new Error('กรุณาเลือกผู้ป่วย');
  const visitData = Object.assign({},data);
  if (isCg_(ctx)) visitData.Visitor = visitorNameForCtx_(ctx);
  if (!String(visitData.Visitor || '').trim()) throw new Error('ไม่พบชื่อผู้เยี่ยม');
  assertPatientAccess_(visitData.PatientID,ctx);
  const result = saveRecord_('VISITS','VisitID','VS',visitData,ctx);
  if (visitData.PatientID && visitData.VisitDate) updateById_('PATIENTS','PatientID',visitData.PatientID,{LastVisit:visitData.VisitDate,NextVisit:visitData.NextVisit || '',UpdatedAt:new Date(),UpdatedBy:ctx.username});
  const alerts = detectAlerts_(visitData, ctx);
  result.alerts = alerts;
  return result;
}

// รูปหลักฐานเก็บเป็นไฟล์ส่วนตัวใน Google Drive และเปิดอ่านผ่าน API หลังยืนยันตัวตนเท่านั้น
function uploadVisitPhotos_(photos, patientId, ctx) {
  assertPatientAccess_(patientId,ctx);
  if (!Array.isArray(photos) || !photos.length) throw new Error('ไม่พบรูปภาพสำหรับอัปโหลด');
  if (photos.length > 4) throw new Error('แนบรูปได้ครั้งละไม่เกิน 4 รูป');
  const folder = getVisitPhotoFolder_();
  const allowed = ['image/jpeg','image/png','image/webp'];
  const saved = photos.map((photo, index) => {
    const match = String(photo && photo.dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
    if (!match || !allowed.includes(match[1])) throw new Error('รูปที่ '+(index+1)+' ไม่ใช่ไฟล์ JPG, PNG หรือ WEBP');
    const bytes = Utilities.base64Decode(match[2]);
    if (bytes.length > 1500000) throw new Error('รูปที่ '+(index+1)+' มีขนาดใหญ่เกิน 1.5 MB');
    const ext = match[1] === 'image/png' ? '.png' : match[1] === 'image/webp' ? '.webp' : '.jpg';
    const base = String(photo.name || 'visit-photo').replace(/\.[^.]+$/, '').replace(/[^a-zA-Z0-9ก-๙_-]/g, '_').substring(0,60) || 'visit-photo';
    const fileName = Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd-HHmmss')+'_'+String(patientId || 'patient').replace(/[^a-zA-Z0-9_-]/g,'')+'_'+(index+1)+'_'+base+ext;
    const file = folder.createFile(Utilities.newBlob(bytes, match[1], fileName));
    file.setDescription('ภาพหลักฐานการเยี่ยมบ้านผู้ป่วย LTC อัปโหลดโดย '+ctx.username);
    return {fileId:file.getId(),name:fileName,mimeType:match[1]};
  });
  audit_(ctx,'UPLOAD_PHOTO','VISITS',String(patientId || ''),'แนบรูปหลักฐาน '+saved.length+' รูป');
  return {ok:true,data:saved};
}

function getVisitPhotos_(visitId, ctx) {
  if (!visitId) throw new Error('ไม่พบรหัสการเยี่ยมบ้าน');
  const visit = rowsAsObjects_('VISITS').find(v => String(v.VisitID) === String(visitId));
  if (!visit) throw new Error('ไม่พบข้อมูลการเยี่ยมบ้าน');
  assertPatientAccess_(visit.PatientID,ctx);
  const refs = parseVisitPhotoRefs_(visit.PhotoURL);
  const data = refs.slice(0,8).map(ref => {
    try {
      const file = DriveApp.getFileById(ref.fileId);
      const blob = file.getBlob();
      const mimeType = blob.getContentType() || ref.mimeType || 'image/jpeg';
      return {name:ref.name || file.getName(),dataUrl:'data:'+mimeType+';base64,'+Utilities.base64Encode(blob.getBytes())};
    } catch (e) { return null; }
  }).filter(Boolean);
  audit_(ctx,'READ_PHOTO','VISITS',visitId,'เปิดรูปหลักฐานการเยี่ยมบ้าน');
  return {ok:true,data:data};
}

function parseVisitPhotoRefs_(value) {
  if (!value) return [];
  try {
    const refs = typeof value === 'string' ? JSON.parse(value) : value;
    return Array.isArray(refs) ? refs.filter(x => x && x.fileId) : [];
  } catch (e) { return []; }
}

function getVisitPhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const key = 'LTC_VISIT_PHOTO_FOLDER_ID';
  const existing = props.getProperty(key);
  if (existing) {
    try { return DriveApp.getFolderById(existing); } catch (e) {}
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const safeName = String(ss.getName() || 'LTC').replace(/[\\/:*?"<>|]/g,'-').substring(0,80);
  const folder = DriveApp.createFolder('LTC Visit Evidence - '+safeName);
  props.setProperty(key, folder.getId());
  return folder;
}

function getContractMaster_(ctx) {
  const settings = rowsAsObjects_('SETTINGS').reduce((map, row) => {
    map[String(row.Key)] = row.Value;
    return map;
  }, {});
  const data = {};
  CONTRACT_MASTER_KEYS.forEach(key => data[key] = settings[key] !== undefined && settings[key] !== '' ? settings[key] : CONTRACT_MASTER_DEFAULTS[key]);
  return {ok:true,data:data};
}

function saveContractMaster_(data, ctx) {
  if (!data || typeof data !== 'object') throw new Error('ข้อมูลสัญญาหลักไม่ถูกต้อง');
  const items = CONTRACT_MASTER_KEYS.filter(key => Object.prototype.hasOwnProperty.call(data,key)).map(key => ({Key:key,Value:sanitizeValue_(data[key])}));
  if (!items.length) throw new Error('ไม่พบข้อมูลสัญญาหลักสำหรับบันทึก');
  saveSettings_(items,ctx);
  audit_(ctx,'UPDATE','CONTRACT_MASTER','','แก้ไขข้อมูลกลางของบันทึกข้อตกลงจ้าง');
  return getContractMaster_(ctx);
}

function saveContract_(data, ctx) {
  if (!data || typeof data !== 'object') throw new Error('ข้อมูลข้อตกลงจ้างไม่ถูกต้อง');
  if (!data.StaffID) throw new Error('กรุณาเลือก CG ผู้รับจ้าง');
  const staff = rowsAsObjects_('STAFF').find(s => String(s.StaffID) === String(data.StaffID));
  if (!staff) throw new Error('ไม่พบข้อมูล CG ที่เลือก');
  if (String(staff.Role || '').trim().toLowerCase() !== 'cg') throw new Error('บันทึกข้อตกลงจ้างนี้ใช้สำหรับบุคลากรบทบาท CG');
  const clean = Object.assign({},data);
  ['DurationMonths','MonthlyAmount','TotalAmount','PaymentEveryMonths','DailyDeduction'].forEach(key => {
    if (clean[key] === '' || clean[key] == null) clean[key] = '';
    else {
      const number = Number(clean[key]);
      if (!isFinite(number) || number < 0) throw new Error('ค่าตัวเลขในข้อตกลงจ้างไม่ถูกต้อง');
      clean[key] = number;
    }
  });
  if ((clean.TotalAmount === '' || clean.TotalAmount == null) && Number(clean.MonthlyAmount) > 0 && Number(clean.DurationMonths) > 0) clean.TotalAmount = Number(clean.MonthlyAmount) * Number(clean.DurationMonths);
  if (!clean.ContractDate) clean.ContractDate = dateOnly_(new Date());
  if (!clean.Status) clean.Status = 'ร่าง';
  return saveRecord_('CONTRACTS','ContractID','CT',clean,ctx);
}

function getContractPrintData_(id, ctx) {
  if (!id) throw new Error('ไม่พบรหัสข้อตกลงจ้าง');
  const contract = rowsAsObjects_('CONTRACTS').find(row => String(row.ContractID) === String(id));
  if (!contract) throw new Error('ไม่พบข้อตกลงจ้าง');
  const staff = rowsAsObjects_('STAFF').find(row => String(row.StaffID) === String(contract.StaffID));
  if (!staff) throw new Error('ไม่พบข้อมูล CG ของข้อตกลงนี้');
  const master = getContractMaster_(ctx).data;
  audit_(ctx,'PRINT','CONTRACTS',id,'เตรียมพิมพ์บันทึกข้อตกลงจ้างรายบุคคล');
  return {ok:true,data:{contract:contract,staff:safeStaffRecord_(staff),master:master}};
}

function saveAssessment_(data, ctx) {
  if (!data || typeof data !== 'object') throw new Error('ข้อมูลแบบประเมินไม่ถูกต้อง');
  if (!data.PatientID) throw new Error('กรุณาเลือกผู้ป่วย');
  assertPatientAccess_(data.PatientID,ctx);
  const isNew = !data.AssessmentID;
  const previous = isNew ? null : rowsAsObjects_('ASSESSMENTS').find(r => String(r.AssessmentID) === String(data.AssessmentID));
  const calculated = calculateAssessment_(data.AssessmentType, data.AssessmentData);
  if (calculated) Object.keys(calculated).forEach(k => data[k] = calculated[k]);
  else if (data.ADLScore !== '' && data.ADLScore != null) {
    const score = Number(data.ADLScore);
    data.DependencyLevel = score <= 4 ? 'ติดเตียง' : score <= 11 ? 'ติดบ้าน' : 'ติดสังคม';
  }
  const result = saveRecord_('ASSESSMENTS','AssessmentID','AS',data,ctx);
  if (data.PatientID && data.ADLScore !== '' && data.ADLScore != null) {
    updateById_('PATIENTS','PatientID',data.PatientID,{ADLScore:data.ADLScore,DependencyLevel:data.DependencyLevel,UpdatedAt:new Date(),UpdatedBy:ctx.username});
  }
  const newlyDetectedSuicideRisk = data.SuicideRisk === 'มีความเสี่ยง' && (!previous || previous.SuicideRisk !== 'มีความเสี่ยง');
  if (data.PatientID && newlyDetectedSuicideRisk) {
    const alert = {
      AlertID:makeId_('AL'),PatientID:data.PatientID,AlertDate:new Date(),Level:'แดง',Category:'สุขภาพจิต',
      Message:'2Q Plus พบความเสี่ยงต่อการฆ่าตัวตาย ต้องประเมิน 8Q และติดต่อ CM/หน่วยบริการทันที',
      Status:'เปิด',AssignedTo:'CM',CreatedBy:ctx.username,CreatedAt:new Date()
    };
    appendObject_('ALERTS',alert);
    result.alert = alert;
  }
  return result;
}

function calculateAssessment_(type, rawData) {
  if (!rawData) return null; // รองรับข้อมูลแบบเดิมที่ยังไม่มีคำตอบรายข้อ
  let detail;
  try { detail = typeof rawData === 'string' ? JSON.parse(rawData) : rawData; }
  catch (e) { throw new Error('รายละเอียดคำตอบแบบประเมินไม่ถูกต้อง'); }
  if (!detail || typeof detail !== 'object') throw new Error('รายละเอียดคำตอบแบบประเมินไม่ถูกต้อง');

  const key = String(type || '').trim().toUpperCase();
  if (key === 'ADL' || key.indexOf('ADL') === 0) {
    const answers = validateScores_(detail.answers, [[0,1,2],[0,1],[0,1,2,3],[0,1,2],[0,1,2,3],[0,1,2],[0,1,2],[0,1],[0,1,2],[0,1,2]], 'ADL');
    const score = answers.reduce((s,n) => s+n, 0);
    const result = score <= 4 ? 'ภาวะพึ่งพาโดยสมบูรณ์' : score <= 8 ? 'ภาวะพึ่งพารุนแรง' : score <= 11 ? 'ภาวะพึ่งพาปานกลาง' : 'ไม่เป็นการพึ่งพา';
    const dependency = score <= 4 ? 'ติดเตียง' : score <= 11 ? 'ติดบ้าน' : 'ติดสังคม';
    const advice = score <= 4
      ? 'ช่วยเหลือกิจวัตรประจำวันทุกด้าน เฝ้าระวังแผลกดทับ การสำลัก และทบทวนแผนดูแลร่วมกับ CM'
      : score <= 8
        ? 'จัดผู้ดูแลช่วยเหลืออย่างใกล้ชิด เน้นกิจกรรมที่ยังทำได้ และประเมินซ้ำตามแผนการดูแล'
        : score <= 11
          ? 'ช่วยเฉพาะกิจกรรมที่ยังติดขัด ส่งเสริมการเคลื่อนไหวอย่างปลอดภัย และประเมินซ้ำตามแผน'
          : 'ส่งเสริมให้ทำกิจวัตรด้วยตนเอง รักษาความแข็งแรง และติดตามการเปลี่ยนแปลงเป็นระยะ';
    return {ADLScore:score,DependencyLevel:dependency,Summary:'ADL '+score+'/20 คะแนน • '+result,Advice:advice,AssessmentData:JSON.stringify({answers:answers})};
  }

  if (key === 'THAI-FRAT' || key === 'THAI_FRAT') {
    const answers = validateScores_(detail.answers, [[0,1],[0,1],[0,2],[0,1,2,3,4],[0,5],[0,1]], 'Thai-FRAT');
    const score = answers.reduce((s,n) => s+n, 0);
    const risk = score <= 3 ? 'ต่ำ' : 'สูง';
    const advice = risk === 'ต่ำ'
      ? 'ออกกำลังกายเพิ่มความแข็งแรงและการทรงตัวอย่างสม่ำเสมอ พร้อมทบทวนความปลอดภัยในบ้าน'
      : 'แจ้ง CM หรือบุคลากรสาธารณสุขเพื่อค้นหาสาเหตุ ทบทวนยา และเร่งแก้จุดเสี่ยงในบ้าน';
    return {FallRiskScore:score,FallRisk:risk,EnvironmentRisk:answers[5] ? 'สูง' : 'ต่ำ',Summary:'Thai-FRAT '+score+' คะแนน • ความเสี่ยงหกล้ม'+risk,Advice:advice,AssessmentData:JSON.stringify({answers:answers})};
  }

  if (key === 'MNA-SF' || key === 'MNA_SF') {
    const answers = validateScores_(detail.answers, [[0,1,2],[0,1,2,3],[0,1,2],[0,2],[0,1,2]], 'MNA-SF');
    const measurement = detail.measurement || {};
    let fScore = 0, normalizedMeasurement;
    if (measurement.method === 'bmi') {
      const weight = Number(measurement.weight), height = Number(measurement.height);
      if (!(weight > 0) || !(height > 0)) throw new Error('กรุณาระบุน้ำหนักและส่วนสูงสำหรับ MNA-SF');
      const bmi = weight / Math.pow(height / 100, 2);
      fScore = bmi < 19 ? 0 : bmi < 21 ? 1 : bmi < 23 ? 2 : 3;
      normalizedMeasurement = {method:'bmi',weight:weight,height:height,bmi:Math.round(bmi*10)/10,score:fScore};
    } else if (measurement.method === 'calf') {
      const calf = Number(measurement.calf);
      if (!(calf > 0)) throw new Error('กรุณาระบุเส้นรอบวงน่องสำหรับ MNA-SF');
      fScore = calf < 31 ? 0 : 3;
      normalizedMeasurement = {method:'calf',calf:calf,score:fScore};
    } else throw new Error('กรุณาเลือกวิธีประเมิน BMI หรือเส้นรอบวงน่อง');
    const score = answers.reduce((s,n) => s+n, 0) + fScore;
    const result = score >= 12 ? 'ภาวะโภชนาการปกติ' : score >= 8 ? 'เสี่ยงต่อภาวะขาดสารอาหาร' : 'ขาดสารอาหาร';
    const risk = score >= 12 ? 'ปกติ' : score >= 8 ? 'เสี่ยงขาดสารอาหาร' : 'ขาดสารอาหาร';
    const advice = score >= 12
      ? 'คงรูปแบบอาหารที่เหมาะสม ชั่งน้ำหนักและติดตามการรับประทานอาหารเป็นระยะ'
      : 'แจ้ง CM และปรึกษาแพทย์หรือนักกำหนดอาหาร ควรประเมิน MNA ฉบับเต็มและวางแผนโภชนบำบัด';
    return {NutritionScore:score,NutritionRisk:risk,Summary:'MNA-SF '+score+'/14 คะแนน • '+result,Advice:advice,AssessmentData:JSON.stringify({answers:answers,measurement:normalizedMeasurement})};
  }

  if (key === '2Q PLUS' || key === '2Q_PLUS' || key === '2Q') {
    const answers = validateScores_(detail.answers, [[0,1],[0,1],[0,1]], '2Q Plus');
    const depressionRisk = answers[0] === 1 || answers[1] === 1;
    const suicideRisk = answers[2] === 1;
    const summary = suicideRisk
      ? '2Q Plus • พบความเสี่ยงต่อการฆ่าตัวตาย'
      : depressionRisk ? '2Q Plus • มีแนวโน้มหรือเสี่ยงต่อภาวะซึมเศร้า' : '2Q Plus • ไม่พบความเสี่ยงจากการคัดกรอง';
    const advice = suicideRisk
      ? 'เร่งด่วน: อยู่กับผู้ป่วย ไม่ปล่อยให้อยู่ลำพัง เก็บสิ่งที่อาจใช้ทำร้ายตนเอง ประเมิน 8Q และติดต่อ CM/หน่วยบริการทันที หากมีอันตรายฉุกเฉินโทร 1669; สายด่วนสุขภาพจิต 1323'
      : depressionRisk
        ? 'ประเมินต่อด้วยแบบ 9Q แจ้ง CM เพื่อติดตาม และแนะนำสายด่วนสุขภาพจิต 1323 หากต้องการคำปรึกษา'
        : 'ให้กำลังใจ ดูแลสุขภาพจิตทั่วไป และคัดกรองซ้ำเมื่ออาการหรือสถานการณ์เปลี่ยนแปลง';
    return {Depression:depressionRisk?'เสี่ยง':'ไม่พบ',SuicideRisk:suicideRisk?'มีความเสี่ยง':'ไม่พบ',Summary:summary,Advice:advice,AssessmentData:JSON.stringify({answers:answers})};
  }
  throw new Error('ไม่รู้จักประเภทแบบประเมิน');
}

function validateScores_(answers, allowed, label) {
  if (!Array.isArray(answers) || answers.length !== allowed.length) throw new Error('กรุณาตอบแบบประเมิน '+label+' ให้ครบทุกข้อ');
  return answers.map((value,i) => {
    const score = Number(value);
    if (!allowed[i].includes(score)) throw new Error('คำตอบข้อที่ '+(i+1)+' ของ '+label+' ไม่ถูกต้อง');
    return score;
  });
}

function detectAlerts_(visit, ctx) {
  const alerts = [];
  const tests = [
    [Number(visit.SpO2) > 0 && Number(visit.SpO2) < 94, 'แดง', 'SpO2', 'ค่าออกซิเจนในเลือดต่ำกว่า 94%'],
    [Number(visit.Systolic) >= 180 || (Number(visit.Systolic) > 0 && Number(visit.Systolic) < 90), 'แดง', 'ความดันโลหิต', 'ค่าความดันตัวบนอยู่ในระดับเร่งด่วน'],
    [Number(visit.DTX) > 400 || (Number(visit.DTX) > 0 && Number(visit.DTX) < 70), 'แดง', 'น้ำตาล', 'ระดับน้ำตาลอยู่ในช่วงอันตราย'],
    [Number(visit.Temperature) >= 38.5, 'ส้ม', 'ไข้', 'อุณหภูมิร่างกายสูงตั้งแต่ 38.5°C']
  ];
  tests.filter(t => t[0]).forEach(t => {
    const obj = {AlertID:makeId_('AL'),PatientID:visit.PatientID,AlertDate:new Date(),Level:t[1],Category:t[2],Message:t[3],Status:'เปิด',AssignedTo:'CM',CreatedBy:ctx.username,CreatedAt:new Date()};
    appendObject_('ALERTS',obj); alerts.push(obj);
  });
  return alerts;
}

function archiveRecord_(sheetName, idField, id, ctx) {
  requireRole_(ctx,['admin','manager','cm']);
  const ok = updateById_(sheetName,idField,id,{Status:'archived',UpdatedAt:new Date(),UpdatedBy:ctx.username});
  if (!ok) throw new Error('ไม่พบข้อมูล');
  audit_(ctx,'ARCHIVE',sheetName,id,'เก็บข้อมูลเข้าคลัง');
  return {ok:true};
}

function listUsers_(ctx) {
  const rows = rowsAsObjects_('USERS').map(u => ({UserID:u.UserID,Username:u.Username,DisplayName:u.DisplayName,Role:u.Role,Status:u.Status,MustChangePassword:isTruthy_(u.MustChangePassword),LastLogin:u.LastLogin,CreatedAt:u.CreatedAt}));
  audit_(ctx,'READ_LIST','USERS','','อ่านรายชื่อผู้ใช้');
  return {ok:true,data:rows};
}

function saveUser_(data, ctx) {
  if (!data.Username) throw new Error('กรุณาระบุชื่อผู้ใช้');
  const existing = rowsAsObjects_('USERS').find(u => String(u.Username).toLowerCase() === String(data.Username).toLowerCase() && u.UserID !== data.UserID);
  if (existing) throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้ว');
  const now = new Date();
  if (data.UserID) {
    const patch = {Username:String(data.Username).toLowerCase(),DisplayName:data.DisplayName,Role:data.Role,Status:data.Status || 'active',UpdatedAt:now};
    if (data.Password) { patch.Salt = makeSalt_(); patch.PasswordHash = hashPassword_(data.Password, patch.Salt); patch.MustChangePassword = true; }
    updateById_('USERS','UserID',data.UserID,patch);
    audit_(ctx,'UPDATE','USERS',data.UserID,'แก้ไขผู้ใช้');
    return {ok:true,id:data.UserID};
  }
  if (!data.Password) throw new Error('กรุณาระบุรหัสผ่าน');
  const salt = makeSalt_();
  const row = {UserID:makeId_('US'),Username:String(data.Username).toLowerCase(),PasswordHash:hashPassword_(data.Password,salt),Salt:salt,DisplayName:data.DisplayName||data.Username,Role:data.Role||'cg',Status:data.Status||'active',CreatedAt:now,UpdatedAt:now,MustChangePassword:true};
  appendObject_('USERS',row); audit_(ctx,'CREATE','USERS',row.UserID,'สร้างผู้ใช้');
  return {ok:true,id:row.UserID};
}

function closeSessionsForUser_(userId) {
  const sheet = getSheet_('SESSIONS');
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 0;
  const values = sheet.getRange(1,1,lastRow,lastCol).getValues(), headers = values[0];
  const tokenCol = headers.indexOf('Token'), userCol = headers.indexOf('UserID'), statusCol = headers.indexOf('Status'), lastSeenCol = headers.indexOf('LastSeen');
  if (tokenCol < 0 || userCol < 0 || statusCol < 0) return 0;
  const tokens = [];
  let closed = 0;
  for (let r=1;r<values.length;r++) {
    if (String(values[r][userCol]) !== String(userId)) continue;
    if (values[r][tokenCol]) tokens.push(authCacheKey_(String(values[r][tokenCol])));
    if (String(values[r][statusCol]).toLowerCase() === 'active') {
      values[r][statusCol] = 'closed';
      if (lastSeenCol >= 0) values[r][lastSeenCol] = new Date();
      closed++;
    }
  }
  if (closed) sheet.getRange(2,1,values.length-1,lastCol).setValues(values.slice(1));
  if (tokens.length) CacheService.getScriptCache().removeAll(tokens);
  return closed;
}

function adminResetUserPassword_(data, ctx) {
  if (!data || !data.userId) throw new Error('กรุณาเลือกบัญชีผู้ใช้');
  const password = String(data.newPassword || '');
  if (password.length < 4) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 4 ตัวอักษร');
  const user = rowsAsObjects_('USERS').find(item => String(item.UserID) === String(data.userId));
  if (!user) throw new Error('ไม่พบบัญชีผู้ใช้');
  const salt = makeSalt_();
  if (!updateById_('USERS','UserID',user.UserID,{Salt:salt,PasswordHash:hashPassword_(password,salt),MustChangePassword:true,UpdatedAt:new Date()})) throw new Error('ไม่สามารถตั้งรหัสผ่านใหม่ได้');
  const sessionsClosed = closeSessionsForUser_(user.UserID);
  audit_(ctx,'ADMIN_RESET_PASSWORD','USERS',user.UserID,'ผู้ดูแลระบบตั้งรหัสผ่านใหม่ให้ '+user.Username);
  return {
    ok:true,
    message:'ตั้งรหัสผ่านใหม่ให้ '+(user.DisplayName || user.Username)+' เรียบร้อยแล้ว',
    userId:user.UserID,
    currentUser:String(user.UserID) === String(ctx.userId),
    sessionsClosed:sessionsClosed,
    mustChangePassword:true
  };
}

function changePassword_(payload, ctx) {
  const user = rowsAsObjects_('USERS').find(u => u.UserID === ctx.userId);
  if (!user || hashPassword_(String(payload.currentPassword||''),String(user.Salt)) !== String(user.PasswordHash)) throw new Error('รหัสผ่านปัจจุบันไม่ถูกต้อง');
  if (String(payload.newPassword||'').length < 4) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 4 ตัวอักษร');
  const salt = makeSalt_();
  updateById_('USERS','UserID',ctx.userId,{Salt:salt,PasswordHash:hashPassword_(payload.newPassword,salt),MustChangePassword:false,UpdatedAt:new Date()});
  CacheService.getScriptCache().remove(authCacheKey_(ctx.token));
  audit_(ctx,'CHANGE_PASSWORD','USERS',ctx.userId,'เปลี่ยนรหัสผ่าน');
  return {ok:true,message:'เปลี่ยนรหัสผ่านเรียบร้อยแล้ว',mustChangePassword:false};
}

function resetAdminPassword() {
  setupSystem();
  const admin = rowsAsObjects_('USERS').find(u => String(u.Username).toLowerCase() === 'admin');
  const salt = makeSalt_();
  updateById_('USERS','UserID',admin.UserID,{Salt:salt,PasswordHash:hashPassword_(APP.DEFAULT_PASSWORD,salt),Status:'active',MustChangePassword:true,UpdatedAt:new Date()});
  SpreadsheetApp.getUi().alert('รีเซ็ตรหัสผ่าน admin เป็น 1234 แล้ว กรุณาเปลี่ยนทันทีหลังเข้าสู่ระบบ');
}

function saveSettings_(items, ctx) {
  (items || []).forEach(item => {
    const key = String(item.Key || '').trim(); if (!key) return;
    if (!updateById_('SETTINGS','Key',key,{Value:item.Value,UpdatedAt:new Date(),UpdatedBy:ctx.username})) appendObject_('SETTINGS',{Key:key,Value:item.Value,UpdatedAt:new Date(),UpdatedBy:ctx.username});
  });
  audit_(ctx,'UPDATE','SETTINGS','','แก้ไขการตั้งค่า');
  return {ok:true};
}

function seedSettings_() {
  const defaults = Object.assign({APP_NAME:'ระบบเยี่ยมบ้านผู้ป่วย LTC',ORGANIZATION:'โรงพยาบาลส่งเสริมสุขภาพตำบลบ้านชมภูพานเหนือ',OWNER:'นางศันสนีย์ มีพรม พยาบาลวิชาชีพชำนาญการ',FISCAL_YEAR:'2570'},CONTRACT_MASTER_DEFAULTS);
  const existing = rowsAsObjects_('SETTINGS').map(r => r.Key);
  Object.keys(defaults).filter(k => !existing.includes(k)).forEach(k => appendObject_('SETTINGS',{Key:k,Value:defaults[k],UpdatedAt:new Date(),UpdatedBy:'system'}));
}

function seedAdmin_() {
  if (rowsAsObjects_('USERS').some(u => String(u.Username).toLowerCase() === APP.DEFAULT_ADMIN)) return;
  const salt = makeSalt_();
  appendObject_('USERS',{UserID:makeId_('US'),Username:APP.DEFAULT_ADMIN,PasswordHash:hashPassword_(APP.DEFAULT_PASSWORD,salt),Salt:salt,DisplayName:'ผู้ดูแลระบบ',Role:'admin',Status:'active',CreatedAt:new Date(),UpdatedAt:new Date()});
}

function seedCgUsers_(markExistingForFirstLogin) {
  const existing = rowsAsObjects_('USERS');
  const usersByUsername = existing.reduce((map,user) => {
    map[String(user.Username || '').trim().toLowerCase()] = user;
    return map;
  },{});
  DEFAULT_CG_USERS.forEach(displayName => {
    const username = String(displayName).trim().toLowerCase();
    const current = usersByUsername[username];
    if (current) {
      if (markExistingForFirstLogin) updateById_('USERS','UserID',current.UserID,{Role:'cg',Status:'active',MustChangePassword:true,UpdatedAt:new Date()});
      return;
    }
    const salt = makeSalt_();
    appendObject_('USERS',{
      UserID:makeId_('US'),Username:username,PasswordHash:hashPassword_(APP.DEFAULT_PASSWORD,salt),Salt:salt,
      DisplayName:displayName,Role:'cg',Status:'active',CreatedAt:new Date(),UpdatedAt:new Date(),MustChangePassword:true
    });
    usersByUsername[username] = {Username:username};
  });
}

function seedCodebook_() {
  if (rowsAsObjects_('CODEBOOK').length) return;
  const entries = [
    ['risk','green','เขียว','ปกติ',true],['risk','yellow','เหลือง','เฝ้าระวัง',true],['risk','orange','ส้ม','ติดตามเร่งด่วน',true],['risk','red','แดง','ส่งต่อ/ตอบสนองทันที',true],
    ['status','active','กำลังดูแล','ผู้ป่วยอยู่ในระบบ',true],['status','referred','ส่งต่อ','ส่งต่อหน่วยบริการอื่น',true],['status','deceased','เสียชีวิต','สิ้นสุดการดูแล',true],
    ['role','admin','ผู้ดูแลระบบ','สิทธิ์ทั้งหมด',true],['role','manager','ผู้บริหาร','Dashboard และรายงาน',true],['role','cm','Care Manager','จัดการผู้ป่วยและ Care Plan',true],['role','cg','Care Giver','บันทึกการเยี่ยม',true]
  ];
  entries.forEach(r => appendObject_('CODEBOOK',{Category:r[0],Code:r[1],ThaiLabel:r[2],Description:r[3],Active:r[4]}));
}

function cleanupSessions() {
  const sheet = getSheet_('SESSIONS');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {ok:true,removed:0};
  const now = Date.now();
  const active = values.slice(1).filter(row => {
    const expiresAt = new Date(row[5]).getTime();
    return Number.isFinite(expiresAt) && expiresAt > now && String(row[7]).toLowerCase() === 'active';
  });
  const oldCount = values.length - 1;
  const lastCol = values[0].length;
  if (oldCount === active.length) return {ok:true,removed:0};
  if (active.length) sheet.getRange(2,1,active.length,lastCol).setValues(active);
  if (oldCount > active.length) sheet.getRange(active.length+2,1,oldCount-active.length,lastCol).clearContent();
  return {ok:true,removed:oldCount-active.length};
}

function touchSession_(token) {
  const cache = CacheService.getScriptCache();
  const key = touchCacheKey_(token);
  if (cache.get(key)) return;
  updateById_('SESSIONS','Token',token,{LastSeen:new Date()});
  cache.put(key, '1', PERFORMANCE.SESSION_TOUCH_TTL);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const current = sheet.getLastColumn() ? sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),headers.length)).getValues()[0] : [];
  const merged = current.filter(String);
  headers.forEach(h => { if (!merged.includes(h)) merged.push(h); });
  sheet.getRange(1,1,1,merged.length).setValues([merged]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#159947').setHorizontalAlignment('center');
  sheet.setFrozenRows(1); sheet.getDataRange().setVerticalAlignment('middle');
  return sheet;
}

function getSheet_(name) { const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); if(!s) throw new Error('ไม่พบชีต '+name); return s; }

function rowsAsObjects_(name) {
  const cached = getCachedRows_(name);
  if (cached) return cached;
  const sheet=getSheet_(name), lastRow=sheet.getLastRow(), lastCol=sheet.getLastColumn();
  if(lastRow<2 || lastCol<1) { putCachedRows_(name,[]); return []; }
  const values=sheet.getRange(1,1,lastRow,lastCol).getValues();
  if(values.length<2) { putCachedRows_(name,[]); return []; }
  const headers=values[0].map(String);
  const rows=values.slice(1).filter(r=>r.some(v=>v!=='' && v!=null)).map(r=>{const o={};headers.forEach((h,i)=>o[h]=serialize_(r[i]));return o;});
  putCachedRows_(name, rows);
  return rows;
}

function appendObject_(name,obj) {
  const sheet=getSheet_(name), headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(h => obj[h] === undefined ? '' : obj[h]));
  clearDataCache_(name);
}

function updateById_(name,idField,id,patch) {
  const sheet=getSheet_(name), lastRow=sheet.getLastRow(), lastCol=sheet.getLastColumn();
  if(lastRow<2 || lastCol<1) return false;
  const values=sheet.getRange(1,1,lastRow,lastCol).getValues(), headers=values[0], idCol=headers.indexOf(idField);
  if(idCol<0) throw new Error('ไม่พบคอลัมน์ '+idField);
  for(let r=1;r<values.length;r++) if(String(values[r][idCol])===String(id)) {
    Object.keys(patch).forEach(k=>{const c=headers.indexOf(k);if(c>=0) values[r][c]=patch[k]===undefined?'':patch[k];});
    sheet.getRange(r+1,1,1,lastCol).setValues([values[r]]); clearDataCache_(name); return true;
  }
  return false;
}

function readyCacheKey_() { return 'LTC_READY_' + APP.VERSION; }
function dataCacheKey_(name) { return 'LTC_DATA_' + APP.VERSION + '_' + name; }
function dataCacheChunkKey_(name,index) { return dataCacheKey_(name) + '_PART_' + index; }
function dashboardCacheKey_(isPublic) { return 'LTC_DASH_' + APP.VERSION + '_' + (isPublic ? 'PUBLIC' : 'PRIVATE'); }
function authCacheKey_(token) { return 'LTC_AUTH_' + token; }
function touchCacheKey_(token) { return 'LTC_TOUCH_' + token; }
function isCacheableSheet_(name) { return PERFORMANCE.CACHEABLE_SHEETS.includes(name); }
function utf8ByteLength_(value) {
  let bytes = 0;
  for (let i=0;i<value.length;i++) {
    const code = value.charCodeAt(i);
    if (code < 0x80) bytes += 1;
    else if (code < 0x800) bytes += 2;
    else if (code >= 0xD800 && code <= 0xDBFF && i+1 < value.length) { bytes += 4; i++; }
    else bytes += 3;
  }
  return bytes;
}
function getCachedRows_(name) {
  if (!isCacheableSheet_(name)) return null;
  try {
    const cache = CacheService.getScriptCache();
    const value = cache.get(dataCacheKey_(name));
    if (!value) return null;
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) return parsed;
    if (!parsed || parsed.version !== 1 || !(parsed.chunks >= 1)) return null;
    const keys = Array.from({length:parsed.chunks},(_,index) => dataCacheChunkKey_(name,index));
    const parts = cache.getAll(keys);
    if (keys.some(key => !parts[key])) return null;
    return keys.reduce((rows,key) => rows.concat(JSON.parse(parts[key])),[]);
  } catch (e) { return null; }
}
function putCachedRows_(name, rows) {
  if (!isCacheableSheet_(name)) return;
  try {
    const chunks = [];
    let batch = [], batchBytes = 2;
    for (let i=0;i<rows.length;i++) {
      const encoded = JSON.stringify(rows[i]);
      const encodedBytes = utf8ByteLength_(encoded);
      if (encodedBytes + 2 > PERFORMANCE.CACHE_CHUNK_BYTES) return;
      const extraBytes = encodedBytes + (batch.length ? 1 : 0);
      if (batch.length && batchBytes + extraBytes > PERFORMANCE.CACHE_CHUNK_BYTES) {
        chunks.push('['+batch.join(',')+']');
        batch = [];
        batchBytes = 2;
      }
      batch.push(encoded);
      batchBytes += encodedBytes + (batch.length > 1 ? 1 : 0);
    }
    chunks.push('['+batch.join(',')+']');
    if (chunks.length > 50) return;
    const values = {};
    values[dataCacheKey_(name)] = JSON.stringify({version:1,chunks:chunks.length});
    chunks.forEach((value,index) => values[dataCacheChunkKey_(name,index)] = value);
    CacheService.getScriptCache().putAll(values,PERFORMANCE.DATA_TTL);
  } catch (e) {}
}
function clearDashboardCaches_() {
  CacheService.getScriptCache().removeAll([dashboardCacheKey_(true),dashboardCacheKey_(false)]);
}
function clearDataCache_(name) {
  if (isCacheableSheet_(name)) CacheService.getScriptCache().remove(dataCacheKey_(name));
  if (['PATIENTS','VISITS','ALERTS','STAFF','USERS','EQUIPMENT','BUDGETS'].includes(name)) clearDashboardCaches_();
}
function clearAllDataCaches_() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(PERFORMANCE.CACHEABLE_SHEETS.map(dataCacheKey_).concat([dashboardCacheKey_(true),dashboardCacheKey_(false)]));
}
function cacheAuth_(token, user, expiresAt) {
  try {
    CacheService.getScriptCache().put(authCacheKey_(token), JSON.stringify({user:user,expiresAt:formatDate_(expiresAt)}), PERFORMANCE.AUTH_TTL);
  } catch (e) {}
}
function getCachedAuth_(token) {
  try {
    const value = CacheService.getScriptCache().get(authCacheKey_(token));
    if (!value) return null;
    const cached = JSON.parse(value);
    if (new Date(cached.expiresAt).getTime() <= Date.now()) return null;
    return cached.user;
  } catch (e) { return null; }
}

function sanitizeObject_(obj,headers) { const out={}; headers.forEach(h=>{if(Object.prototype.hasOwnProperty.call(obj,h)) out[h]=sanitizeValue_(obj[h]);}); return out; }
function sanitizeValue_(v) { if(typeof v==='string' && /^[=+\-@]/.test(v)) return "'"+v; return v==null?'':v; }
function serialize_(v) { return Object.prototype.toString.call(v)==='[object Date]' ? formatDate_(v) : v; }
function formatDate_(d) { return Utilities.formatDate(new Date(d),APP.TZ,"yyyy-MM-dd'T'HH:mm:ssXXX"); }
function dateOnly_(v) { if(!v) return ''; if(typeof v==='string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.substring(0,10); try{return Utilities.formatDate(new Date(v),APP.TZ,'yyyy-MM-dd');}catch(e){return String(v).substring(0,10);} }
function now_() { return formatDate_(new Date()); }
function makeId_(prefix) { return prefix+'-'+Utilities.formatDate(new Date(),APP.TZ,'yyyyMMddHHmmss')+'-'+Math.floor(Math.random()*9000+1000); }
function makeSalt_() { return Utilities.getUuid().replace(/-/g,''); }
function hashPassword_(password,salt) { const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt+'|'+password,Utilities.Charset.UTF_8); return bytes.map(b=>('0'+((b<0?b+256:b).toString(16))).slice(-2)).join(''); }
function countBy_(rows,key) { return rows.reduce((a,r)=>{const k=String(r[key]||'ไม่ระบุ');a[k]=(a[k]||0)+1;return a;},{}); }
function getSetting_(key) { const r=rowsAsObjects_('SETTINGS').find(x=>x.Key===key); return r?r.Value:''; }
function clientIp_(e) { return e && e.parameter && e.parameter.ip ? String(e.parameter.ip).substring(0,64) : ''; }
function safeError_(e) { return String(e && e.message ? e.message : e).substring(0,500); }
function jsonResponse_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function audit_(ctx,action,entity,id,detail) { try{appendObject_('AUDIT_LOG',{LogID:makeId_('LG'),Timestamp:new Date(),Username:ctx&&ctx.username||'-',Role:ctx&&ctx.role||'-',Action:action,Entity:entity,EntityID:id||'',Detail:detail||'',IPAddress:ctx&&ctx.ip||''});}catch(e){} }
