// ==========================================
// GLOBAL CONFIGURATION & DEFAULT SETTINGS
// ==========================================
const GLOBAL_SETTINGS = {
mainSpreadsheetId: 'MOCK_SPREADSHEET_ID'
};

const DEFAULT_TEMPLATES = {
"Mock Template A": {
targetWorkflow: 'MOCK_WORKFLOW_A',
ccEmails: 'mock.user@example.com, mock.user@example.com, mock.user@example.com, mock.user@example.com, mock.user@example.com, mock.user@example.com',
emailSubject: '{{AGENT_ID}} || Mock Event || {{DATE}}',
driveFileId: 'MOCK_DRIVE_FILE_ID',
emailBody: 'Hi Team,<br><br>This is to inform you that <strong>{{AGENT_NAME}}</strong> will be in Mock Event mode on <strong>{{DATE}}</strong>.<br><br>The agent will recap all the updates that we had during these days.<br>You can access the personalized plan by clicking here: <a href="{{FILE_URL}}">Mock Event File</a>.<br><br>In the afternoon, the agent will have shadowing / reverse shadowing with a colleague.<br><br>Please let me know if you have any questions or doubts.<br><br>Regards,'
}
};

// ==========================================
// WEB APP INITIALIZATION
// ==========================================
function doGet(e) {
return HtmlService.createHtmlOutputFromFile('Index')
.setTitle('ExampleCorp - Email Automation')
.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// BATCH LOAD
// ==========================================
function getInitialData() {
try {
const workflows = getWorkflows();
const templates = getTemplates();
return { success: true, workflows: workflows, templates: templates };
} catch (e) {
return { success: false, error: e.message };
}
}

// ==========================================
// WORKFLOW (SHEETS) MANAGEMENT
// ==========================================
function getWorkflows() {
try {
const ss = SpreadsheetApp.openById(GLOBAL_SETTINGS.mainSpreadsheetId);
return ss.getSheets().map(sheet => sheet.getName());
} catch (e) {
throw new Error('Could not load workflows: ' + e.message);
}
}

function createWorkflow(workflowName) {
try {
if (!workflowName || !workflowName.trim()) throw new Error("Workflow name cannot be empty.");
const cleanName = workflowName.trim();
const ss = SpreadsheetApp.openById(GLOBAL_SETTINGS.mainSpreadsheetId);
if (ss.getSheetByName(cleanName)) {
return { success: false, message: `Workflow '${cleanName}' already exists.` };
}
const newSheet = ss.insertSheet(cleanName);
newSheet.getRange('A1:D2').setBackground('#141821').setFontColor('#d4af37').setFontWeight('bold');
newSheet.getRange('A2').setValue('Nº');
newSheet.getRange('B2').setValue('User ID');
newSheet.getRange('C2').setValue('Name');
newSheet.getRange('D2').setValue('Email');
newSheet.setFrozenRows(2);
return { success: true, message: `Workflow '${cleanName}' created successfully!`, workflows: getWorkflows() };
} catch (e) {
return { success: false, error: 'Could not create workflow: ' + e.message };
}
}

// ==========================================
// TEMPLATE MANAGEMENT
// ==========================================
function getTemplates() {
try {
const props = PropertiesService.getScriptProperties();
const savedTemplates = props.getProperty('EXAMPLECORP_EMAIL_TEMPLATES');
return savedTemplates ? JSON.parse(savedTemplates) : DEFAULT_TEMPLATES;
} catch (e) {
return DEFAULT_TEMPLATES;
}
}

function saveTemplate(templateName, templateData) {
try {
if (!templateName || templateName.trim() === '') throw new Error("Template name is required.");
const props = PropertiesService.getScriptProperties();
let templates = getTemplates();
templates[templateName.trim()] = templateData;
props.setProperty('EXAMPLECORP_EMAIL_TEMPLATES', JSON.stringify(templates));
return { success: true, message: `Template '${templateName}' saved!`, templates: templates };
} catch (e) {
return { success: false, error: 'Could not save template: ' + e.message };
}
}

function deleteTemplate(templateName) {
try {
const props = PropertiesService.getScriptProperties();
let templates = getTemplates();
if (templates[templateName]) {
delete templates[templateName];
if (Object.keys(templates).length === 0) templates = DEFAULT_TEMPLATES;
props.setProperty('EXAMPLECORP_EMAIL_TEMPLATES', JSON.stringify(templates));
return { success: true, message: `Template '${templateName}' deleted.`, templates: templates };
} else {
return { success: false, error: "Template not found." };
}
} catch (e) {
return { success: false, error: 'Could not delete template: ' + e.message };
}
}

// ==========================================
// AGENT LOGIC (DYNAMIC WORKFLOWS)
// ==========================================
function getAgentsByWorkflow(workflowName) {
try {
if(!workflowName) return [];
const cacheKey = `EXAMPLECORP_AGENTS_${workflowName.replace(/\s+/g, '_')}`;
const cache = CacheService.getScriptCache();
const cachedAgents = cache.get(cacheKey);
if (cachedAgents) return JSON.parse(cachedAgents);

const sheet = SpreadsheetApp.openById(GLOBAL_SETTINGS.mainSpreadsheetId).getSheetByName(workflowName);
if (!sheet) return [];
const lastRow = sheet.getLastRow();
if (lastRow < 3) return [];
const values = sheet.getRange('B3:D' + lastRow).getValues();
const agents = values.map(row => {
const [id, name, email] = row;
if (id && String(id).trim() && email && String(email).trim()) {
return { id: String(id).trim(), name: name ? String(name).trim() : String(id).trim(), email: String(email).trim() };
}
return null;
}).filter(agent => agent);

cache.put(cacheKey, JSON.stringify(agents), 3600);
return agents;
} catch (e) {
throw new Error(`Could not load agents for ${workflowName}: ` + e.message);
}
}

function addAgentsBulk(agentsString, workflowName) {
if (!workflowName) throw new Error("Please select a workflow first.");
if (!agentsString || !agentsString.trim()) throw new Error("Please provide valid User IDs.");
const rawUserIds = agentsString.split(',').map(s => s.trim()).filter(s => s !== '');
if (rawUserIds.length === 0) throw new Error("No valid User IDs found.");

try {
const ss = SpreadsheetApp.openById(GLOBAL_SETTINGS.mainSpreadsheetId);
const sheet = ss.getSheetByName(workflowName);
if (!sheet) throw new Error(`Workflow '${workflowName}' not found.`);

const columnBValues = sheet.getRange('B1:B').getValues().flat().map(v => v.toString().trim());
let addedCount = 0;
let skipped = [];

rawUserIds.forEach(userId => {
if (columnBValues.includes(userId)) {
skipped.push(userId);
} else {
let targetRow = sheet.getLastRow() + 1;
if (targetRow < 3) targetRow = 3;
sheet.getRange(targetRow, 2).setValue(userId);
sheet.getRange(targetRow, 3).setValue(userId);
sheet.getRange(targetRow, 4).setValue(`${userId}@example.com`);
columnBValues.push(userId);
addedCount++;
}
});
const cacheKey = `EXAMPLECORP_AGENTS_${workflowName.replace(/\s+/g, '_')}`;
CacheService.getScriptCache().remove(cacheKey);
let msg = `Successfully added ${addedCount} agent(s).`;
if(skipped.length > 0) msg += ` Skipped ${skipped.length} existing: ${skipped.slice(0,3).join(', ')}${skipped.length>3?'...':''}`;
return { success: true, message: msg };
} catch (e) {
throw new Error(`Insertion error: ${e.message}`);
}
}

function deleteAgentsBatch(agentUserIds, workflowName) {
if (!workflowName) throw new Error("Workflow not selected.");
if (!agentUserIds || agentUserIds.length === 0) throw new Error("No agents provided for deletion.");
try {
const ss = SpreadsheetApp.openById(GLOBAL_SETTINGS.mainSpreadsheetId);
const sheet = ss.getSheetByName(workflowName);
if (!sheet) throw new Error(`Workflow '${workflowName}' not found.`);
let deletedCount = 0;
const columnBValues = sheet.getRange('B1:B' + sheet.getLastRow()).getValues().flat().map(v => v.toString().trim());
for (let i = columnBValues.length - 1; i >= 0; i--) {
if (agentUserIds.includes(columnBValues[i])) {
sheet.deleteRow(i + 1);
deletedCount++;
}
}
if (deletedCount > 0) {
const cacheKey = `EXAMPLECORP_AGENTS_${workflowName.replace(/\s+/g, '_')}`;
CacheService.getScriptCache().remove(cacheKey);
return { success: true, message: `Successfully removed ${deletedCount} agent(s) from ${workflowName}.` };
} else {
return { success: false, message: `Agents not found in ${workflowName}.` };
}
} catch (e) {
throw new Error(`Removal error: ${e.message}`);
}
}

// ==========================================
// EMAIL BATCH GENERATION LOGIC (UPDATED FOR OPTIONAL FILES)
// ==========================================
function createAndPrepareEmailsBatch(workflowName, agentIds, dateValue, customEmailData) {
try {
const agents = getAgentsByWorkflow(workflowName);
const fileIdToCopy = customEmailData.driveFileId;
// FLAG PER CAPIRE SE DEVO CLONARE UN FILE O NO
const requiresFile = (fileIdToCopy && fileIdToCopy.trim() !== '');

const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const date = new Date(dateValue);
const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
const managerEmails = customEmailData.ccEmails.split(',').map(e => e.trim());
const validManagerEmails = managerEmails.filter(email => emailRegex.test(email));
let templateFile = null;
// PRENDO IL FILE SOLO SE È STATO INSERITO UN ID
if (requiresFile) {
try {
templateFile = DriveApp.getFileById(fileIdToCopy.trim());
} catch(err) {
throw new Error("Could not find the Drive File. Ensure the ID is correct and you have access.");
}
}

let createdCount = 0;
let firstDraftUrl = "";

agentIds.forEach(agentId => {
let selectedAgent = agents.find(agent => agent.id === agentId);
if (!selectedAgent) return;
let newFileUrl = ""; // Se non c'è file, l'URL resta vuoto
// CLONO IL FILE SOLO SE RICHIESTO
if (requiresFile && templateFile) {
const newFileName = `Mock Event - ${selectedAgent.name} - ${formattedDate}`;
const newFile = templateFile.makeCopy(newFileName);
const newFileId = newFile.getId();
newFileUrl = newFile.getUrl();
SpreadsheetApp.openById(newFileId).getSheets()[0].getRange('A2').setValue(`${selectedAgent.name} ${formattedDate}`);
const allEditorEmails = [selectedAgent.email, ...validManagerEmails];
allEditorEmails.forEach(email => {
try { Drive.Permissions.insert({ 'role': 'writer', 'type': 'user', 'value': email }, newFileId, { 'sendNotificationEmails': false }); } catch (err) {}
});
}
const parseTemplate = (text) => {
return text.replace(/{{AGENT_ID}}/g, selectedAgent.id)
.replace(/{{AGENT_NAME}}/g, selectedAgent.name)
.replace(/{{DATE}}/g, formattedDate)
.replace(/{{FILE_URL}}/g, newFileUrl); // Se non c'è file, toglie il tag senza rompere l'HTML
};

const finalSubject = parseTemplate(customEmailData.emailSubject);
const finalBody = parseTemplate(customEmailData.emailBody);
const draft = GmailApp.createDraft(
[selectedAgent.email, ...validManagerEmails].join(','),
finalSubject,
'',
{ htmlBody: finalBody }
);
if (createdCount === 0) {
firstDraftUrl = `https://example.com/mail/u/0/#drafts?compose=${draft.getId()}`;
}
createdCount++;
});

const returnUrl = createdCount === 1 ? firstDraftUrl : 'https://example.com/mail/u/0/#drafts';
return { success: true, count: createdCount, url: returnUrl };
} catch (e) {
throw new Error(`Fatal error: ${e.message}`);
}
}



