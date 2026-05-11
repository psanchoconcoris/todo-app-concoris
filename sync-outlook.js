#!/usr/bin/env node
/**
 * SYNC-OUTLOOK.JS
 * Script que sincroniza emails de Outlook con la aplicación de tareas
 */

const https = require('https');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// CONFIGURACIÓN
const OUTLOOK_EMAIL = 'psancho@concoris.es';

const TEAM_MAPPING = {
  'admanistracion@concoris.es': 'Jordi',
  'usp@urbanservicepoint.es': 'Laura',
  'lszalacian@concoris.es': 'Luis Sancho Zalacain',
  'pthielen@corporalfisiocenter.com': 'Pablo Thielen',
  'acastro@adiari.es': 'Anna Castro',
  'rsancho@adiari.es': 'Rosario Sancho'
};

const COMPANIES = [
  'USP Urban Infrastructure',
  'USP Urban Logistics',
  'Corporal',
  'ADIARI',
  'Kerner Hill',
  'Taulies Capital',
  'Soportes',
  'Harvard',
  'GEL',
  'KH'
];

// FUNCIONES UTILITARIAS
function httpRequest(options, data = null) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        try {
          const parsed = body ? JSON.parse(body) : {};
          if (res.statusCode >= 400) {
            reject(new Error(`HTTP ${res.statusCode}: ${body}`));
          } else {
            resolve(parsed);
          }
        } catch (e) {
          reject(e);
        }
      });
    });
    req.on('error', reject);
    if (data) req.write(JSON.stringify(data));
    req.end();
  });
}

async function getAccessToken() {
  const clientId = process.env.OUTLOOK_CLIENT_ID;
  const clientSecret = process.env.OUTLOOK_CLIENT_SECRET;
  const tenantId = process.env.OUTLOOK_TENANT_ID || 'common';

  const options = {
    hostname: 'login.microsoftonline.com',
    path: `/${tenantId}/oauth2/v2.0/token`,
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  };

  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (res.statusCode >= 400) {
            reject(new Error(`Auth failed: ${data}`));
          } else {
            resolve(parsed.access_token);
          }
        } catch (e) {
          reject(e);
        }
      });
    });
    req.on('error', reject);
    req.write(new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    }).toString());
    req.end();
  });
}

async function getEmailsFromOutlook(accessToken) {
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const startDate = yesterday.toISOString().split('T')[0];

  const options = {
    hostname: 'graph.microsoft.com',
    path: `/v1.0/me/mailFolders/inbox/messages?$filter=receivedDateTime ge ${startDate}T00:00:00Z&$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,internetMessageId&$top=50`,
    method: 'GET',
    headers: { 'Authorization': `Bearer ${accessToken}` }
  };

  return httpRequest(options);
}

function parseSubject(subject) {
  const companies = [];
  let taskTitle = subject;
  const match = subject.match(/^([^-/]+(?:\s*\/\s*[^-/]+)*?)\s*-\s*(.+)$/);

  if (match) {
    const companyPart = match[1];
    taskTitle = match[2];
    const companyNames = companyPart.split('/').map(c => c.trim());

    for (let name of companyNames) {
      const found = COMPANIES.find(comp =>
        comp.toUpperCase().includes(name.toUpperCase()) ||
        name.toUpperCase().includes(comp.toUpperCase())
      );
      if (found && !companies.includes(found)) {
        companies.push(found);
      }
    }
  }

  return {
    companies: companies.length > 0 ? companies : [''],
    taskTitle: taskTitle.trim()
  };
}

function identifyPerson(email) {
  return TEAM_MAPPING[email.toLowerCase()] || null;
}

function generateTaskHash(email) {
  return crypto.createHash('md5').update(email.internetMessageId || email.id).digest('hex');
}

function loadProcessedEmails() {
  const filepath = path.join(__dirname, 'processed-emails.json');
  if (fs.existsSync(filepath)) {
    return JSON.parse(fs.readFileSync(filepath, 'utf8'));
  }
  return {};
}

function saveProcessedEmails(data) {
  const filepath = path.join(__dirname, 'processed-emails.json');
  fs.writeFileSync(filepath, JSON.stringify(data, null, 2));
}

function loadTasks() {
  const filepath = path.join(__dirname, 'tasks.json');
  if (fs.existsSync(filepath)) {
    return JSON.parse(fs.readFileSync(filepath, 'utf8'));
  }
  return [];
}

function saveTasks(tasks) {
  const filepath = path.join(__dirname, 'tasks.json');
  fs.writeFileSync(filepath, JSON.stringify(tasks, null, 2));
}

function createTask(email, subject, person, companies) {
  const parsed = parseSubject(subject);
  const company = companies.length > 0 ? companies[0] : (parsed.companies[0] || '');

  return {
    id: Date.now() + Math.random(),
    title: parsed.taskTitle,
    desc: `Email de ${email.from.emailAddress.address}`,
    urgency: 'Media',
    company: company,
    type: 'Auto',
    completed: false,
    date: new Date().toISOString().split('T')[0],
    person: person,
    emailHash: generateTaskHash(email),
    emailId: email.id,
    receivedDateTime: email.receivedDateTime
  };
}

async function syncEmails() {
  try {
    console.log('🚀 Iniciando sincronización de Outlook...');

    const token = await getAccessToken();
    console.log('✅ Token obtenido');

    const response = await getEmailsFromOutlook(token);
    const emails = response.value || [];
    console.log(`✉️  Encontrados ${emails.length} emails`);

    const processedEmails = loadProcessedEmails();
    const tasks = loadTasks();
    let newTasksCount = 0;

    for (const email of emails) {
      const hash = generateTaskHash(email);

      if (processedEmails[hash]) {
        console.log(`⏭️  Saltando (ya procesado): ${email.subject}`);
        continue;
      }

      const fromEmail = email.from.emailAddress.address.toLowerCase();
      const toRecipients = email.toRecipients.map(r => r.emailAddress.address.toLowerCase());
      let assignedPerson = null;
      let isPersonalTask = false;

      if (toRecipients.includes(OUTLOOK_EMAIL.toLowerCase())) {
        isPersonalTask = true;
        assignedPerson = null;
      } else {
        for (const recipient of toRecipients) {
          const person = identifyPerson(recipient);
          if (person) {
            assignedPerson = person;
            break;
          }
        }
      }

      const parsed = parseSubject(email.subject);

      const newTask = createTask(
        email,
        email.subject,
        assignedPerson,
        parsed.companies
      );

      if (!newTask.company) {
        newTask.company = '';
      }
      if (!newTask.person && !isPersonalTask) {
        newTask.person = null;
      }

      tasks.push(newTask);
      processedEmails[hash] = {
        subject: email.subject,
        processedAt: new Date().toISOString(),
        emailId: email.id
      };

      newTasksCount++;
      console.log(`✅ Nueva tarea: "${newTask.title}" → ${newTask.person || 'Mis Tareas'}`);
    }

    saveTasks(tasks);
    saveProcessedEmails(processedEmails);

    console.log(`\n✨ Sincronización completada`);
    console.log(`📊 Nuevas tareas: ${newTasksCount}`);
    console.log(`💾 Total tareas en sistema: ${tasks.length}`);

    return { success: true, newTasks: newTasksCount, totalTasks: tasks.length };

  } catch (error) {
    console.error('❌ Error en sincronización:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  syncEmails();
}

module.exports = { syncEmails, parseSubject, identifyPerson };
