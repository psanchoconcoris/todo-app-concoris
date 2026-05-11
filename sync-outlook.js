#!/usr/bin/env node
/**
 * SYNC-OUTLOOK.JS - Versión IMAP
 * Script que sincroniza emails de Outlook con la aplicación de tareas
 * Usa IMAP en lugar de Microsoft Graph
 */

const Imap = require('imap');
const { simpleParser } = require('mailparser');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// CONFIGURACIÓN
const OUTLOOK_EMAIL = 'psancho@concoris.es';
const OUTLOOK_PASSWORD = process.env.OUTLOOK_PASSWORD;

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

function generateTaskHash(messageId) {
  return crypto.createHash('md5').update(messageId).digest('hex');
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
    desc: `Email de ${email.from || 'desconocido'}`,
    urgency: 'Media',
    company: company,
    type: 'Auto',
    completed: false,
    date: new Date().toISOString().split('T')[0],
    person: person,
    emailHash: email.hash,
    emailId: email.messageId,
    receivedDateTime: email.date
  };
}

async function syncEmails() {
  return new Promise((resolve, reject) => {
    const imap = new Imap({
      user: OUTLOOK_EMAIL,
      password: OUTLOOK_PASSWORD,
      host: 'outlook.office365.com',
      port: 993,
      tls: true
    });

    const processedEmails = loadProcessedEmails();
    const tasks = loadTasks();
    let newTasksCount = 0;

    function openInbox(cb) {
      imap.openBox('INBOX', false, cb);
    }

    imap.openBox('INBOX', false, function(err, box) {
      if (err) {
        reject(err);
        return;
      }

      // Buscar emails del último día
      const oneDayAgo = new Date();
      oneDayAgo.setDate(oneDayAgo.getDate() - 1);

      imap.search(['SINCE', oneDayAgo], function(err, results) {
        if (err) {
          reject(err);
          return;
        }

        if (results.length === 0) {
          console.log('✉️  No hay emails nuevos');
          imap.end();
          resolve({ success: true, newTasks: 0, totalTasks: tasks.length });
          return;
        }

        console.log(`✉️  Encontrados ${results.length} emails`);

        const f = imap.fetch(results, { bodies: '' });
        let processedCount = 0;

        f.on('message', function(msg, seqno) {
          let emailData = {
            messageId: null,
            subject: '',
            from: '',
            to: [],
            date: new Date()
          };

          msg.on('attributes', function(attrs) {
            emailData.messageId = attrs.uid.toString();
          });

          msg.on('structure', function(structure) {
            // Parse headers
            const headers = {};
            for (let part of structure) {
              if (part.params) {
                headers[part.type] = part;
              }
            }
          });

          simpleParser(msg, async (err, parsed) => {
            if (err) {
              console.error('Error parsing email:', err);
              return;
            }

            emailData.subject = parsed.subject || '';
            emailData.from = parsed.from.text || '';
            emailData.to = parsed.to ? parsed.to.text.split(',').map(e => e.trim()) : [];
            emailData.date = parsed.date || new Date();
            emailData.hash = generateTaskHash(emailData.messageId);

            // Verificar si ya fue procesado
            if (processedEmails[emailData.hash]) {
              console.log(`⏭️  Saltando (ya procesado): ${emailData.subject}`);
              processedCount++;
              if (processedCount === results.length) {
                imap.end();
              }
              return;
            }

            // Identificar persona
            let assignedPerson = null;
            for (const recipient of emailData.to) {
              const person = identifyPerson(recipient);
              if (person) {
                assignedPerson = person;
                break;
              }
            }

            const parsed_subject = parseSubject(emailData.subject);

            // Crear tarea
            const newTask = createTask(
              emailData,
              emailData.subject,
              assignedPerson,
              parsed_subject.companies
            );

            tasks.push(newTask);
            processedEmails[emailData.hash] = {
              subject: emailData.subject,
              processedAt: new Date().toISOString(),
              emailId: emailData.messageId
            };

            newTasksCount++;
            console.log(`✅ Nueva tarea: "${newTask.title}" → ${newTask.person || 'Mis Tareas'}`);

            processedCount++;
            if (processedCount === results.length) {
              imap.end();
            }
          });
        });

        f.on('error', reject);
        f.on('end', function() {
          setTimeout(() => {
            saveTasks(tasks);
            saveProcessedEmails(processedEmails);

            console.log(`\n✨ Sincronización completada`);
            console.log(`📊 Nuevas tareas: ${newTasksCount}`);
            console.log(`💾 Total tareas en sistema: ${tasks.length}`);

            resolve({ success: true, newTasks: newTasksCount, totalTasks: tasks.length });
          }, 1000);
        });
      });
    });

    imap.on('error', reject);
    imap.on('end', function() {
      console.log('Conexión IMAP cerrada');
    });

    imap.openBox('INBOX', false, function(err, box) {
      if (err) reject(err);
    });

    imap.openBox('INBOX', false, function(err) {
      if (err) {
        reject(err);
      } else {
        imap.search(['UNSEEN'], function(err, results) {
          if (err) reject(err);
        });
      }
    });
  });
}

if (require.main === module) {
  syncEmails()
    .then(result => {
      console.log('Success:', result);
      process.exit(0);
    })
    .catch(err => {
      console.error('❌ Error:', err.message);
      process.exit(1);
    });
}

module.exports = { syncEmails };
