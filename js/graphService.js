// js/graphService.js

import { getToken, signIn } from './auth.js';
import * as utils            from './utils.js';
import { Client }            from 'https://cdn.jsdelivr.net/npm/@microsoft/microsoft-graph-client/lib/graph-js-sdk-web.js';

let EMP_LIST_ID, SCHED_LIST_ID;

async function getSiteId(g) {
  const site = await g
    .api(`/sites/tcco.sharepoint.com:/sites/SchedulingToolTest:/`)
    .get();
  return site.id;
}

function client(token) {
  return Client.init({ authProvider: done => done(null, token) });
}

// —— all of these must be exported —— //

export async function initLists() {
  utils.showLoading();
  await signIn();
  const token = await getToken(['Sites.Read.All']);
  const g     = client(token);
  const sid   = await getSiteId(g);
  const res   = await g.api(`/sites/${sid}/lists`).get();

  EMP_LIST_ID   = res.value.find(l => l.displayName === 'SchedulingEmployeeTest').id;
  SCHED_LIST_ID = res.value.find(l => l.displayName === 'TestTableScheduling'   ).id;
  utils.hideLoading();
}

export async function fetchEmployees() {
  utils.showLoading();
  const token = await getToken(['Sites.Read.All']);
  const g     = client(token);
  const sid   = await getSiteId(g);
  const r     = await g
    .api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`)
    .expand('fields')
    .get();

  window.employees = r.value.map(i => ({
    id:    i.id,
    name:  i.fields.Title,
    email: i.fields.Employeesemail
  }));
  utils.hideLoading();
}

export async function fetchEntries() {
  utils.showLoading();
  const token = await getToken(['Sites.Read.All']);
  const g     = client(token);
  const sid   = await getSiteId(g);
  const r     = await g
    .api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`)
    .expand('fields')
    .get();

  window.entries = r.value.map(i => ({
    id:                   i.id,
    title:                i.fields.Title || '',
    topic:                i.fields.Topic || '',
    status:               i.fields.RegistrationStatus || '',
    start:                i.fields.StartDate || '',
    end:                  i.fields.EndDate || '',
    location:             i.fields.Location || '',
    link:                 i.fields.Link || '',
    industry:             i.fields.Industry || '',
    desc:                 i.fields.Description || '',
    applicationdeadline:  i.fields.ApplicationDeadline || '',
    internalExternal:     i.fields.Internal_x002f_External || '',
    AssignedEmployee:     i.fields.AssignedEmployee || '',
    notes:                i.fields.Notes || ''
  }));
  utils.hideLoading();
}

export async function addEmployee() {
  const name  = prompt('Enter employee name:');
  const email = prompt('Enter employee email:');
  if (!name || !email) return;

  utils.showLoading();
  const token = await getToken(['Sites.ReadWrite.All']);
  const g     = client(token);
  const sid   = await getSiteId(g);

  await g
    .api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`)
    .post({ fields: { Title: name, Employeesemail: email } });

  await fetchEmployees();
  utils.hideLoading();
}

export async function addEntry() {
  utils.showLoading();
  try {
    const token = await getToken(['Sites.ReadWrite.All']);
    const g     = client(token);
    const sid   = await getSiteId(g);

    // build your new item fields:
    const assignedName  = document.getElementById('new-AssignedEmployee').value;
    const fields = {
      Title:                document.getElementById('new-title').value,
      Topic:                document.getElementById('new-topic').value,
      RegistrationStatus:   document.getElementById('new-status').value,
      StartDate:            document.getElementById('new-start').value,
      EndDate:              document.getElementById('new-end').value,
      Location:             document.getElementById('new-location').value,
      Link:                 document.getElementById('new-link').value,
      Industry:             document.getElementById('new-industry').value,
      Description:          document.getElementById('new-desc').value,
      ApplicationDeadline:  document.getElementById('new-applicationdeadline').value,
      Internal_x002f_External: assignedName,
      AssignedEmployee:       assignedName,
      Notes:                document.getElementById('new-notes').value
    };

    await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`).post({ fields });
    await fetchEntries();
  } catch (err) {
    console.error('addEntry error:', err);
    alert(`Error adding entry:\n${err.message}`);
  } finally {
    utils.hideLoading();
  }
}
