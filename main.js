// main.js

// Loading indicator
function showLoading() {
  document.getElementById('loading-bar').classList.add('active');
}

function hideLoading() {
  setTimeout(() => {
    document.getElementById('loading-bar').classList.remove('active');
  }, 300);
}

document.addEventListener('DOMContentLoaded', async () => {
  const msalConfig = {
    auth: {
      clientId: '8e23d112-104b-4e6a-a57d-7e3a2a61d837',
      authority:
        'https://login.microsoftonline.com/20e27700-b670-4553-a27c-d8e2583b3289',
      redirectUri:
        'https://work-qdc.github.io/Scheduling-Dashboard/'
    },
    cache: { cacheLocation: 'localStorage' }
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);

  async function signIn() {
    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) {
      await msalInstance.loginPopup({
        scopes: ['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All']
      });
    }
  }

  async function getToken(scopes) {
    try {
      return (await msalInstance.acquireTokenSilent({ scopes })).accessToken;
    } catch {
      return (await msalInstance.acquireTokenPopup({ scopes })).accessToken;
    }
  }

  function client(token) {
    return MicrosoftGraph.Client.init({ authProvider: done => done(null, token) });
  }

  let EMP_LIST_ID,
      SCHED_LIST_ID;

  async function getSiteId(g) {
    const site = await g
      .api(`/sites/tcco.sharepoint.com:/sites/SchedulingToolTest:/`)
      .get();
    return site.id;
  }

  async function initLists() {
    showLoading();
    await signIn();
    const token = await getToken(['Sites.Read.All']);
    const g = client(token);
    const sid = await getSiteId(g);
    const res = await g.api(`/sites/${sid}/lists`).get();
    EMP_LIST_ID = res.value.find(
      l => l.displayName === 'SchedulingEmployeeTest'
    ).id;
    SCHED_LIST_ID = res.value.find(
      l => l.displayName === 'TestTableScheduling'
    ).id;
    hideLoading();
  }

  async function fetchEmployees() {
    showLoading();
    const token = await getToken(['Sites.Read.All']);
    const g = client(token);
    const sid = await getSiteId(g);
    const r = await g
      .api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`)
      .expand('fields')
      .get();
    window.employees = r.value.map(i => ({
      id: i.id,
      name: i.fields.Title,
      email: i.fields.Employeesemail
    }));
    populateEmployeeSelects();
    renderEmployeeTable();
    hideLoading();
  }

  async function fetchEntries() {
    showLoading();
    const token = await getToken(['Sites.Read.All']);
    const g = client(token);
    const sid = await getSiteId(g);
    const r = await g
      .api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`)
      .expand('fields')
      .get();
    window.entries = r.value.map(i => ({
      id: i.id,
      title: i.fields.Title || '',
      topic: i.fields.Topic || '',
      status: i.fields.RegistrationStatus || '',
      start: i.fields.StartDate || '',
      end: i.fields.EndDate || '',
      location: i.fields.Location || '',
      link: i.fields.Link || '',
      industry: i.fields.Industry || '',
      desc: i.fields.Description || '',
      applicationdeadline: i.fields.ApplicationDeadline || '',
      internalExternal: i.fields.Internal_x002f_External || '',
      AssignedEmployee: i.fields.AssignedEmployee || '',
      notes: i.fields.Notes || ''
    }));
    renderList();
    hideLoading();
  }

  async function addEmployee() {
    const name = prompt('Enter employee name:');
    const email = prompt('Enter employee email:');
    if (!name || !email) return;
    showLoading();
    const token = await getToken(['Sites.ReadWrite.All']);
    const g = client(token);
    const sid = await getSiteId(g);
    await g.api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`).post({
      fields: { Title: name, Employeesemail: email }
    });
    await fetchEmployees();
  }

  async function addEntry() {
    try {
      showLoading();
      const token = await getToken(['Sites.ReadWrite.All']);
      const g = client(token);
      const sid = await getSiteId(g);

      const sel = document.getElementById('new-AssignedEmployee');
      const assignedEmployeeName = sel.value;
      const assignedEmployeeEmail =
        (window.employees.find(e => e.name === assignedEmployeeName) || {})
          .email;

      const fields = {
        Title: document.getElementById('new-title').value,
        Topic: document.getElementById('new-topic').value,
        RegistrationStatus: document.getElementById('new-status').value,
        TypeofEngagement: document.getElementById('new-type').value,
        StartDate: document.getElementById('new-start').value,
        EndDate: document.getElementById('new-end').value,
        Location: document.getElementById('new-location').value,
        Link: document.getElementById('new-link').value,
        Industry: document.getElementById('new-industry').value,
        Description: document.getElementById('new-desc').value,
        ApplicationDeadline: document.getElementById(
          'new-applicationdeadline'
        ).value,
        Internal_x002f_External: document.getElementById('new-type').value,
        AssignedEmployee: assignedEmployeeName,
        Notes: document.getElementById('new-notes').value
      };

      await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`).post({
        fields
      });

      // Send calendar invite if email valid...
      if (assignedEmployeeEmail && assignedEmployeeEmail.includes('@')) {
        const event = {
          subject: fields.Title || 'Scheduled Event',
          body: {
            contentType: 'Text',
            content:
              fields.Description || 'Scheduled event via scheduling tool.'
          },
          start: {
            dateTime: `${fields.StartDate}T09:00:00`,
            timeZone: 'UTC'
          },
          end: {
            dateTime: `${fields.EndDate}T17:00:00`,
            timeZone: 'UTC'
          },
          location: { displayName: fields.Location || 'TBD' },
          attendees: [
            {
              emailAddress: {
                address: assignedEmployeeEmail,
                name: assignedEmployeeName
              },
              type: 'required'
            }
          ],
          responseRequested: true,
          isOnlineMeeting: false
        };

        try {
          await g
            .api('/me/events')
            .query({ sendInvitations: true })
            .post(event);
          console.log('✅ Calendar invite sent to:', assignedEmployeeEmail);
        } catch (inviteErr) {
          console.error('❌ Failed to send calendar invite:', inviteErr);
          alert(
            `Event created, but failed to send invite:\n${inviteErr.message}`
          );
        }
      }

      await fetchEntries();
      clearAddForm();
    } catch (err) {
      console.error('addEntry failed:', err);
      alert(`Error adding entry:\n${err.message}`);
    } finally {
      hideLoading();
    }
  }

  // expose addEntry globally
  window.addEntry = addEntry;

  function clearAddForm() {
    document.getElementById('new-title').value = '';
    document.getElementById('new-topic').value = '';
    document.getElementById('new-status').value = 'Pending';
    document.getElementById('new-start').value = '';
    document.getElementById('new-end').value = '';
    document.getElementById('new-location').value = '';
    document.getElementById('new-link').value = '';
    document.getElementById('new-industry').value = '';
    document.getElementById('new-desc').value = '';
    document.getElementById('new-applicationdeadline').value = '';
    document.getElementById('new-AssignedEmployee').value = '';
    document.getElementById('new-notes').value = '';
  }

  function populateEmployeeSelects() {
    const sel       = document.getElementById('new-AssignedEmployee');
    const filterSel = document.getElementById('employee-filter');
    const calSel    = document.getElementById('calendar-employee-filter');

    if (sel)       sel.innerHTML       = '<option value="">Select Speaker</option>';
    if (filterSel) filterSel.innerHTML = '<option value="">All Speakers</option>';
    if (calSel)    calSel.innerHTML    = '<option value="">All Speakers</option>';

    window.employees.forEach(e => {
      if (sel) {
        const o1 = document.createElement('option');
        o1.value = e.name;
        o1.textContent = e.name;
        sel.appendChild(o1);
      }
      if (filterSel) {
        const o2 = document.createElement('option');
        o2.value = e.name;
        o2.textContent = e.name;
        filterSel.appendChild(o2);
      }
      if (calSel) {
        const o3 = document.createElement('option');
        o3.value = e.name;
        o3.textContent = e.name;
        calSel.appendChild(o3);
      }
    });
  }

  function renderEmployeeTable() {
    const body = document.getElementById('employee-table-body');
    body.innerHTML = '';
    window.employees.forEach(e => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${e.name}</td>
        <td>${e.email}</td>
        <td class="text-center">
          <button class="action-btn btn-delete" data-id="${e.id}">
            <i class="bi bi-trash"></i>
          </button>
        </td>`;
      tr.querySelector('button').addEventListener('click', async () => {
        if (confirm(`Delete employee ${e.name}?`)) {
          showLoading();
          const token = await getToken(['Sites.ReadWrite.All']);
          const g = client(token);
          const sid = await getSiteId(g);
          await g
            .api(`/sites/${sid}/lists/${EMP_LIST_ID}/items/${e.id}`)
            .delete();
          await fetchEmployees();
        }
      });
      body.appendChild(tr);
    });
  }

  function getStatusBadge(status) {
    const map = {
      Confirmed: 'status-confirmed',
      Pending:   'status-pending',
      Cancelled: 'status-cancelled',
      Completed: 'status-completed'
    };
    const cls = map[status] || 'status-pending';
    return `<span class="status-badge ${cls}">${status||'Pending'}</span>`;
  }

  function getTypePill(type) {
    const cls = type === 'Internal' ? 'type-internal' : 'type-external';
    return `<span class="type-pill ${cls}">${type || 'Internal'}</span>`;
  }

  function formatDate(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return d.toLocaleDateString('en-US', {
      month: 'short',
      day:   'numeric',
      year:  'numeric'
    });
  }

  function formatDateOnly(iso) {
    if (!iso || typeof iso !== 'string') return '';
    return iso.split('T')[0];
  }

  function truncate(str, len) {
    if (!str) return '';
    return str.length > len ? str.slice(0, len) + '...' : str;
  }

  function renderCalendarView() {
    const container = document.getElementById('calendar-table-wrapper');
    container.innerHTML = '';
    const calendarEl = document.createElement('div');
    container.appendChild(calendarEl);

    const selectedEmployee = document.getElementById(
      'calendar-employee-filter'
    ).value;

    const filteredEvents = window.entries
      .filter(e => !selectedEmployee || e.AssignedEmployee === selectedEmployee)
      .map(e => ({
        id: e.id,
        title: e.title || '(No Title)',
        start: e.start,
        end:   e.end,
        backgroundColor:
          e.status === 'Confirmed'
            ? '#2DCE89'
            : e.status === 'Cancelled'
            ? '#F5365C'
            : e.status === 'Completed'
            ? '#11CDEF'
            : '#FB6340',
        extendedProps: {
          location: e.location,
          notes:    e.notes,
          status:   e.status,
          topic:    e.topic
        }
      }));

    const calendar = new FullCalendar.Calendar(calendarEl, {
      initialView: 'dayGridMonth',
      height:      'auto',
      headerToolbar: {
        left:   'prev,next today',
        center: 'title',
        right:  'dayGridMonth,timeGridWeek,listWeek'
      },
      events: filteredEvents,
      eventClick(info) {
        const { title, start, end, extendedProps, id } = info.event;
        const modal =
          document.getElementById('eventModal') ||
          createEventModal();

        // 1) populate the modal body
        document.getElementById('eventModalBody').innerHTML = `
          <p><strong>Title:</strong> ${title}</p>
          <p><strong>Topic:</strong> ${extendedProps.topic || 'N/A'}</p>
          <p><strong>Status:</strong> ${getStatusBadge(extendedProps.status)}</p>
          <p><strong>Start:</strong> ${formatDate(start)}</p>
          <p><strong>End:</strong> ${formatDate(end)}</p>
          <p><strong>Location:</strong> ${extendedProps.location || 'N/A'}</p>
          <p><strong>Notes:</strong> ${extendedProps.notes || 'None'}</p>
        `;

        // 2) hook up the Edit button
        modal.querySelector('#editEventBtn').onclick = () =>
          editEntryFromCalendar(id);

        // 3) show the modal
        new bootstrap.Modal(modal).show();
      }
    });

    calendar.render();
  }

  // === UPDATED createEventModal with footer Edit button ===
  function createEventModal() {
    const modal = document.createElement('div');
    modal.className = 'modal fade';
    modal.id        = 'eventModal';
    modal.innerHTML = `
      <div class="modal-dialog">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title">Event Details</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body" id="eventModalBody"></div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
              Close
            </button>
            <button type="button" class="btn btn-primary" id="editEventBtn">
              <i class="bi bi-pencil me-1"></i>Edit
            </button>
          </div>
        </div>
      </div>`;
    document.body.appendChild(modal);
    return modal;
  }

  // === NEW helper to jump into list-view edit ===
  function editEntryFromCalendar(entryId) {
    const modalEl = document.getElementById('eventModal');
    const bsModal = bootstrap.Modal.getInstance(modalEl);
    if (bsModal) bsModal.hide();

    // switch to list view
    document.getElementById('list-view-btn').click();

    // allow panel swap then click the edit button
    setTimeout(() => {
      document.querySelectorAll('#entries-tbody tr').forEach(row => {
        const btn = row.querySelector('.edit-entry');
        if (btn && btn.dataset.id === String(entryId)) {
          btn.click();
        }
      });
    }, 200);
  }

  function renderList() {
    const tbody = document.getElementById('entries-tbody');
    tbody.innerHTML = '';

    if (!window.entries.length) {
      tbody.innerHTML = `
        <tr>
          <td colspan="15" class="empty-state">
            <i class="bi bi-calendar-x"></i>
            <p>No conferences scheduled yet. Add your first conference using the new Add Entry button!</p>
          </td>
        </tr>`;
      return;
    }

    window.entries.forEach((e, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td><strong>${e.title}</strong></td>
        <td>${e.topic}</td>
        <td>${getStatusBadge(e.status)}</td>
        <td>${formatDate(e.start)}</td>
        <td>${formatDate(e.end)}</td>
        <td>${e.location}</td>
        <td>${e.link ? `<a href="${e.link}" target="_blank"><i class="bi bi-link-45deg"></i></a>` : ''}</td>
        <td>${e.industry}</td>
        <td><small>${truncate(e.desc, 50)}</small></td>
        <td>${formatDate(e.applicationdeadline)}</td>
        <td>${getTypePill(e.internalExternal)}</td>
        <td>${e.AssignedEmployee || '<em>Unassigned</em>'}</td>
        <td><small>${truncate(e.notes, 50)}</small></td>
        <td class="text-center">
          <button class="action-btn btn-edit edit-entry" data-id="${e.id}" title="Edit">
            <i class="bi bi-pencil"></i>
          </button>
          <button class="action-btn btn-delete del-entry" data-id="${e.id}" title="Delete">
            <i class="bi bi-trash"></i>
          </button>
        </td>`;

      tr.querySelector('.del-entry').addEventListener('click', async () => {
        if (confirm(`Delete "${e.title}"?`)) {
          showLoading();
          const token = await getToken(['Sites.ReadWrite.All']);
          const g = client(token);
          const sid = await getSiteId(g);
          await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items/${e.id}`).delete();
          await fetchEntries();
        }
      });

      tr.querySelector('.edit-entry').addEventListener('click', () => {
        enableEditMode(tr, e);
      });

      tbody.appendChild(tr);
    });

    applyFilters();
  }

  function applyFilters() {
    const selectedSpeaker = document.getElementById('employee-filter').value;
    const month           = document.getElementById('start-date-filter').value;
    const query           = document.getElementById('keyword-search').value.trim().toLowerCase();
    const rows = document.querySelectorAll('#entries-tbody tr');

    rows.forEach((row, index) => {
      const entry = window.entries[index];
      if (!entry) return;
      let show = true;

      if (selectedSpeaker && entry.AssignedEmployee !== selectedSpeaker) show = false;

      if (month && entry.start) {
        const entryMonth = entry.start.substring(0, 7);
        if (entryMonth !== month) show = false;
      }

      if (query) {
        const text = Array.from(row.cells).map(td => td.textContent.toLowerCase()).join(' ');
        if (!text.includes(query)) show = false;
      }

      row.style.display = show ? '' : 'none';
    });
  }

  function enableEditMode(row, entry) {
    // ... your existing inline‐edit logic ...
  }

  // ---- EXISTING BUTTON & VIEW HANDLERS ----

  document
    .getElementById('edit-employees-btn')
    .addEventListener('click', () =>
      new bootstrap.Modal(document.getElementById('employeeModal')).show()
    );
  document.getElementById('add-emp-btn').addEventListener('click', addEmployee);

  document
  .getElementById('open-add-entry-ui-btn')
  .addEventListener('click', () => {
    const modalEl = document.getElementById('addEntryModal');
    const addEntryModal = new bootstrap.Modal(modalEl);
    addEntryModal.show();
  });
// Cancel = clear the form & hide the modal
document
  .getElementById('cancel-entry-btn')
  .addEventListener('click', () => {
    clearAddForm();
    const modalEl = document.getElementById('addEntryModal');
    bootstrap.Modal.getInstance(modalEl).hide();
  });

// Save = add entry, clear form & hide modal
document
  .getElementById('submit-entry-btn')
  .addEventListener('click', async () => {
    await addEntry();
    clearAddForm();
    const modalEl = document.getElementById('addEntryModal');
    bootstrap.Modal.getInstance(modalEl).hide();
  });

  document
    .getElementById('employee-filter')
    .addEventListener('change', applyFilters);
  document
    .getElementById('start-date-filter')
    .addEventListener('change', applyFilters);
  document
    .getElementById('keyword-search')
    .addEventListener('input', applyFilters);
  document
    .getElementById('calendar-employee-filter')
    .addEventListener('change', renderCalendarView);

  document
    .getElementById('calendar-view-btn')
    .addEventListener('click', () => {
      document.getElementById('list-view-btn').classList.remove('active');
      document.getElementById('calendar-view-btn').classList.add('active');
      document.getElementById('calendar-container').style.display = 'block';
      document.getElementById('list-container').style.display     = 'none';
      renderCalendarView();
    });
  document
    .getElementById('list-view-btn')
    .addEventListener('click', () => {
      document.getElementById('calendar-view-btn').classList.remove('active');
      document.getElementById('list-view-btn').classList.add('active');
      document.getElementById('calendar-container').style.display = 'none';
      document.getElementById('list-container').style.display     = 'block';
    });
  document
    .getElementById('calendar-list-view-btn')
    .addEventListener('click', () =>
      document.getElementById('list-view-btn').click()
    );

  // Initialize data
  try {
    await initLists();
    await fetchEmployees();
    await fetchEntries();
  } catch (err) {
    console.error('Initialization error:', err);
    alert(`Initialization error:\n${err.message}`);
    hideLoading();
  }
});
