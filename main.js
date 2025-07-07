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
      authority: 'https://login.microsoftonline.com/20e27700-b670-4553-a27c-d8e2583b3289',
      redirectUri: 'https://work-qdc.github.io/Scheduling-Dashboard/'
    },
    cache: { cacheLocation: 'localStorage' }
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);

  async function signIn() {
    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) {
      await msalInstance.loginPopup({ scopes: ['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All'] });
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

  let EMP_LIST_ID, SCHED_LIST_ID;

  async function getSiteId(g) {
    const site = await g.api(`/sites/tcco.sharepoint.com:/sites/SchedulingToolTest:/`).get();
    return site.id;
  }

  async function initLists() {
    showLoading();
    await signIn();
    const token = await getToken(['Sites.Read.All']);
    const g = client(token);
    const sid = await getSiteId(g);
    const res = await g.api(`/sites/${sid}/lists`).get();
    EMP_LIST_ID = res.value.find(l => l.displayName === 'SchedulingEmployeeTest').id;
    SCHED_LIST_ID = res.value.find(l => l.displayName === 'TestTableScheduling').id;
    hideLoading();
  }

  async function fetchEmployees() {
    showLoading();
    const token = await getToken(['Sites.Read.All']);
    const g = client(token);
    const sid = await getSiteId(g);
    const r = await g.api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`).expand('fields').get();
    window.employees = r.value.map(i => ({ id: i.id, name: i.fields.Title, email: i.fields.Employeesemail }));
    populateEmployeeSelects();
    renderEmployeeTable();
    hideLoading();
  }

  async function fetchEntries() {
    showLoading();
    const token = await getToken(['Sites.Read.All']);
    const g = client(token);
    const sid = await getSiteId(g);
    const r = await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`).expand('fields').get();
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
    await g.api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`).post({ fields: { Title: name, Employeesemail: email } });
    await fetchEmployees();
  }

  async function addEntry() {
    try {
      showLoading();
      const token = await getToken(['Sites.ReadWrite.All']);
      const g = client(token);
      const sid = await getSiteId(g);

      const sel = document.getElementById('new-AssignedEmployee');
      const selectedOption = sel.options[sel.selectedIndex];
      const assignedEmployeeName = selectedOption.value;
      const assignedEmployeeEmail = (window.employees.find(e => e.name === assignedEmployeeName) || {}).email;

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
        ApplicationDeadline: document.getElementById('new-applicationdeadline').value,
        Internal_x002f_External: document.getElementById('new-type').value,
        AssignedEmployee: assignedEmployeeName,
        Notes: document.getElementById('new-notes').value
      };

      await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`).post({ fields });

      // Send calendar invite
      if (assignedEmployeeEmail && assignedEmployeeEmail.includes('@')) {
        const event = {
          subject: fields.Title || 'Scheduled Event',
          body: {
            contentType: 'Text',
            content: fields.Description || 'Scheduled event via scheduling tool.'
          },
          start: {
            dateTime: `${fields.StartDate}T09:00:00`,
            timeZone: 'UTC'
          },
          end: {
            dateTime: `${fields.EndDate}T17:00:00`,
            timeZone: 'UTC'
          },
          location: {
            displayName: fields.Location || 'TBD'
          },
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
          await g.api('/me/events')
            .query({ sendInvitations: true })
            .post(event);
          console.log('✅ Calendar invite sent to:', assignedEmployeeEmail);
        } catch (inviteErr) {
          console.error('❌ Failed to send calendar invite:', inviteErr);
          alert(`Event created, but failed to send invite:\n${inviteErr.message}`);
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
    const sel = document.getElementById('new-AssignedEmployee');
    const filterSel = document.getElementById('employee-filter');
    const calSel = document.getElementById('calendar-employee-filter');

    sel.innerHTML = '<option value="">Select Speaker</option>';
    filterSel.innerHTML = '<option value="">All Employees</option>';
    calSel.innerHTML = '<option value="">All Employees</option>';

    window.employees.forEach(e => {
      const o1 = document.createElement('option');
      o1.value = e.name;
      o1.textContent = e.name;
      sel.appendChild(o1);

      const o2 = o1.cloneNode(true);
      filterSel.appendChild(o2);

      const o3 = o1.cloneNode(true);
      calSel.appendChild(o3);
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
          await g.api(`/sites/${sid}/lists/${EMP_LIST_ID}/items/${e.id}`).delete();
          await fetchEmployees();
        }
      });
      body.appendChild(tr);
    });
  }

  function getStatusBadge(status) {
    const statusClass = {
      'Confirmed': 'status-confirmed',
      'Pending': 'status-pending',
      'Cancelled': 'status-cancelled',
      'Completed': 'status-completed'
    }[status] || 'status-pending';
    
    return `<span class="status-badge ${statusClass}">${status || 'Pending'}</span>`;
  }

  function getTypePill(type) {
    const typeClass = type === 'Internal' ? 'type-internal' : 'type-external';
    return `<span class="type-pill ${typeClass}">${type || 'Internal'}</span>`;
  }

  function formatDate(dateStr) {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function formatDateOnly(isoString) {
    if (!isoString || typeof isoString !== 'string') return '';
    return isoString.split('T')[0];
  }

  function truncate(str, len) {
    if (!str) return '';
    return str.length > len ? str.substring(0, len) + '...' : str;
  }

  function renderCalendarView() {
    const container = document.getElementById('calendar-table-wrapper');
    container.innerHTML = '';

    const calendarEl = document.createElement('div');
    container.appendChild(calendarEl);

    const selectedEmployee = document.getElementById('calendar-employee-filter').value;

    const filteredEvents = window.entries
      .filter(e => !selectedEmployee || e.AssignedEmployee === selectedEmployee)
      .map(e => ({
        id: e.id,
        title: e.title || '(No Title)',
        start: e.start,
        end: e.end,
        backgroundColor: e.status === 'Confirmed' ? '#2DCE89' : 
                       e.status === 'Cancelled' ? '#F5365C' : 
                       e.status === 'Completed' ? '#11CDEF' : '#FB6340',
        extendedProps: {
          location: e.location,
          notes: e.notes,
          status: e.status,
          topic: e.topic
        }
      }));

    const calendar = new FullCalendar.Calendar(calendarEl, {
      initialView: 'dayGridMonth',
      height: 'auto',
      headerToolbar: {
        left: 'prev,next today',
        center: 'title',
        right: 'dayGridMonth,timeGridWeek,listWeek'
      },
      events: filteredEvents,
      eventClick: function(info) {
        const { title, start, end, extendedProps } = info.event;
        const modal = new bootstrap.Modal(document.getElementById('eventModal') || createEventModal());
        document.getElementById('eventModalBody').innerHTML = `
          <p><strong>Title:</strong> ${title}</p>
          <p><strong>Topic:</strong> ${extendedProps.topic || 'N/A'}</p>
          <p><strong>Status:</strong> ${getStatusBadge(extendedProps.status)}</p>
          <p><strong>Start:</strong> ${formatDate(start)}</p>
          <p><strong>End:</strong> ${formatDate(end)}</p>
          <p><strong>Location:</strong> ${extendedProps.location || 'N/A'}</p>
          <p><strong>Notes:</strong> ${extendedProps.notes || 'None'}</p>
        `;
        modal.show();
      }
    });

    calendar.render();
  }

  function createEventModal() {
    const modal = document.createElement('div');
    modal.className = 'modal fade';
    modal.id = 'eventModal';
    modal.innerHTML = `
      <div class="modal-dialog">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title">Event Details</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body" id="eventModalBody"></div>
        </div>
      </div>
    `;
    document.body.appendChild(modal);
    return modal;
  }

  function renderList() {
    const tbody = document.getElementById('entries-tbody');
    tbody.innerHTML = '';
    
    if (window.entries.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="15" class="empty-state">
            <i class="bi bi-calendar-x"></i>
            <p>No conferences scheduled yet. Add your first conference using the form below!</p>
          </td>
        </tr>
      `;
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
        </td>
      `;

      // Delete logic
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

      // Edit logic
      tr.querySelector('.edit-entry').addEventListener('click', () => {
        enableEditMode(tr, e);
      });

      tbody.appendChild(tr);
    });
    applyFilters();
  }

  function applyFilters() {
    const selected = document.getElementById('employee-filter').value;
    const month = document.getElementById('start-date-filter').value;
    const rows = document.querySelectorAll('#entries-tbody tr');
    
    rows.forEach((row, index) => {
      const entry = window.entries[index];
      if (!entry) return;
      
      let show = true;
      
      if (selected && entry.AssignedEmployee !== selected) show = false;
      
      if (month && entry.start) {
        const entryMonth = entry.start.substring(0, 7);
        if (entryMonth !== month) show = false;
      }
      
      row.style.display = show ? '' : 'none';
    });
  }

  function enableEditMode(row, entry) {
    row.classList.add('editing-row');
    const td = row.querySelectorAll('td');
    
    td[1].innerHTML = `<input class="form-control form-control-sm" value="${entry.title}" />`;
    td[2].innerHTML = `<input class="form-control form-control-sm" value="${entry.topic}" />`;
    td[3].innerHTML = `<select class="form-select form-select-sm">
      <option ${entry.status === 'Pending' ? 'selected' : ''}>Pending</option>
      <option ${entry.status === 'Confirmed' ? 'selected' : ''}>Confirmed</option>
      <option ${entry.status === 'Cancelled' ? 'selected' : ''}>Cancelled</option>
      <option ${entry.status === 'Completed' ? 'selected' : ''}>Completed</option>
    </select>`;
    td[4].innerHTML = `<input type="date" class="form-control form-control-sm" value="${formatDateOnly(entry.start)}" />`;
    td[5].innerHTML = `<input type="date" class="form-control form-control-sm" value="${formatDateOnly(entry.end)}" />`;
    td[6].innerHTML = `<input class="form-control form-control-sm" value="${entry.location}" />`;
    td[7].innerHTML = `<input class="form-control form-control-sm" value="${entry.link}" />`;
    td[8].innerHTML = `<input class="form-control form-control-sm" value="${entry.industry}" />`;
    td[9].innerHTML = `<textarea class="form-control form-control-sm" rows="2">${entry.desc}</textarea>`;
    td[10].innerHTML = `<input type="date" class="form-control form-control-sm" value="${formatDateOnly(entry.applicationdeadline)}" />`;
    td[11].innerHTML = `<select class="form-select form-select-sm">
      <option ${entry.internalExternal === 'Internal' ? 'selected' : ''}>Internal</option>
      <option ${entry.internalExternal === 'External' ? 'selected' : ''}>External</option>
    </select>`;
    td[12].innerHTML = `<select class="form-select form-select-sm">
      <option value="">Unassigned</option>
      ${window.employees.map(e => 
        `<option value="${e.name}" ${e.name === entry.AssignedEmployee ? 'selected' : ''}>
          ${e.name}
        </option>`
      ).join('')}
    </select>`;
    td[13].innerHTML = `<textarea class="form-control form-control-sm" rows="2">${entry.notes}</textarea>`;
    td[14].innerHTML = `
      <button class="action-btn btn-save save-edit" title="Save">
        <i class="bi bi-check-lg"></i>
      </button>
      <button class="action-btn btn-cancel cancel-edit" title="Cancel">
        <i class="bi bi-x-lg"></i>
      </button>`;

    td[14].querySelector('.save-edit').addEventListener('click', async () => {
      showLoading();
      const updated = {};
      const set = (key, val) => {
        if (val !== '') updated[key] = val;
      };

      set('Title', td[1].querySelector('input').value);
      set('Topic', td[2].querySelector('input').value);
      set('RegistrationStatus', td[3].querySelector('select').value);
      set('StartDate', td[4].querySelector('input').value);
      set('EndDate', td[5].querySelector('input').value);
      set('Location', td[6].querySelector('input').value);
      set('Link', td[7].querySelector('input').value);
      set('Industry', td[8].querySelector('input').value);
      set('Description', td[9].querySelector('textarea').value);
      set('ApplicationDeadline', td[10].querySelector('input').value);
      set('Internal_x002f_External', td[11].querySelector('select').value);
      set('AssignedEmployee', td[12].querySelector('select').value);
      set('Notes', td[13].querySelector('textarea').value);

      try {
        const token = await getToken(['Sites.ReadWrite.All']);
        const g = client(token);
        const sid = await getSiteId(g);

        await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items/${entry.id}/fields`)
          .patch(updated);

        // Send calendar invite when an employee is assigned
        const assignedName = td[12].querySelector('select').value;
        const assignedEmail = (window.employees.find(e => e.name === assignedName) || {}).email;

        if (assignedEmail && assignedEmail.includes('@')) {
          const event = {
            subject: updated.Title || 'Scheduled Event',
            body: {
              contentType: 'Text',
              content: updated.Description || 'Scheduled event via scheduling tool.'
            },
            start: {
              dateTime: `${updated.StartDate || entry.start}T09:00:00`,
              timeZone: 'UTC'
            },
            end: {
              dateTime: `${updated.EndDate || entry.end}T17:00:00`,
              timeZone: 'UTC'
            },
            location: {
              displayName: updated.Location || 'TBD'
            },
            attendees: [
              {
                emailAddress: {
                  address: assignedEmail,
                  name: assignedName
                },
                type: 'required'
              }
            ],
            responseRequested: true,
            isOnlineMeeting: false
          };

          try {
            await g.api('/me/events')
              .query({ sendInvitations: true })
              .post(event);
            console.log('✅ Calendar invite sent to:', assignedEmail);
          } catch (inviteErr) {
            console.error('❌ Failed to send calendar invite:', inviteErr);
          }
        }

        await fetchEntries();
      } catch (err) {
        console.error('editEntry failed:', err);
        alert(`Failed to save edits:\n${err.message}`);
        hideLoading();
      }
    });

    td[14].querySelector('.cancel-edit').addEventListener('click', () => {
      renderList();
    });
  }

  // View switching
  document.getElementById('calendar-view-btn').addEventListener('click', () => {
    document.getElementById('list-view-btn').classList.remove('active');
    document.getElementById('calendar-view-btn').classList.add('active');
    document.getElementById('calendar-container').style.display = 'block';
    document.getElementById('list-container').style.display = 'none';
    renderCalendarView();
  });

  document.getElementById('list-view-btn').addEventListener('click', () => {
    document.getElementById('calendar-view-btn').classList.remove('active');
    document.getElementById('list-view-btn').classList.add('active');
    document.getElementById('calendar-container').style.display = 'none';
    document.getElementById('list-container').style.display = 'block';
  });

  document.getElementById('calendar-list-view-btn').addEventListener('click', () => {
    document.getElementById('list-view-btn').click();
  });

  // Modal and button handlers
  document.getElementById('edit-employees-btn').addEventListener('click', () =>
    new bootstrap.Modal(document.getElementById('employeeModal')).show()
  );
  document.getElementById('add-emp-btn').addEventListener('click', addEmployee);
  document.getElementById('add-entry-btn').addEventListener('click', addEntry);

  // Filter handlers
  document.getElementById('employee-filter').addEventListener('change', applyFilters);
  document.getElementById('start-date-filter').addEventListener('change', applyFilters);
  document.getElementById('calendar-employee-filter').addEventListener('change', renderCalendarView);

  // Initialize
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
